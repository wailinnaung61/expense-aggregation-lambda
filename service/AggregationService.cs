using Amazon.DynamoDBv2;
using Amazon.DynamoDBv2.Model;
using Amazon.Lambda.Core;
using Amazon.Lambda.DynamoDBEvents;
using System.Globalization;

namespace ExpenseAggregationLambda.service;

public class AggregationService
{
    private readonly IAmazonDynamoDB _ddb = new AmazonDynamoDBClient();
    private const string AGG_TABLE = "Aggregations";

    public async Task HandleRecordAsync(DynamoDBEvent.DynamodbStreamRecord record, ILambdaContext context)
    {
        context.Logger.LogInformation($"[HandleRecord] Event: {record.EventName}");

        if (record.EventName == "INSERT")
        {
            context.Logger.LogInformation("[INSERT] Processing INSERT event");
            var item = record.Dynamodb.NewImage.ToLedgerItem();

            context.Logger.LogInformation($"[INSERT] Transaction: {item.TransactionId}, Type: {item.Type}, Amount: {item.Amount}, Status: {item.PaymentStatus}");

            if (item.PaymentStatus == "COMPLETED")
            {
                context.Logger.LogInformation("[INSERT] Payment status is COMPLETED - applying aggregation (+1)");
                await ApplyAsync(item, +1, context);
            }
            else
            {
                context.Logger.LogInformation($"[INSERT] Payment status is {item.PaymentStatus} - skipping aggregation");
            }
        }
        else if (record.EventName == "REMOVE")
        {
            context.Logger.LogInformation("[REMOVE] Processing REMOVE event");
            var item = record.Dynamodb.OldImage.ToLedgerItem();

            context.Logger.LogInformation($"[REMOVE] Transaction: {item.TransactionId}, Type: {item.Type}, Amount: {item.Amount}, Status: {item.PaymentStatus}");

            if (item.PaymentStatus == "COMPLETED")
            {
                context.Logger.LogInformation("[REMOVE] Payment status is COMPLETED - reversing aggregation (-1)");
                await ApplyAsync(item, -1, context);
            }
            else
            {
                context.Logger.LogInformation($"[REMOVE] Payment status is {item.PaymentStatus} - skipping aggregation");
            }
        }
        else if (record.EventName == "MODIFY")
        {
            context.Logger.LogInformation("[MODIFY] Processing MODIFY event");
            var oldItem = record.Dynamodb.OldImage.ToLedgerItem();
            var newItem = record.Dynamodb.NewImage.ToLedgerItem();

            context.Logger.LogInformation($"[MODIFY] Transaction: {newItem.TransactionId}");
            context.Logger.LogInformation($"[MODIFY] Old status: {oldItem.PaymentStatus}, New status: {newItem.PaymentStatus}");

            if (oldItem.PaymentStatus == "COMPLETED")
            {
                context.Logger.LogInformation("[MODIFY] Old item was COMPLETED - reversing aggregation (-1)");
                await ApplyAsync(oldItem, -1, context);
            }

            if (newItem.PaymentStatus == "COMPLETED")
            {
                context.Logger.LogInformation("[MODIFY] New item is COMPLETED - applying aggregation (+1)");
                await ApplyAsync(newItem, +1, context);
            }

            if (oldItem.PaymentStatus != "COMPLETED" && newItem.PaymentStatus != "COMPLETED")
            {
                context.Logger.LogInformation("[MODIFY] Neither old nor new status is COMPLETED - no aggregation needed");
            }
        }
        else
        {
            context.Logger.LogWarning($"[UNKNOWN] Unknown event type: {record.EventName}");
        }
    }

    private async Task ApplyAsync(LedgerItem item, int direction, ILambdaContext context)
    {
        context.Logger.LogInformation($"[ApplyAsync] Starting aggregation for User: {item.UserId}, Type: {item.Type}, Amount: {item.Amount}, Direction: {direction}");

        var dayKey = item.Date.ToString("yyyy-MM-dd");
        var weekKey = ISOWeek.GetWeekOfYear(item.Date);
        var year = item.Date.Year;

        var monthKey = GetSalaryMonth(item.Date);

        context.Logger.LogInformation($"[ApplyAsync] Aggregation keys - Day: {dayKey}, Week: W{weekKey}, Month: {monthKey}, Year: {year}");

        context.Logger.LogInformation("[ApplyAsync] Updating DAY aggregation...");
        await UpdateAggAsync(item, $"AGG#DAY#{dayKey}", direction, context);

        context.Logger.LogInformation("[ApplyAsync] Updating WEEK aggregation...");
        await UpdateAggAsync(item, $"AGG#WEEK#{year}-W{weekKey:D2}", direction, context);

        context.Logger.LogInformation("[ApplyAsync] Updating MONTH aggregation...");
        await UpdateAggAsync(item, $"AGG#MONTH#{monthKey}", direction, context);

        context.Logger.LogInformation("[ApplyAsync] Updating YEAR aggregation...");
        await UpdateAggAsync(item, $"AGG#YEAR#{year}", direction, context);

        context.Logger.LogInformation("[ApplyAsync] Updating CATEGORY aggregation...");
        await UpdateCategoryAggAsync(item, monthKey, direction, context);

        context.Logger.LogInformation("[ApplyAsync] All aggregations completed successfully");
    }

    private async Task UpdateAggAsync(LedgerItem item, string sortKey, int direction, ILambdaContext context)
    {
        var pk = $"USER#{item.UserId}";
        var sk = sortKey;

        context.Logger.LogInformation($"[UpdateAggAsync] Updating {AGG_TABLE} - PK: {pk}, SK: {sk}");

        var updateExpression = new List<string>();
        var expressionAttributeNames = new Dictionary<string, string>();
        var expressionAttributeValues = new Dictionary<string, AttributeValue>();

        var amount = item.Amount * direction;

        if (item.Type == "INCOME")
        {
            updateExpression.Add("#income = if_not_exists(#income, :zero) + :amount");
            expressionAttributeNames["#income"] = "income";
            context.Logger.LogInformation($"[UpdateAggAsync] Type: INCOME, Amount change: {amount}");
        }
        else if (item.Type == "EXPENSE")
        {
            updateExpression.Add("#expense = if_not_exists(#expense, :zero) + :amount");
            expressionAttributeNames["#expense"] = "expense";
            context.Logger.LogInformation($"[UpdateAggAsync] Type: EXPENSE, Amount change: {amount}");
        }
        else if (item.Type == "SAVING")
        {
            updateExpression.Add("#saving = if_not_exists(#saving, :zero) + :amount");
            expressionAttributeNames["#saving"] = "saving";
            context.Logger.LogInformation($"[UpdateAggAsync] Type: SAVING, Amount change: {amount}");
        }
        else if (item.Type == "INVESTMENT")
        {
            updateExpression.Add("#investment = if_not_exists(#investment, :zero) + :amount");
            expressionAttributeNames["#investment"] = "investment";
            context.Logger.LogInformation($"[UpdateAggAsync] Type: INVESTMENT, Amount change: {amount}");
        }

        updateExpression.Add("#tranactionCount = if_not_exists(#tranactionCount, :zero) + :direction");
        expressionAttributeNames["#tranactionCount"] = "tranactionCount";

        expressionAttributeValues[":amount"] = new AttributeValue { N = amount.ToString() };
        expressionAttributeValues[":zero"] = new AttributeValue { N = "0" };
        expressionAttributeValues[":direction"] = new AttributeValue { N = direction.ToString() };

        if (sk.StartsWith("AGG#MONTH#"))
        {
            var (periodStart, periodEnd) = GetSalaryMonthPeriod(item.Date);
            updateExpression.Add("#periodStart = if_not_exists(#periodStart, :periodStart)");
            updateExpression.Add("#periodEnd = if_not_exists(#periodEnd, :periodEnd)");
            updateExpression.Add("#period = if_not_exists(#period, :period)");

            expressionAttributeNames["#periodStart"] = "periodStart";
            expressionAttributeNames["#periodEnd"] = "periodEnd";
            expressionAttributeNames["#period"] = "period";

            expressionAttributeValues[":periodStart"] = new AttributeValue { S = periodStart.ToString("yyyy/MM/dd") };
            expressionAttributeValues[":periodEnd"] = new AttributeValue { S = periodEnd.ToString("yyyy/MM/dd") };
            expressionAttributeValues[":period"] = new AttributeValue { S = sk.Replace("AGG#MONTH#", "") };
        }
        else if (sk.StartsWith("AGG#DAY#") || sk.StartsWith("AGG#WEEK#") || sk.StartsWith("AGG#YEAR#"))
        {
            updateExpression.Add("#period = if_not_exists(#period, :period)");
            expressionAttributeNames["#period"] = "period";

            string periodValue = sk.Replace("AGG#DAY#", "").Replace("AGG#WEEK#", "").Replace("AGG#YEAR#", "");
            expressionAttributeValues[":period"] = new AttributeValue { S = periodValue };
        }

        var request = new UpdateItemRequest
        {
            TableName = AGG_TABLE,
            Key = new Dictionary<string, AttributeValue>
            {
                ["PK"] = new AttributeValue { S = pk },
                ["SK"] = new AttributeValue { S = sk }
            },
            UpdateExpression = "SET " + string.Join(", ", updateExpression),
            ExpressionAttributeNames = expressionAttributeNames,
            ExpressionAttributeValues = expressionAttributeValues
        };

        context.Logger.LogInformation($"[UpdateAggAsync] Executing DynamoDB UpdateItem...");

        try
        {
            await _ddb.UpdateItemAsync(request);
            context.Logger.LogInformation($"[UpdateAggAsync] ✅ Successfully updated {sk}");
        }
        catch (Exception ex)
        {
            context.Logger.LogError($"[UpdateAggAsync] ❌ Failed to update {sk}: {ex.Message}");
            throw;
        }
    }

    private async Task UpdateCategoryAggAsync(LedgerItem item, string monthKey, int direction, ILambdaContext context)
    {
        var pk = $"USER#{item.UserId}";
        var sk = $"CAT#{item.Type}#{item.CategoryId}#{monthKey}";

        context.Logger.LogInformation($"[UpdateCategoryAgg] Updating category aggregation - PK: {pk}, SK: {sk}");
        context.Logger.LogInformation($"[UpdateCategoryAgg] Category: {item.CategoryId}, Type: {item.Type}, Month: {monthKey}");

        var amount = item.Amount * direction;

        var (periodStart, periodEnd) = GetSalaryMonthPeriod(item.Date);

        var updateExpression = new List<string>
        {
            "#totalAmount = if_not_exists(#totalAmount, :zero) + :amount",
            "#tranactionCount = if_not_exists(#tranactionCount, :zero) + :direction",
            "#categoryId = if_not_exists(#categoryId, :categoryId)",
            "#period = if_not_exists(#period, :period)",
            "#periodStart = if_not_exists(#periodStart, :periodStart)",
            "#periodEnd = if_not_exists(#periodEnd, :periodEnd)"
        };

        var request = new UpdateItemRequest
        {
            TableName = AGG_TABLE,
            Key = new Dictionary<string, AttributeValue>
            {
                ["PK"] = new AttributeValue { S = pk },
                ["SK"] = new AttributeValue { S = sk }
            },
            UpdateExpression = "SET " + string.Join(", ", updateExpression),
            ExpressionAttributeNames = new Dictionary<string, string>
            {
                ["#totalAmount"] = "totalAmount",
                ["#tranactionCount"] = "tranactionCount",
                ["#categoryId"] = "categoryId",
                ["#period"] = "period",
                ["#periodStart"] = "periodStart",
                ["#periodEnd"] = "periodEnd"
            },
            ExpressionAttributeValues = new Dictionary<string, AttributeValue>
            {
                [":amount"] = new AttributeValue { N = amount.ToString() },
                [":zero"] = new AttributeValue { N = "0" },
                [":direction"] = new AttributeValue { N = direction.ToString() },
                [":categoryId"] = new AttributeValue { S = item.CategoryId },
                [":period"] = new AttributeValue { S = monthKey },
                [":periodStart"] = new AttributeValue { S = periodStart.ToString("yyyy/MM/dd") },
                [":periodEnd"] = new AttributeValue { S = periodEnd.ToString("yyyy/MM/dd") }
            }
        };

        context.Logger.LogInformation($"[UpdateCategoryAgg] Amount change: {amount}, Direction: {direction}");

        try
        {
            await _ddb.UpdateItemAsync(request);
            context.Logger.LogInformation($"[UpdateCategoryAgg] ✅ Successfully updated category aggregation");
        }
        catch (Exception ex)
        {
            context.Logger.LogError($"[UpdateCategoryAgg] ❌ Failed to update category aggregation: {ex.Message}");
            throw;
        }
    }

    private string GetSalaryMonth(DateTime date)
    {
        return $"{date.Year}-{date.Month:D2}";
    }

    private (DateTime periodStart, DateTime periodEnd) GetSalaryMonthPeriod(DateTime date)
    {
        var periodStart = new DateTime(date.Year, date.Month, 1);
        var periodEnd = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

        return (periodStart, periodEnd);
    }
}

public class LedgerItem
{
    public string UserId { get; set; } = string.Empty;
    public string TransactionId { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string CategoryId { get; set; } = string.Empty;
    public DateTime Date { get; set; }
    public string PaymentStatus { get; set; } = string.Empty;
}

public static class DynamoDBExtensions
{
    public static LedgerItem ToLedgerItem(this Dictionary<string, DynamoDBEvent.AttributeValue> image)
    {
        return new LedgerItem
        {
            UserId = image.ContainsKey("userId") ? image["userId"].S : string.Empty,
            TransactionId = image.ContainsKey("tranactionId") ? image["tranactionId"].S : string.Empty,
            Type = image.ContainsKey("type") ? image["type"].S : string.Empty,
            Amount = image.ContainsKey("amount") && decimal.TryParse(image["amount"].N, out var amount) ? amount : 0,
            CategoryId = image.ContainsKey("categoryId") ? image["categoryId"].S : string.Empty,
            Date = image.ContainsKey("date") && DateTime.TryParse(image["date"].S, out var date) ? date : DateTime.UtcNow,
            PaymentStatus = image.ContainsKey("paymentStatus") ? image["paymentStatus"].S : string.Empty
        };
    }
}