using Amazon.Lambda.Core;
using Amazon.Lambda.DynamoDBEvents;
using ExpenseAggregationLambda.service;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace ExpenseAggregationLambda;
    public class Function
    {
        private readonly AggregationService _service = new();

        public async Task FunctionHandler(DynamoDBEvent dynamoEvent, ILambdaContext context)
        {
            context.Logger.LogInformation($"[START] Processing {dynamoEvent.Records.Count} DynamoDB records");

            int processed = 0;
            int skipped = 0;
            int errors = 0;

            foreach (var record in dynamoEvent.Records)
            {
                try
                {
                    context.Logger.LogInformation($"[RECORD {processed + 1}] EventName: {record.EventName}, EventID: {record.EventID}");

                    await _service.HandleRecordAsync(record, context);
                    processed++;

                    context.Logger.LogInformation($"[SUCCESS] Record {processed} processed successfully");
                }
                catch (Exception ex)
                {
                    errors++;
                    context.Logger.LogError($"[ERROR] Failed to process record {processed + 1}: {ex.Message}");
                    context.Logger.LogError($"[ERROR] Stack trace: {ex.StackTrace}");
                }
            }

            context.Logger.LogInformation($"[COMPLETE] Processed: {processed}, Skipped: {skipped}, Errors: {errors}");
        }
    }
