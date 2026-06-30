# AGENTS.md

## Cursor Cloud specific instructions

### What this is
`ExpenseAggregationLambda` is a **.NET 10 AWS Lambda** (C#, class library — no standalone entry point). It is triggered by a **DynamoDB stream** (`DynamoDBEvent`) and writes rolling aggregations (day / week / month / year / category) into a DynamoDB table named `Aggregations`. Core logic lives in `service/AggregationService.cs`; the Lambda entry point is `Function.FunctionHandler` in `Function.cs`.

### Toolchain
- The .NET 10 SDK is installed at `~/.dotnet` and added to `PATH` via `~/.bashrc` (`DOTNET_ROOT`, `PATH`, telemetry opt-out). New non-login shells inherit it; if `dotnet` is not found, run `source ~/.bashrc` or use `~/.dotnet/dotnet`.
- The update script installs the SDK (idempotent) and runs `dotnet restore`.

### Build / lint
- Build (also serves as the lint/compile check — there is no separate linter): `dotnet build` (or `dotnet build -c Release`).
- Package for deployment: `dotnet lambda package` (requires `dotnet tool install -g Amazon.Lambda.Tools`).

### Tests
- There is **no automated test project**. `xunit` is referenced in `ExpenseAggregationLambda.csproj`, but there are no test files and no test runner (`Microsoft.NET.Test.Sdk` / `xunit.runner.visualstudio`) is referenced, so `dotnet test` discovers nothing.

### Running / verifying end-to-end (no AWS account needed)
The handler constructs `new AmazonDynamoDBClient()` with **no explicit endpoint**, so to run locally you point the AWS SDK at a local DynamoDB via environment variables, then invoke the handler with one of the sample events in `test-events/`.

1. Start **DynamoDB Local** (Java 21 is preinstalled; no Docker available). Download once and run:
   ```bash
   curl -fsSL https://d1ni2b6xgvw0s0.cloudfront.net/v2.x/dynamodb_local_latest.tar.gz -o /tmp/ddb.tar.gz
   mkdir -p /tmp/dynamodb-local && tar -xzf /tmp/ddb.tar.gz -C /tmp/dynamodb-local
   java -Djava.library.path=/tmp/dynamodb-local/DynamoDBLocal_lib \
        -jar /tmp/dynamodb-local/DynamoDBLocal.jar -sharedDb -port 8000   # runs in foreground
   ```
2. Point the SDK at it (the AWSSDK v4 service-specific endpoint env var is what makes the Lambda's no-arg client connect locally):
   ```bash
   export AWS_REGION=us-east-1 AWS_DEFAULT_REGION=us-east-1
   export AWS_ACCESS_KEY_ID=fake AWS_SECRET_ACCESS_KEY=fake
   export AWS_ENDPOINT_URL_DYNAMODB=http://localhost:8000
   ```
3. Create the `Aggregations` table (key schema: `PK` string HASH, `SK` string RANGE) and invoke `Function.FunctionHandler` with a deserialized `test-events/*.json` event (use `DefaultLambdaJsonSerializer` to deserialize into `DynamoDBEvent`, and `Amazon.Lambda.TestUtilities.TestLambdaContext` for the context). A small throwaway console project referencing `ExpenseAggregationLambda.csproj` is the easiest harness.
- Alternatively use the AWS Lambda Test Tool: `dotnet tool install -g Amazon.Lambda.TestTool` (note: `Properties/launchSettings.json` references a Windows path to `dotnet-lambda-test-tool-10.0.exe` that does not exist on this Linux VM).

### Gotchas
- `test-events/update-pending-to-complete.json` uses `paymentStatus` value `"COMPLETE"` (not `"COMPLETED"`), so the handler intentionally performs **no aggregation** for it — this is sample-data behavior, not a bug.
- The DynamoDB attribute for the transaction id is spelled `tranactionId` (and the code field `tranactionCount`) — preserve the existing (mis)spelling to stay compatible with the data model.
