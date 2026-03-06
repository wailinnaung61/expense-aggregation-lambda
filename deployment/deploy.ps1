# AWS Lambda Deployment Script - Build & Deploy Only
# Prerequisites: IAM Role and DynamoDB Stream must be configured manually

Write-Host "🚀 Starting Lambda Deployment..." -ForegroundColor Cyan

# Configuration
$FUNCTION_NAME = "ExpenseAggregationLambda"
$REGION = "us-east-1"

# Step 1: Build and Package Lambda
Write-Host "`n📦 Step 1: Building Lambda package..." -ForegroundColor Yellow

dotnet lambda package -c Release -o bin/Release/deployment-package.zip

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Build failed!" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Package created: bin/Release/deployment-package.zip" -ForegroundColor Green

# Step 2: Deploy Lambda Function
Write-Host "`n🚀 Step 2: Deploying Lambda function..." -ForegroundColor Yellow

$functionExists = aws lambda get-function --function-name $FUNCTION_NAME --region $REGION 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Function does not exist. Please create it manually in AWS Console first." -ForegroundColor Red
    Write-Host "   Or use the full deployment script: .\deployment\deploy-full.ps1" -ForegroundColor Yellow
    exit 1
} else {
    # Update existing function
    Write-Host "Updating Lambda function code: $FUNCTION_NAME"

    aws lambda update-function-code `
        --function-name $FUNCTION_NAME `
        --zip-file fileb://bin/Release/deployment-package.zip `
        --region $REGION

    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Function code updated successfully!" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to update function code!" -ForegroundColor Red
        exit 1
    }
}

# Summary
Write-Host "`n✨ Deployment Complete!" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "Function Name: $FUNCTION_NAME" -ForegroundColor White 
Write-Host "Region: $REGION" -ForegroundColor White
Write-Host "Package: bin/Release/deployment-package.zip" -ForegroundColor White
Write-Host "`nNext Steps:" -ForegroundColor Yellow
Write-Host "1. Test in AWS Console: Lambda → $FUNCTION_NAME → Test tab" -ForegroundColor Gray
Write-Host "2. View logs: aws logs tail /aws/lambda/$FUNCTION_NAME --follow --region $REGION" -ForegroundColor Gray
Write-Host "3. Insert test data: aws dynamodb put-item --table-name FinanceLedger ..." -ForegroundColor Gray
