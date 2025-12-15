# ============================================================================
# ChargeScrub Deployment Script
# ============================================================================
# 
# PREREQUISITES:
# 1. AWS CLI configured with appropriate credentials
# 2. Python 3.12 in PATH (for building openpyxl layer)
# 3. Route 53 hosted zone for your domain
#
# USAGE:
# 1. Copy deploy-config.example.ps1 to deploy-config.ps1
# 2. Fill in your AWS account details
# 3. Run: .\deploy.ps1
#
# ============================================================================

# Load configuration
if (Test-Path ".\deploy-config.ps1") {
    . .\deploy-config.ps1
} else {
    Write-Host "ERROR: deploy-config.ps1 not found!" -ForegroundColor Red
    Write-Host "Copy deploy-config.example.ps1 to deploy-config.ps1 and fill in your values." -ForegroundColor Yellow
    exit 1
}

# Validate required variables
$requiredVars = @("Domain", "Region", "AccountId", "HostedZoneId")
foreach ($var in $requiredVars) {
    if (-not (Get-Variable -Name $var -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: Required variable '$var' not set in deploy-config.ps1" -ForegroundColor Red
        exit 1
    }
}

# Derived names
$ProjectPrefix = "ChargeScrub"
$BucketName = "chargescrub-static-$AccountId"
$RoleName = "$ProjectPrefix-Lambda-Role"
$LambdaName = "$ProjectPrefix-Processor"
$RoleArn = "arn:aws:iam::${AccountId}:role/$RoleName"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "ChargeScrub Deployment Script" -ForegroundColor Cyan
Write-Host "Domain: $Domain" -ForegroundColor Cyan
Write-Host "Region: $Region" -ForegroundColor Cyan
Write-Host "Account: $AccountId" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# ============================================================================
# STEP 1: Create IAM Role
# ============================================================================
Write-Host "`n[Step 1/10] Creating IAM Role for Lambda..." -ForegroundColor Yellow

aws iam create-role `
    --role-name $RoleName `
    --assume-role-policy-document '{\"Version\":\"2012-10-17\",\"Statement\":[{\"Effect\":\"Allow\",\"Principal\":{\"Service\":\"lambda.amazonaws.com\"},\"Action\":\"sts:AssumeRole\"}]}' `
    --description "Execution role for ChargeScrub Lambda function" 2>$null

aws iam attach-role-policy `
    --role-name $RoleName `
    --policy-arn "arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole"

Write-Host "IAM Role created: $RoleName" -ForegroundColor Green
Write-Host "Waiting 15 seconds for IAM role propagation..." -ForegroundColor Gray
Start-Sleep -Seconds 15

# ============================================================================
# STEP 2: Create S3 Bucket
# ============================================================================
Write-Host "`n[Step 2/10] Creating S3 Bucket..." -ForegroundColor Yellow

aws s3api create-bucket `
    --bucket $BucketName `
    --region $Region `
    --create-bucket-configuration LocationConstraint=$Region 2>$null

aws s3api put-public-access-block `
    --bucket $BucketName `
    --public-access-block-configuration "BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=true,RestrictPublicBuckets=true"

Write-Host "S3 Bucket created: $BucketName" -ForegroundColor Green

# ============================================================================
# STEP 3: Upload Static Site Files
# ============================================================================
Write-Host "`n[Step 3/10] Uploading Static Site Files..." -ForegroundColor Yellow

aws s3 sync .\frontend\ s3://$BucketName/ --region $Region
Write-Host "Static files uploaded to S3" -ForegroundColor Green

# ============================================================================
# STEP 4: Setup Lambda Layers
# ============================================================================
Write-Host "`n[Step 4/10] Setting up Lambda Layers..." -ForegroundColor Yellow

$PandasLayerArn = "arn:aws:lambda:${Region}:336392948345:layer:AWSSDKPandas-Python312:15"
Write-Host "Using AWS public pandas layer: $PandasLayerArn" -ForegroundColor Green

Write-Host "Creating openpyxl layer..." -ForegroundColor Gray
New-Item -ItemType Directory -Force -Path ".\layer\python" | Out-Null
pip install openpyxl et-xmlfile -t .\layer\python --quiet --upgrade --no-deps

Set-Location .\layer
Compress-Archive -Path python -DestinationPath ..\openpyxl-layer.zip -Force
Set-Location ..

$OpenpyxlLayerArn = aws lambda publish-layer-version `
    --layer-name "$ProjectPrefix-Openpyxl" `
    --description "openpyxl for ChargeScrub Excel processing" `
    --zip-file fileb://openpyxl-layer.zip `
    --compatible-runtimes python3.12 python3.11 `
    --region $Region `
    --query 'LayerVersionArn' `
    --output text

Write-Host "Openpyxl layer created: $OpenpyxlLayerArn" -ForegroundColor Green

# ============================================================================
# STEP 5: Create Lambda Function
# ============================================================================
Write-Host "`n[Step 5/10] Creating Lambda Function..." -ForegroundColor Yellow

Compress-Archive -Path .\lambda\lambda_function.py -DestinationPath lambda-code.zip -Force

aws lambda create-function `
    --function-name $LambdaName `
    --runtime python3.12 `
    --role $RoleArn `
    --handler lambda_function.lambda_handler `
    --zip-file fileb://lambda-code.zip `
    --timeout 60 `
    --memory-size 512 `
    --layers $PandasLayerArn $OpenpyxlLayerArn `
    --environment "Variables={ALLOWED_ORIGIN=https://$Domain}" `
    --region $Region `
    --description "ChargeScrub charge processing function"

Write-Host "Lambda function created: $LambdaName" -ForegroundColor Green

# ============================================================================
# STEP 6: Create API Gateway
# ============================================================================
Write-Host "`n[Step 6/10] Creating API Gateway..." -ForegroundColor Yellow

$ApiResponse = aws apigatewayv2 create-api `
    --name "$ProjectPrefix-API" `
    --protocol-type HTTP `
    --cors-configuration "AllowOrigins=https://$Domain,AllowMethods=POST,GET,OPTIONS,AllowHeaders=Content-Type" `
    --region $Region `
    --output json | ConvertFrom-Json

$ApiId = $ApiResponse.ApiId
$ApiEndpoint = $ApiResponse.ApiEndpoint

$IntegrationId = aws apigatewayv2 create-integration `
    --api-id $ApiId `
    --integration-type AWS_PROXY `
    --integration-uri "arn:aws:lambda:${Region}:${AccountId}:function:$LambdaName" `
    --payload-format-version "2.0" `
    --region $Region `
    --query 'IntegrationId' `
    --output text

aws apigatewayv2 create-route --api-id $ApiId --route-key "POST /api/upload" --target "integrations/$IntegrationId" --region $Region | Out-Null
aws apigatewayv2 create-route --api-id $ApiId --route-key "GET /api/health" --target "integrations/$IntegrationId" --region $Region | Out-Null

aws apigatewayv2 create-stage --api-id $ApiId --stage-name "`$default" --auto-deploy --region $Region | Out-Null

aws lambda add-permission `
    --function-name $LambdaName `
    --statement-id "AllowAPIGateway" `
    --action "lambda:InvokeFunction" `
    --principal "apigateway.amazonaws.com" `
    --source-arn "arn:aws:execute-api:${Region}:${AccountId}:${ApiId}/*" `
    --region $Region | Out-Null

Write-Host "API Gateway created: $ApiId" -ForegroundColor Green
Write-Host "API Endpoint: $ApiEndpoint" -ForegroundColor Green

# ============================================================================
# STEP 7: Request ACM Certificate
# ============================================================================
Write-Host "`n[Step 7/10] Requesting ACM Certificate..." -ForegroundColor Yellow
Write-Host "Certificate must be in us-east-1 for CloudFront" -ForegroundColor Gray

$CertArn = aws acm request-certificate `
    --domain-name $Domain `
    --validation-method DNS `
    --region us-east-1 `
    --query 'CertificateArn' `
    --output text

Write-Host "Certificate requested: $CertArn" -ForegroundColor Green

Start-Sleep -Seconds 5

$CertDetails = aws acm describe-certificate --certificate-arn $CertArn --region us-east-1 --output json | ConvertFrom-Json
$DnsRecord = $CertDetails.Certificate.DomainValidationOptions[0].ResourceRecord

Write-Host "Adding DNS validation record to Route 53..." -ForegroundColor Gray

$ChangeBatch = @{
    Changes = @(@{
        Action = "UPSERT"
        ResourceRecordSet = @{
            Name = $DnsRecord.Name
            Type = $DnsRecord.Type
            TTL = 300
            ResourceRecords = @(@{Value = $DnsRecord.Value})
        }
    })
} | ConvertTo-Json -Depth 10 -Compress

[System.IO.File]::WriteAllText("$PWD\dns-validation.json", $ChangeBatch)
aws route53 change-resource-record-sets --hosted-zone-id $HostedZoneId --change-batch file://dns-validation.json | Out-Null

Write-Host "Waiting for certificate validation (this may take 5-10 minutes)..." -ForegroundColor Gray

do {
    Start-Sleep -Seconds 30
    $status = aws acm describe-certificate --certificate-arn $CertArn --region us-east-1 --query 'Certificate.Status' --output text
    Write-Host "  Certificate status: $status" -ForegroundColor Gray
} while ($status -eq "PENDING_VALIDATION")

Write-Host "Certificate validated!" -ForegroundColor Green

# ============================================================================
# STEP 8: Create CloudFront OAC
# ============================================================================
Write-Host "`n[Step 8/10] Creating CloudFront Distribution..." -ForegroundColor Yellow

$OacResponse = aws cloudfront create-origin-access-control `
    --origin-access-control-config "Name=$ProjectPrefix-OAC,Description=OAC for ChargeScrub S3,SigningProtocol=sigv4,SigningBehavior=always,OriginAccessControlOriginType=s3" `
    --output json | ConvertFrom-Json

$OacId = $OacResponse.OriginAccessControl.Id
Write-Host "OAC Created: $OacId" -ForegroundColor Green

# ============================================================================
# STEP 9: Create CloudFront Distribution
# ============================================================================
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$cfConfig = @"
{
    "CallerReference": "$ProjectPrefix-$timestamp",
    "Aliases": {"Quantity": 1, "Items": ["$Domain"]},
    "DefaultRootObject": "index.html",
    "Origins": {
        "Quantity": 2,
        "Items": [
            {
                "Id": "S3-chargescrub",
                "DomainName": "$BucketName.s3.$Region.amazonaws.com",
                "OriginAccessControlId": "$OacId",
                "S3OriginConfig": {"OriginAccessIdentity": ""}
            },
            {
                "Id": "API-Gateway",
                "DomainName": "$ApiId.execute-api.$Region.amazonaws.com",
                "CustomOriginConfig": {
                    "HTTPPort": 80,
                    "HTTPSPort": 443,
                    "OriginProtocolPolicy": "https-only",
                    "OriginSslProtocols": {"Quantity": 1, "Items": ["TLSv1.2"]}
                }
            }
        ]
    },
    "DefaultCacheBehavior": {
        "TargetOriginId": "S3-chargescrub",
        "ViewerProtocolPolicy": "redirect-to-https",
        "AllowedMethods": {"Quantity": 2, "Items": ["GET", "HEAD"], "CachedMethods": {"Quantity": 2, "Items": ["GET", "HEAD"]}},
        "CachePolicyId": "658327ea-f89d-4fab-a63d-7e88639e58f6",
        "Compress": true
    },
    "CacheBehaviors": {
        "Quantity": 1,
        "Items": [{
            "PathPattern": "/api/*",
            "TargetOriginId": "API-Gateway",
            "ViewerProtocolPolicy": "https-only",
            "AllowedMethods": {"Quantity": 7, "Items": ["GET", "HEAD", "OPTIONS", "PUT", "POST", "PATCH", "DELETE"], "CachedMethods": {"Quantity": 2, "Items": ["GET", "HEAD"]}},
            "CachePolicyId": "4135ea2d-6df8-44a3-9df3-4b5a84be39ad",
            "OriginRequestPolicyId": "b689b0a8-53d0-40ab-baf2-68738e2966ac",
            "Compress": true
        }]
    },
    "Comment": "ChargeScrub CDN",
    "Enabled": true,
    "ViewerCertificate": {
        "ACMCertificateArn": "$CertArn",
        "SSLSupportMethod": "sni-only",
        "MinimumProtocolVersion": "TLSv1.2_2021"
    },
    "HttpVersion": "http2and3"
}
"@

[System.IO.File]::WriteAllText("$PWD\cf-config.json", $cfConfig)

$DistResponse = aws cloudfront create-distribution --distribution-config file://cf-config.json --output json | ConvertFrom-Json
$DistributionId = $DistResponse.Distribution.Id
$DistributionDomain = $DistResponse.Distribution.DomainName

Write-Host "CloudFront Distribution created: $DistributionId" -ForegroundColor Green
Write-Host "CloudFront Domain: $DistributionDomain" -ForegroundColor Green

# ============================================================================
# STEP 10: Update S3 Bucket Policy & Create Route 53 Record
# ============================================================================
Write-Host "`n[Step 10/10] Finalizing Configuration..." -ForegroundColor Yellow

$bucketPolicy = @"
{
    "Version": "2012-10-17",
    "Statement": [{
        "Sid": "AllowCloudFrontServicePrincipal",
        "Effect": "Allow",
        "Principal": {"Service": "cloudfront.amazonaws.com"},
        "Action": "s3:GetObject",
        "Resource": "arn:aws:s3:::$BucketName/*",
        "Condition": {
            "StringEquals": {
                "AWS:SourceArn": "arn:aws:cloudfront::${AccountId}:distribution/$DistributionId"
            }
        }
    }]
}
"@

[System.IO.File]::WriteAllText("$PWD\bucket-policy.json", $bucketPolicy)
aws s3api put-bucket-policy --bucket $BucketName --policy file://bucket-policy.json

$route53Change = @{
    Changes = @(@{
        Action = "UPSERT"
        ResourceRecordSet = @{
            Name = $Domain
            Type = "A"
            AliasTarget = @{
                HostedZoneId = "Z2FDTNDATAQYW2"
                DNSName = $DistributionDomain
                EvaluateTargetHealth = $false
            }
        }
    })
} | ConvertTo-Json -Depth 10 -Compress

[System.IO.File]::WriteAllText("$PWD\route53-alias.json", $route53Change)
aws route53 change-resource-record-sets --hosted-zone-id $HostedZoneId --change-batch file://route53-alias.json | Out-Null

Write-Host "Route 53 A record created!" -ForegroundColor Green

# ============================================================================
# DEPLOYMENT COMPLETE
# ============================================================================
Write-Host "`n============================================" -ForegroundColor Green
Write-Host "DEPLOYMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

Write-Host "`nYour site will be available at:" -ForegroundColor White
Write-Host "  https://$Domain" -ForegroundColor Cyan

Write-Host "`nNote: CloudFront distribution may take 10-15 minutes to fully deploy." -ForegroundColor Yellow
Write-Host "Check status with:" -ForegroundColor Gray
Write-Host "  aws cloudfront get-distribution --id $DistributionId --query 'Distribution.Status'" -ForegroundColor Gray

# Save deployment info (keep private!)
$DeploymentInfo = @"

# ChargeScrub Deployment Info
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# 
# KEEP THIS FILE PRIVATE - Contains account-specific resource IDs

Domain=$Domain
Region=$Region
AccountId=$AccountId
HostedZoneId=$HostedZoneId
BucketName=$BucketName
LambdaName=$LambdaName
ApiId=$ApiId
ApiEndpoint=$ApiEndpoint
DistributionId=$DistributionId
DistributionDomain=$DistributionDomain
CertArn=$CertArn
OacId=$OacId
"@

$DeploymentInfo | Out-File -FilePath ".\deployment-info.txt" -Encoding UTF8
Write-Host "`nDeployment info saved to deployment-info.txt (keep private!)" -ForegroundColor Yellow
