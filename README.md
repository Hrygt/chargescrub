# ChargeScrub

**Hospital Billing E/M Charge Duplicate & Overlap Review Tool**

ChargeScrub automates the review of hospitalist E/M charges, instantly identifying duplicate codes, overlapping visits, and charges requiring timeline verification. Reduce billing errors and audit risk while saving hours of manual review.

## Features

- 🔍 **Automatic Analysis** — Upload Excel charge reports, get instant recommendations
- ✅ **Color-Coded Output** — Green (Charge), Red (Deny), Blue (Review)
- 🔒 **HIPAA-Aligned Security** — Zero storage, no PHI logging, in-memory processing only
- ☁️ **Serverless Architecture** — Low cost (~$0.50/month), auto-scaling
- 📊 **Summary Statistics** — Quick overview of charge distribution

## Billing Rules (v1.3)

| Rule | Condition | Result |
|------|-----------|--------|
| 1 | 99499 (Unlisted E/M) | **Deny** |
| 2 | 99497 (ACP) different provider than admit | **Charge** |
| 2b | 99497 (ACP) same provider as admit | **Deny** |
| 3 | Critical + Subsequent + Discharge (no Initial) | Critical: **Charge**, Others: **Deny** |
| 4 | Critical + Subsequent only (no Initial) | Both: **Review** (timeline needed) |
| 5 | Initial + Critical both present | Both: **Review** (timeline needed) |
| 6 | Initial present, no Critical | Initial: **Charge**, Others: **Deny** |
| 7 | Critical only (no Initial/Subsequent) | Critical: **Charge** |
| 8 | Discharge only (no Initial/Critical) | Discharge: **Charge**, Subsequent: **Deny** |
| 9 | Single codes | **Charge** |

## Architecture

```
User Browser
     │
     ▼
┌─────────────────────────────────────────┐
│  CloudFront (CDN + SSL)                 │
│  Custom domain with ACM certificate     │
└─────────────────────────────────────────┘
     │                    │
     │ Static files       │ /api/*
     ▼                    ▼
┌──────────────┐  ┌─────────────────────────┐
│  S3 Bucket   │  │  API Gateway (HTTP)     │
│  (private)   │  │  POST /api/upload       │
└──────────────┘  │  GET  /api/health       │
                  └─────────────────────────┘
                              │
                              ▼
                  ┌─────────────────────────┐
                  │  Lambda Function        │
                  │  Python 3.12            │
                  │  + pandas layer         │
                  │  + openpyxl layer       │
                  └─────────────────────────┘
```

## Prerequisites

- AWS Account with CLI configured
- Python 3.12
- Route 53 hosted zone for your domain
- PowerShell (Windows) or adapt script for bash

## Deployment

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/chargescrub.git
cd chargescrub
```

### 2. Configure your deployment

```powershell
# Copy the example config
Copy-Item deploy-config.example.ps1 deploy-config.ps1

# Edit with your values
notepad deploy-config.ps1
```

Fill in your AWS details:
- `$Domain` — Your custom domain (e.g., `scrub.yourdomain.com`)
- `$Region` — AWS region (e.g., `us-west-2`)
- `$AccountId` — Your 12-digit AWS account ID
- `$HostedZoneId` — Route 53 hosted zone ID for your domain

### 3. Run deployment

```powershell
# May need to set execution policy first
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Deploy
.\deploy.ps1
```

The script will:
1. Create IAM role for Lambda
2. Create S3 bucket for static files
3. Build and deploy Lambda layers (pandas + openpyxl)
4. Create Lambda function
5. Create API Gateway
6. Request and validate ACM certificate
7. Create CloudFront distribution
8. Configure Route 53 DNS

**Total deployment time: ~15-20 minutes** (mostly waiting for certificate validation and CloudFront)

## Usage

1. Navigate to your deployed URL
2. Upload an Excel file (.xlsx) with columns:
   - `PATIENT NAME` — Patient identifier
   - `DATE OF SERVICE` — Service date
   - `PROCEDURECODE` — E/M CPT code
   - `BILLINGPROVIDER` — Provider name
3. Click "Process Charges"
4. Download the color-coded results

## Input File Format

| Column | Description | Required |
|--------|-------------|----------|
| PATIENT NAME | Patient identifier | ✅ |
| DATE OF SERVICE | Service date | ✅ |
| PROCEDURECODE | E/M CPT code (99221-99239, 99291, 99497, 99499) | ✅ |
| BILLINGPROVIDER | Provider name | ✅ |
| DIAGNOSIS CODES | ICD-10 codes | Optional |

## Security

ChargeScrub is designed with HIPAA alignment in mind:

- **Zero Storage** — Files processed in-memory only, never written to disk
- **No PHI Logging** — Patient names and charges are never logged to CloudWatch
- **Encrypted Transit** — TLS 1.2+ via CloudFront
- **Serverless** — No persistent servers, ephemeral compute only
- **Private S3** — Bucket accessible only via CloudFront OAC
- **CORS Restricted** — API only accepts requests from your domain

### Future Security Enhancements

- Add Cognito authentication for user management
- Implement API Gateway throttling
- Add AWS WAF for additional protection

## Costs

| Service | Typical Usage | Monthly Cost |
|---------|---------------|--------------|
| Lambda | ~4 invocations | Free tier |
| API Gateway | ~4 requests | Free tier |
| S3 | ~50KB stored | $0.00 |
| CloudFront | ~1MB transfer | Free tier |
| Route 53 | Hosted zone | $0.50 |
| ACM | Certificate | Free |

**Estimated total: ~$0.50/month** for light usage

## Development

### Update Lambda code

```powershell
Compress-Archive -Path .\lambda\lambda_function.py -DestinationPath lambda-code.zip -Force
aws lambda update-function-code --function-name ChargeScrub-Processor --zip-file fileb://lambda-code.zip --region us-west-2
```

### Update frontend

```powershell
aws s3 sync .\frontend\ s3://your-bucket-name/ --region us-west-2
aws cloudfront create-invalidation --distribution-id YOUR_DIST_ID --paths "/*"
```

### View Lambda logs

```powershell
aws logs tail /aws/lambda/ChargeScrub-Processor --region us-west-2 --follow
```

## License

MIT License — See [LICENSE](LICENSE) for details.

## Author

**RIGGSMED LLC**

---

*Built for hospitalist billing workflows. Not a substitute for professional coding judgment.*
