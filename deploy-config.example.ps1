# ============================================================================
# ChargeScrub Deployment Configuration
# ============================================================================
# 
# Copy this file to deploy-config.ps1 and fill in your values.
# DO NOT commit deploy-config.ps1 to git (it's in .gitignore)
#
# ============================================================================

# Your custom domain (must have Route 53 hosted zone)
$Domain = "scrub.yourdomain.com"

# AWS Region for Lambda, S3, API Gateway
$Region = "us-west-2"

# Your AWS Account ID (12 digits)
$AccountId = "123456789012"

# Route 53 Hosted Zone ID for your domain
$HostedZoneId = "ZXXXXXXXXXXXXX"
