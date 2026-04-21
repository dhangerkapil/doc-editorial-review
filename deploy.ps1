# ════════════════════════════════════════════════════════════════════════════════
# Editorial QA Agent — Production Deployment Guide
# ════════════════════════════════════════════════════════════════════════════════
#
# Prerequisites:
#   1. Azure CLI logged in: az login
#   2. Docker installed
#   3. Azure AI Foundry project with model deployments
#
# Usage:
#   # Set your values
#   $RG = "rg-editorial-qa"
#   $LOCATION = "eastus2"
#   $FOUNDRY_ENDPOINT = "https://your-resource.services.ai.azure.com/api/projects/your-project"
#
#   # Run the deployment (interactive — prompts for confirmation)
#   .\deploy.ps1 -ResourceGroup $RG -Location $LOCATION -FoundryEndpoint $FOUNDRY_ENDPOINT
# ════════════════════════════════════════════════════════════════════════════════

param(
    [Parameter(Mandatory)]
    [string]$ResourceGroup,

    [Parameter(Mandatory)]
    [string]$Location,

    [Parameter(Mandatory)]
    [string]$FoundryEndpoint,

    [string]$BaseName = "editorial-qa",
    [string]$DefaultModel = "gpt-4o",
    [int]$MinReplicas = 0,
    [int]$MaxReplicas = 10,
    [string]$ImageTag = "latest"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Write-Host "`n═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Editorial QA Agent — Azure Deployment" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# ── Step 1: Create Resource Group ──
Write-Host "[1/5] Creating resource group '$ResourceGroup' in '$Location'..." -ForegroundColor Yellow
az group create --name $ResourceGroup --location $Location --output none
Write-Host "  ✅ Resource group ready" -ForegroundColor Green

# ── Step 2: Deploy Infrastructure (Bicep) ──
Write-Host "`n[2/5] Deploying infrastructure via Bicep..." -ForegroundColor Yellow
$deployOutput = az deployment group create `
    --resource-group $ResourceGroup `
    --template-file infra/main.bicep `
    --parameters baseName=$BaseName `
                 location=$Location `
                 foundryEndpoint=$FoundryEndpoint `
                 defaultModel=$DefaultModel `
                 containerImage="mcr.microsoft.com/hello-world:latest" `
                 minReplicas=$MinReplicas `
                 maxReplicas=$MaxReplicas `
    --query "properties.outputs" `
    --output json | ConvertFrom-Json

$ACR_SERVER = $deployOutput.acrLoginServer.value
$APP_URL = $deployOutput.appUrl.value
$MI_ID = $deployOutput.managedIdentityId.value

Write-Host "  ✅ Infrastructure deployed" -ForegroundColor Green
Write-Host "     ACR: $ACR_SERVER" -ForegroundColor Gray
Write-Host "     App URL: $APP_URL" -ForegroundColor Gray

# ── Step 3: Build & Push Container Image ──
Write-Host "`n[3/5] Building and pushing container image..." -ForegroundColor Yellow
az acr login --name ($ACR_SERVER -replace '\.azurecr\.io','')
$IMAGE = "${ACR_SERVER}/editorial-qa:${ImageTag}"
docker build -t $IMAGE .
docker push $IMAGE
Write-Host "  ✅ Image pushed: $IMAGE" -ForegroundColor Green

# ── Step 4: Update Container App with real image ──
Write-Host "`n[4/5] Updating container app with production image..." -ForegroundColor Yellow
az containerapp update `
    --name "$BaseName-app" `
    --resource-group $ResourceGroup `
    --image $IMAGE `
    --output none
Write-Host "  ✅ Container app updated" -ForegroundColor Green

# ── Step 5: Grant Managed Identity access to AI Foundry ──
Write-Host "`n[5/5] Granting Managed Identity access to AI Foundry..." -ForegroundColor Yellow
Write-Host "  ℹ️  Managed Identity Principal ID: $MI_ID" -ForegroundColor Gray
Write-Host "  ℹ️  You need to grant this identity 'Azure AI Developer' role on your AI Foundry resource" -ForegroundColor Gray
Write-Host "  ℹ️  Run: az role assignment create --assignee $MI_ID --role 'Azure AI Developer' --scope <foundry-resource-id>" -ForegroundColor Gray

# ── Done ──
Write-Host "`n═══════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  ✅ Deployment Complete!" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "`n  🌐 App URL: $APP_URL" -ForegroundColor Cyan
Write-Host "  🐳 Image:   $IMAGE" -ForegroundColor Cyan
Write-Host "  🔑 MI ID:   $MI_ID" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Yellow
Write-Host "    1. Grant MI 'Azure AI Developer' role on Foundry resource" -ForegroundColor White
Write-Host "    2. Update app.py to use DefaultAzureCredential() instead of AzureCliCredential()" -ForegroundColor White
Write-Host "    3. Configure Azure Front Door for production traffic" -ForegroundColor White
Write-Host "    4. Set up Entra ID authentication on the Container App" -ForegroundColor White
Write-Host ""
