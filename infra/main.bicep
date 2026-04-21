// ════════════════════════════════════════════════════════════════════════════════
// Bicep template: Editorial QA Agent — Production Azure Deployment
// Deploy with: az deployment sub create --location eastus2 --template-file infra/main.bicep
// ════════════════════════════════════════════════════════════════════════════════

targetScope = 'resourceGroup'

@description('Base name for all resources')
param baseName string = 'editorial-qa'

@description('Azure region')
param location string = resourceGroup().location

@description('Azure AI Foundry project endpoint')
@secure()
param foundryEndpoint string

@description('Default model deployment name')
param defaultModel string = 'gpt-4o'

@description('Container image (ACR login server/repo:tag)')
param containerImage string

@description('Minimum replicas (0 = scale to zero)')
@minValue(0)
@maxValue(10)
param minReplicas int = 0

@description('Maximum replicas')
@minValue(1)
@maxValue(50)
param maxReplicas int = 10

// ── Variables ──
var uniqueSuffix = uniqueString(resourceGroup().id, baseName)
var acrName = replace('acr${baseName}${uniqueSuffix}', '-', '')
var caeName = '${baseName}-cae-${uniqueSuffix}'
var caName = '${baseName}-app'
var kvName = '${baseName}-kv-${uniqueSuffix}'
var logName = '${baseName}-log-${uniqueSuffix}'
var aiName = '${baseName}-ai-${uniqueSuffix}'
var storageName = replace('st${baseName}${uniqueSuffix}', '-', '')

// ════════════════════════════════════════════════════════════════════════════════
// Log Analytics Workspace
// ════════════════════════════════════════════════════════════════════════════════
resource logWorkspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: logName
  location: location
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Application Insights
// ════════════════════════════════════════════════════════════════════════════════
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: aiName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logWorkspace.id
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Azure Container Registry
// ════════════════════════════════════════════════════════════════════════════════
resource acr 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  name: acrName
  location: location
  sku: { name: 'Basic' }
  properties: {
    adminUserEnabled: false
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Azure Key Vault
// ════════════════════════════════════════════════════════════════════════════════
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: kvName
  location: location
  properties: {
    sku: { family: 'A', name: 'standard' }
    tenantId: subscription().tenantId
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
  }
}

// Store Foundry endpoint as a secret
resource kvSecretEndpoint 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'FoundryEndpoint'
  properties: {
    value: foundryEndpoint
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Azure Blob Storage (for PPTX uploads & results)
// ════════════════════════════════════════════════════════════════════════════════
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' = {
  parent: storageAccount
  name: 'default'
}

resource uploadsContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: blobService
  name: 'pptx-uploads'
}

resource resultsContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: blobService
  name: 'review-results'
}

// ════════════════════════════════════════════════════════════════════════════════
// Container Apps Environment
// ════════════════════════════════════════════════════════════════════════════════
resource cae 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: caeName
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logWorkspace.properties.customerId
        sharedKey: logWorkspace.listKeys().primarySharedKey
      }
    }
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Container App — Editorial QA Agent
// ════════════════════════════════════════════════════════════════════════════════
resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: caName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    managedEnvironmentId: cae.id
    configuration: {
      ingress: {
        external: true
        targetPort: 7860
        transport: 'auto'
        allowInsecure: false
      }
      registries: [
        {
          server: acr.properties.loginServer
          identity: 'system'
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'editorial-qa'
          image: containerImage
          resources: {
            cpu: json('1.0')
            memory: '2Gi'
          }
          env: [
            { name: 'AZURE_AI_PROJECT_ENDPOINT', secretRef: 'foundry-endpoint' }
            { name: 'AZURE_AI_MODEL_DEPLOYMENT_NAME', value: defaultModel }
            { name: 'CONTAINER', value: 'true' }
            { name: 'PORT', value: '7860' }
            { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: appInsights.properties.ConnectionString }
          ]
        }
      ]
      scale: {
        minReplicas: minReplicas
        maxReplicas: maxReplicas
        rules: [
          {
            name: 'http-scaling'
            http: {
              metadata: {
                concurrentRequests: '10'
              }
            }
          }
        ]
      }
    }
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// RBAC: Container App Managed Identity → ACR Pull
// ════════════════════════════════════════════════════════════════════════════════
var acrPullRoleId = '7f951dda-4ed3-4680-a7ca-43fe172d538d'
resource acrPullRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(acr.id, containerApp.id, acrPullRoleId)
  scope: acr
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', acrPullRoleId)
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// RBAC: Container App Managed Identity → Key Vault Secrets User
// ════════════════════════════════════════════════════════════════════════════════
var kvSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'
resource kvSecretsRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, containerApp.id, kvSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', kvSecretsUserRoleId)
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// RBAC: Container App Managed Identity → Storage Blob Data Contributor
// ════════════════════════════════════════════════════════════════════════════════
var storageBlobContribRoleId = 'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, containerApp.id, storageBlobContribRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobContribRoleId)
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// Outputs
// ════════════════════════════════════════════════════════════════════════════════
output appUrl string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output acrLoginServer string = acr.properties.loginServer
output appInsightsKey string = appInsights.properties.InstrumentationKey
output managedIdentityId string = containerApp.identity.principalId
