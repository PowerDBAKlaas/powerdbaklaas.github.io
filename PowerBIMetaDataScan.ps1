# ============================================================================
# Power BI Metadata Scanning - Complete Script
# Generated bij Claude AI !!!
# To be reviewed !!!
# ============================================================================

# 1. AUTHENTICATION - Get Bearer Token
# ============================================================================

# Method A: Interactive (easiest for testing)
Connect-PowerBIServiceAccount

# Get token from the session
$tokenResponse = Invoke-PowerBIRestMethod -Url "admin/workspaces" -Method Get
# Extract token from the module's internal session
$token = (Get-PowerBIAccessToken)["Authorization"].Replace("Bearer ", "")

# Method B: Service Principal (for automation)
<#
$tenantId = "your-tenant-id"
$clientId = "your-app-id"
$clientSecret = "your-client-secret"

$body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    resource      = "https://analysis.windows.net/powerbi/api"
}

$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $body
$token = $tokenResponse.access_token
#>

# 2. INITIATE SCAN
# ============================================================================

$headers = @{
    'Authorization' = "Bearer $token"
    'Content-Type'  = 'application/json'
}

# Base URI - NO personalization needed, this is the standard endpoint
$baseUri = "https://api.powerbi.com/v1.0/myorg/admin"

# Scan request body - specify what metadata to retrieve
$scanBody = @{
    workspaces = @()  # Empty array = scan ALL workspaces in tenant
} | ConvertTo-Json

# Optional: Scan specific workspaces only
<#
$scanBody = @{
    workspaces = @(
        "workspace-guid-1",
        "workspace-guid-2"
    )
} | ConvertTo-Json
#>

# Parameters for scan (all optional, but recommended for complete metadata)
$scanParams = @(
    "lineage=True",                    # Data lineage information
    "datasourceDetails=True",          # Connection strings, credentials info
    "datasetSchema=True",              # Table/column schemas
    "datasetExpressions=True",         # DAX measures, calculated columns
    "getArtifactUsers=True"            # User access permissions
) -join "&"

$scanUrl = "$baseUri/workspaces/getInfo?$scanParams"

Write-Host "Initiating metadata scan..." -ForegroundColor Cyan

try {
    $scanResponse = Invoke-RestMethod -Method Post -Uri $scanUrl -Headers $headers -Body $scanBody
    $scanId = $scanResponse.id
    Write-Host "Scan initiated successfully. Scan ID: $scanId" -ForegroundColor Green
}
catch {
    Write-Host "Error initiating scan: $_" -ForegroundColor Red
    exit
}

# 3. POLL SCAN STATUS
# ============================================================================

$statusUrl = "$baseUri/workspaces/scanStatus/$scanId"
$maxWaitMinutes = 30
$waitSeconds = 10
$elapsedMinutes = 0

Write-Host "Waiting for scan to complete..." -ForegroundColor Cyan

do {
    Start-Sleep -Seconds $waitSeconds
    $elapsedMinutes += $waitSeconds / 60
    
    $status = Invoke-RestMethod -Method Get -Uri $statusUrl -Headers $headers
    
    Write-Host "Status: $($status.status) - Elapsed: $([math]::Round($elapsedMinutes, 1)) minutes" -ForegroundColor Yellow
    
    if ($elapsedMinutes -ge $maxWaitMinutes) {
        Write-Host "Scan timeout after $maxWaitMinutes minutes" -ForegroundColor Red
        exit
    }
    
} while ($status.status -ne "Succeeded")

Write-Host "Scan completed successfully!" -ForegroundColor Green

# 4. GET SCAN RESULTS
# ============================================================================

$resultUrl = "$baseUri/workspaces/scanResult/$scanId"

Write-Host "Retrieving scan results..." -ForegroundColor Cyan

$scanResults = Invoke-RestMethod -Method Get -Uri $resultUrl -Headers $headers

# Save raw results
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$rawOutputPath = "PowerBI_Metadata_Raw_$timestamp.json"
$scanResults | ConvertTo-Json -Depth 100 | Out-File -FilePath $rawOutputPath -Encoding UTF8

Write-Host "Raw results saved to: $rawOutputPath" -ForegroundColor Green

# 5. PARSE AND STRUCTURE RESULTS
# ============================================================================

Write-Host "Parsing and structuring results..." -ForegroundColor Cyan

# --- WORKSPACES ---
$workspaceInventory = foreach ($ws in $scanResults.workspaces) {
    [PSCustomObject]@{
        WorkspaceId          = $ws.id
        WorkspaceName        = $ws.name
        Type                 = $ws.type
        State                = $ws.state
        CapacityId           = $ws.capacityId
        DatasetCount         = ($ws.datasets).Count
        ReportCount          = ($ws.reports).Count
        DashboardCount       = ($ws.dashboards).Count
        DataflowCount        = ($ws.dataflows).Count
        WorkbookCount        = ($ws.workbooks).Count
    }
}

$workspaceInventory | Export-Csv -Path "Workspaces_$timestamp.csv" -NoTypeInformation

# --- DATASETS ---
$datasetInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        
        # Parse endorsement
        $endorsement = if ($ds.endorsementDetails) { 
            $ds.endorsementDetails.endorsement 
        } else { 
            "None" 
        }
        
        # Parse sensitivity label
        $sensitivityLabel = if ($ds.sensitivityLabel) { 
            $ds.sensitivityLabel.labelId 
        } else { 
            "None" 
        }
        
        [PSCustomObject]@{
            WorkspaceId              = $ws.id
            WorkspaceName            = $ws.name
            DatasetId                = $ds.id
            DatasetName              = $ds.name
            ConfiguredBy             = $ds.configuredBy
            CreatedDate              = $ds.createdDate
            ContentProviderType      = $ds.contentProviderType
            Endorsement              = $endorsement
            SensitivityLabel         = $sensitivityLabel
            IsRefreshable            = $ds.isRefreshable
            IsEffectiveIdentityRequired = $ds.isEffectiveIdentityRequired
            IsOnPremGatewayRequired  = $ds.isOnPremGatewayRequired
            TableCount               = ($ds.tables).Count
            DatasourceCount          = ($ds.datasources).Count
            UpstreamDataflowCount    = ($ds.upstreamDataflows).Count
        }
    }
}

$datasetInventory | Export-Csv -Path "Datasets_$timestamp.csv" -NoTypeInformation

# --- DATA SOURCES ---
$datasourceInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        foreach ($source in $ds.datasources) {
            [PSCustomObject]@{
                WorkspaceId       = $ws.id
                WorkspaceName     = $ws.name
                DatasetId         = $ds.id
                DatasetName       = $ds.name
                DatasourceType    = $source.datasourceType
                ConnectionDetails = $source.connectionDetails.url ?? $source.connectionDetails.server ?? $source.connectionDetails.path
                Database          = $source.connectionDetails.database
                GatewayId         = $source.gatewayId
            }
        }
    }
}

$datasourceInventory | Export-Csv -Path "DataSources_$timestamp.csv" -NoTypeInformation

# --- DATASET TABLES & COLUMNS ---
$schemaInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        foreach ($table in $ds.tables) {
            foreach ($column in $table.columns) {
                [PSCustomObject]@{
                    WorkspaceId     = $ws.id
                    DatasetId       = $ds.id
                    DatasetName     = $ds.name
                    TableName       = $table.name
                    ColumnName      = $column.name
                    DataType        = $column.dataType
                    IsHidden        = $column.isHidden
                    ColumnType      = $column.columnType  # Data, Calculated, RowNumber
                }
            }
        }
    }
}

$schemaInventory | Export-Csv -Path "DatasetSchema_$timestamp.csv" -NoTypeInformation

# --- MEASURES (DAX Expressions) ---
$measureInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        foreach ($table in $ds.tables) {
            foreach ($measure in $table.measures) {
                [PSCustomObject]@{
                    WorkspaceId   = $ws.id
                    DatasetId     = $ds.id
                    DatasetName   = $ds.name
                    TableName     = $table.name
                    MeasureName   = $measure.name
                    Expression    = $measure.expression
                    IsHidden      = $measure.isHidden
                }
            }
        }
    }
}

$measureInventory | Export-Csv -Path "Measures_$timestamp.csv" -NoTypeInformation

# --- REPORTS ---
$reportInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($report in $ws.reports) {
        [PSCustomObject]@{
            WorkspaceId      = $ws.id
            WorkspaceName    = $ws.name
            ReportId         = $report.id
            ReportName       = $report.name
            DatasetId        = $report.datasetId
            CreatedBy        = $report.createdBy
            CreatedDate      = $report.createdDateTime
            ModifiedBy       = $report.modifiedBy
            ModifiedDate     = $report.modifiedDateTime
            ReportType       = $report.reportType
            Endorsement      = $report.endorsementDetails.endorsement ?? "None"
        }
    }
}

$reportInventory | Export-Csv -Path "Reports_$timestamp.csv" -NoTypeInformation

# --- LINEAGE (Dataset Dependencies) ---
$lineageInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        # Upstream dataflows
        foreach ($dataflow in $ds.upstreamDataflows) {
            [PSCustomObject]@{
                SourceType        = "Dataflow"
                SourceWorkspaceId = $dataflow.groupId
                SourceId          = $dataflow.objectId
                TargetType        = "Dataset"
                TargetWorkspaceId = $ws.id
                TargetId          = $ds.id
                TargetName        = $ds.name
            }
        }
        
        # Upstream datasets (composite models)
        foreach ($upstreamDs in $ds.upstreamDatasets) {
            [PSCustomObject]@{
                SourceType        = "Dataset"
                SourceWorkspaceId = $upstreamDs.groupId
                SourceId          = $upstreamDs.objectId
                TargetType        = "Dataset"
                TargetWorkspaceId = $ws.id
                TargetId          = $ds.id
                TargetName        = $ds.name
            }
        }
    }
}

$lineageInventory | Export-Csv -Path "Lineage_$timestamp.csv" -NoTypeInformation

# --- USER ACCESS (if getArtifactUsers=True) ---
$userAccessInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        foreach ($user in $ds.users) {
            [PSCustomObject]@{
                WorkspaceId       = $ws.id
                ArtifactType      = "Dataset"
                ArtifactId        = $ds.id
                ArtifactName      = $ds.name
                UserEmailAddress  = $user.identifier
                PrincipalType     = $user.principalType
                AccessRight       = $user.datasetUserAccessRight
            }
        }
    }
    
    foreach ($report in $ws.reports) {
        foreach ($user in $report.users) {
            [PSCustomObject]@{
                WorkspaceId       = $ws.id
                ArtifactType      = "Report"
                ArtifactId        = $report.id
                ArtifactName      = $report.name
                UserEmailAddress  = $user.identifier
                PrincipalType     = $user.principalType
                AccessRight       = $user.reportUserAccessRight
            }
        }
    }
}

$userAccessInventory | Export-Csv -Path "UserAccess_$timestamp.csv" -NoTypeInformation

# 6. SUMMARY REPORT
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SCAN COMPLETE - SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Workspaces:   $($workspaceInventory.Count)" -ForegroundColor White
Write-Host "Total Datasets:     $($datasetInventory.Count)" -ForegroundColor White
Write-Host "Total Reports:      $($reportInventory.Count)" -ForegroundColor White
Write-Host "Total Data Sources: $($datasourceInventory.Count)" -ForegroundColor White
Write-Host "Total Measures:     $($measureInventory.Count)" -ForegroundColor White
Write-Host "`nFiles created:" -ForegroundColor Cyan
Write-Host "  - $rawOutputPath (raw JSON)" -ForegroundColor White
Write-Host "  - Workspaces_$timestamp.csv" -ForegroundColor White
Write-Host "  - Datasets_$timestamp.csv" -ForegroundColor White
Write-Host "  - DataSources_$timestamp.csv" -ForegroundColor White
Write-Host "  - DatasetSchema_$timestamp.csv" -ForegroundColor White
Write-Host "  - Measures_$timestamp.csv" -ForegroundColor White
Write-Host "  - Reports_$timestamp.csv" -ForegroundColor White
Write-Host "  - Lineage_$timestamp.csv" -ForegroundColor White
Write-Host "  - UserAccess_$timestamp.csv" -ForegroundColor White
```

## Key Points Explained

### **Token Acquisition**
- **Method A (Interactive)**: Uses `Connect-PowerBIServiceAccount` - simplest for ad-hoc runs
- **Method B (Service Principal)**: For scheduled automation - requires Azure AD app registration

### **URI Personalization**
- **NO personalization needed** - the URI is a standard Microsoft endpoint
- The only variable is `$scanId` which is returned by the API itself
- Your tenant context is determined by the authentication token

### **Scan Results Structure**

The JSON response has this hierarchy:
```
workspaces[]
├── id, name, type, state, capacityId
├── datasets[]
│   ├── id, name, tables[], measures[], datasources[]
│   ├── endorsementDetails
│   ├── sensitivityLabel
│   └── upstreamDataflows[], upstreamDatasets[]
├── reports[]
│   ├── id, name, datasetId, users[]
│   └── endorsementDetails
├── dashboards[]
├── dataflows[]
└── workbooks[]
