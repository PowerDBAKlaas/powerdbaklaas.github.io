# ============================================================================
# Power BI Metadata Scanning - Complete Script
# Generated bij Claude AI !!!
# To be reviewed !!!
# ============================================================================

# 1. AUTHENTICATION
# ============================================================================

Connect-PowerBIServiceAccount
$token = (Get-PowerBIAccessToken)["Authorization"].Replace("Bearer ", "")

$headers = @{
    'Authorization' = "Bearer $token"
    'Content-Type'  = 'application/json; charset=utf-8'
}

$baseUri = "https://api.powerbi.com/v1.0/myorg/admin"
$scanParams = "lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
$scanUrl = "$baseUri/workspaces/getInfo?$scanParams"

# 2. GET ALL WORKSPACES
# ============================================================================

Write-Host "Retrieving all workspaces..." -ForegroundColor Cyan
$allWorkspaces = Get-PowerBIWorkspace -Scope Organization -All
$workspaceIds = $allWorkspaces | ForEach-Object { $_.Id.ToString() }

Write-Host "Found $($workspaceIds.Count) total workspaces" -ForegroundColor Green

# 3. SPLIT INTO BATCHES OF 100
# ============================================================================

$batchSize = 100
$totalBatches = [Math]::Ceiling($workspaceIds.Count / $batchSize)
$allScanResults = @()

Write-Host "Will process $totalBatches batches of up to $batchSize workspaces each`n" -ForegroundColor Yellow

for ($batchNum = 0; $batchNum -lt $totalBatches; $batchNum++) {
    
    $startIndex = $batchNum * $batchSize
    $endIndex = [Math]::Min($startIndex + $batchSize - 1, $workspaceIds.Count - 1)
    $batchIds = $workspaceIds[$startIndex..$endIndex]
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "BATCH $($batchNum + 1) of $totalBatches" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Workspaces: $($batchIds.Count) (indices $startIndex to $endIndex)" -ForegroundColor White
    
    # 4. BUILD SCAN REQUEST FOR THIS BATCH
    # ========================================================================
    
    $workspaceJsonArray = ($batchIds | ForEach-Object { "`"$_`"" }) -join ",`n    "
    $scanBody = @"
{
  "workspaces": [
    $workspaceJsonArray
  ]
}
"@
    
    # 5. INITIATE SCAN
    # ========================================================================
    
    Write-Host "Initiating scan..." -ForegroundColor Yellow
    
    try {
        $scanResponse = Invoke-RestMethod -Method Post -Uri $scanUrl -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($scanBody))
        $scanId = $scanResponse.id
        Write-Host "✓ Scan initiated - ID: $scanId" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Scan initiation failed!" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        continue  # Skip to next batch
    }
    
    # 6. POLL FOR COMPLETION
    # ========================================================================
    
    $statusUrl = "$baseUri/workspaces/scanStatus/$scanId"
    $maxWaitMinutes = 30
    $waitSeconds = 10
    $elapsedSeconds = 0
    
    Write-Host "Waiting for scan completion..." -ForegroundColor Yellow
    
    do {
        Start-Sleep -Seconds $waitSeconds
        $elapsedSeconds += $waitSeconds
        
        try {
            $status = Invoke-RestMethod -Method Get -Uri $statusUrl -Headers $headers
            Write-Host "  Status: $($status.status) - Elapsed: $([Math]::Round($elapsedSeconds / 60, 1)) min" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Error checking status: $($_.Exception.Message)" -ForegroundColor Red
            break
        }
        
        if ($elapsedSeconds -ge ($maxWaitMinutes * 60)) {
            Write-Host "  Timeout after $maxWaitMinutes minutes" -ForegroundColor Red
            break
        }
        
    } while ($status.status -ne "Succeeded")
    
    if ($status.status -ne "Succeeded") {
        Write-Host "✗ Batch $($batchNum + 1) did not complete successfully (Status: $($status.status))" -ForegroundColor Red
        continue  # Skip to next batch
    }
    
    # 7. GET SCAN RESULTS
    # ========================================================================
    
    Write-Host "Retrieving results..." -ForegroundColor Yellow
    
    try {
        $resultUrl = "$baseUri/workspaces/scanResult/$scanId"
        $batchResults = Invoke-RestMethod -Method Get -Uri $resultUrl -Headers $headers
        
        # Add to accumulated results
        $allScanResults += $batchResults.workspaces
        
        Write-Host "✓ Batch $($batchNum + 1) complete: $($batchResults.workspaces.Count) workspaces retrieved`n" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to retrieve results: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }
    
    # Small delay between batches to avoid rate limiting
    if ($batchNum -lt ($totalBatches - 1)) {
        Write-Host "Pausing 5 seconds before next batch...`n" -ForegroundColor Gray
        Start-Sleep -Seconds 5
    }
}

# 8. CREATE COMBINED RESULTS OBJECT
# ============================================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ALL BATCHES COMPLETE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total workspaces retrieved: $($allScanResults.Count) of $($workspaceIds.Count) requested`n" -ForegroundColor Green

$scanResults = [PSCustomObject]@{
    workspaces = $allScanResults
}

# 9. SAVE RAW COMBINED RESULTS
# ============================================================================

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$rawOutputPath = "PowerBI_Metadata_Raw_$timestamp.json"

Write-Host "Saving raw results..." -ForegroundColor Cyan
$scanResults | ConvertTo-Json -Depth 100 | Out-File -FilePath $rawOutputPath -Encoding UTF8
Write-Host "✓ Raw results saved to: $rawOutputPath`n" -ForegroundColor Green

# 10. PARSE AND STRUCTURE RESULTS
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
Write-Host "✓ Workspaces exported: $($workspaceInventory.Count)" -ForegroundColor White

# --- DATASETS ---
$datasetInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
        
        $endorsement = if ($ds.endorsementDetails) { 
            $ds.endorsementDetails.endorsement 
        } else { 
            "None" 
        }
        
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
Write-Host "✓ Datasets exported: $($datasetInventory.Count)" -ForegroundColor White

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
Write-Host "✓ Data sources exported: $($datasourceInventory.Count)" -ForegroundColor White

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
                    ColumnType      = $column.columnType
                }
            }
        }
    }
}

$schemaInventory | Export-Csv -Path "DatasetSchema_$timestamp.csv" -NoTypeInformation
Write-Host "✓ Schema exported: $($schemaInventory.Count) columns" -ForegroundColor White

# --- MEASURES ---
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
Write-Host "✓ Measures exported: $($measureInventory.Count)" -ForegroundColor White

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
Write-Host "✓ Reports exported: $($reportInventory.Count)" -ForegroundColor White

# --- LINEAGE ---
$lineageInventory = foreach ($ws in $scanResults.workspaces) {
    foreach ($ds in $ws.datasets) {
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
Write-Host "✓ Lineage exported: $($lineageInventory.Count) relationships" -ForegroundColor White

# --- USER ACCESS ---
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
Write-Host "✓ User access exported: $($userAccessInventory.Count) permissions" -ForegroundColor White

# 11. FINAL SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "EXPORT COMPLETE - SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Workspaces:   $($workspaceInventory.Count)" -ForegroundColor White
Write-Host "Total Datasets:     $($datasetInventory.Count)" -ForegroundColor White
Write-Host "Total Reports:      $($reportInventory.Count)" -ForegroundColor White
Write-Host "Total Data Sources: $($datasourceInventory.Count)" -ForegroundColor White
Write-Host "Total Measures:     $($measureInventory.Count)" -ForegroundColor White
Write-Host "Total Lineage:      $($lineageInventory.Count)" -ForegroundColor White
Write-Host "`nFiles created:" -ForegroundColor Cyan
Write-Host "  - $rawOutputPath" -ForegroundColor White
Write-Host "  - Workspaces_$timestamp.csv" -ForegroundColor White
Write-Host "  - Datasets_$timestamp.csv" -ForegroundColor White
Write-Host "  - DataSources_$timestamp.csv" -ForegroundColor White
Write-Host "  - DatasetSchema_$timestamp.csv" -ForegroundColor White
Write-Host "  - Measures_$timestamp.csv" -ForegroundColor White
Write-Host "  - Reports_$timestamp.csv" -ForegroundColor White
Write-Host "  - Lineage_$timestamp.csv" -ForegroundColor White
Write-Host "  - UserAccess_$timestamp.csv" -ForegroundColor White
Write-Host "`n✓ All done!" -ForegroundColor Green


<#
```

## Key Features of This Batched Version

1. **Automatic Batching**: Splits all workspaces into groups of 100
2. **Progress Tracking**: Shows which batch is processing (e.g., "BATCH 3 of 15")
3. **Error Resilience**: If one batch fails, continues with the next
4. **Combined Results**: Merges all batches into single output files
5. **Rate Limit Protection**: 5-second pause between batches
6. **Detailed Logging**: Shows status for each batch

## Expected Output

For a tenant with 450 workspaces, you'll see:
```
Found 450 total workspaces
Will process 5 batches of up to 100 workspaces each

========================================
BATCH 1 of 5
========================================
Workspaces: 100 (indices 0 to 99)
Initiating scan...
✓ Scan initiated - ID: abc123...
...
✓ Batch 1 complete: 100 workspaces retrieved

[repeats for batches 2-5]

#>
