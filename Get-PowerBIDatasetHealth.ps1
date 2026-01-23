<#
Created by claude AI
to be reviewed!
#>

# ============================================================================
# Power BI Health Analysis - WITH GATEWAY & CAPACITY MONITORING
# ============================================================================

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "POWER BI HEALTH ANALYSIS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Import your existing inventory
Write-Host "Loading existing inventory..." -ForegroundColor Yellow
$datasets = Import-Csv "Datasets_*.csv" | Sort-Object DatasetId -Unique
$workspaces = Import-Csv "Workspaces_*.csv" | Sort-Object WorkspaceId -Unique
$datasources = Import-Csv "DataSources_*.csv"

Write-Host "✓ Loaded $($datasets.Count) datasets from $($workspaces.Count) workspaces`n" -ForegroundColor Green

# Connect
Write-Host "Connecting to Power BI..." -ForegroundColor Yellow
Connect-PowerBIServiceAccount
$token = (Get-PowerBIAccessToken)["Authorization"].Replace("Bearer ", "")

$headers = @{
    'Authorization' = $token
}

Write-Host "✓ Connected`n" -ForegroundColor Green

# ============================================================================
# 1. GET REFRESH HISTORY FOR ALL DATASETS
# ============================================================================

Write-Host "Retrieving refresh history (this may take a few minutes)..." -ForegroundColor Yellow
Write-Host "Progress: " -NoNewline -ForegroundColor Gray

$refreshHistory = @()
$totalDatasets = $datasets.Count
$currentCount = 0

foreach ($dataset in $datasets) {
    $currentCount++
    
    if ($currentCount % 10 -eq 0) {
       
Write-host "$currentCount/$totalDatasets " -NoNewline -ForegroundColor Gray
    }
    
    if ($dataset.IsRefreshable -ne "True") {
        continue
    }
    
    try {
        $uri = "https://api.powerbi.com/v1.0/myorg/datasets/$($dataset.DatasetId)/refreshes?`$top=5"
        $refreshes = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        
        foreach ($refresh in $refreshes.value) {
            $refreshHistory += [PSCustomObject]@{
                WorkspaceId          = $dataset.WorkspaceId
                WorkspaceName        = $dataset.WorkspaceName
                DatasetId            = $dataset.DatasetId
                DatasetName          = $dataset.DatasetName
                RefreshType          = $refresh.refreshType
                StartTime            = $refresh.startTime
                EndTime              = $refresh.endTime
                Status               = $refresh.status
                ServiceExceptionJson = $refresh.serviceExceptionJson
                RequestId            = $refresh.requestId
            }
        }
    }
    catch {
       
Write-host "!" -NoNewline -ForegroundColor Red
    }
    
    if ($currentCount % 100 -eq 0) {
        Start-Sleep -Seconds 2
    }
}

Write-Host "`n✓ Retrieved refresh history for $($refreshHistory.Count) refresh attempts`n" -ForegroundColor Green

$refreshHistory | Export-Csv "RefreshHistory_Raw_$timestamp.csv" -NoTypeInformation

# ============================================================================
# 2. GET GATEWAY INFORMATION
# ============================================================================

Write-Host "Retrieving gateway information..." -ForegroundColor Yellow

try {
    $gatewaysUri = "https://api.powerbi.com/v1.0/myorg/gateways"
    $gatewaysResponse = Invoke-RestMethod -Uri $gatewaysUri -Headers $headers -Method Get
    $gateways = $gatewaysResponse.value
    
   
Write-host "✓ Found $($gateways.Count) gateways`n" -ForegroundColor Green
}
catch {
   
Write-host "⚠️  Could not retrieve gateway list: $($_.Exception.Message)`n" -ForegroundColor Yellow
    $gateways = @()
}

# Get gateway details including datasources
$gatewayDetails = foreach ($gateway in $gateways) {
    try {
        # Get gateway datasources
        $datasourcesUri = "https://api.powerbi.com/v1.0/myorg/gateways/$($gateway.id)/datasources"
        $gatewayDatasources = Invoke-RestMethod -Uri $datasourcesUri -Headers $headers -Method Get -ErrorAction Stop
        
        [PSCustomObject]@{
            GatewayId              = $gateway.id
            GatewayName            = $gateway.name
            GatewayType            = $gateway.type
            GatewayStatus          = $gateway.gatewayStatus
            GatewayVersion         = $gateway.gatewayVersion
            PublicKey              = $gateway.publicKey.exponent
            DatasourceCount        = $gatewayDatasources.value.Count
            GatewayAnnotation      = $gateway.gatewayAnnotation
        }
    }
    catch {
        [PSCustomObject]@{
            GatewayId              = $gateway.id
            GatewayName            = $gateway.name
            GatewayType            = $gateway.type
            GatewayStatus          = "Unknown"
            GatewayVersion         = $gateway.gatewayVersion
            PublicKey              = $null
            DatasourceCount        = 0
            GatewayAnnotation      = $gateway.gatewayAnnotation
        }
    }
}

if ($gatewayDetails) {
    $gatewayDetails | Export-Csv "Gateways_$timestamp.csv" -NoTypeInformation
}

# ============================================================================
# 3. ANALYZE GATEWAY USAGE & HEALTH
# ============================================================================

Write-Host "Analyzing gateway health..." -ForegroundColor Yellow

# Map datasets to gateways
$gatewayDatasetMapping = $datasources | Where-Object { $_.GatewayId } | 
    Select-Object DatasetId, DatasetName, WorkspaceName, GatewayId -Unique

# Count datasets per gateway
$gatewayUsage = $gatewayDatasetMapping | Group-Object GatewayId | ForEach-Object {
    $gatewayId = $_.Name
    $gatewayInfo = $gatewayDetails | Where-Object { $_.GatewayId -eq $gatewayId }
    
    # Get refresh failures for datasets on this gateway
    $gatewayDatasetIds = ($_.Group | Select-Object -ExpandProperty DatasetId)
    $gatewayRefreshes = $refreshHistory | Where-Object { $gatewayDatasetIds -contains $_.DatasetId }
    $gatewayFailures = ($gatewayRefreshes | Where-Object { $_.Status -eq "Failed" }).Count
    $gatewayTotal = $gatewayRefreshes.Count
    
    $failureRate = if ($gatewayTotal -gt 0) {
        [Math]::Round(($gatewayFailures / $gatewayTotal) * 100, 1)
    } else {
        0
    }
    
    [PSCustomObject]@{
        GatewayId       = $gatewayId
        GatewayName     = $gatewayInfo.GatewayName
        GatewayStatus   = $gatewayInfo.GatewayStatus
        GatewayVersion  = $gatewayInfo.GatewayVersion
        DatasetCount    = $_.Count
        RefreshAttempts = $gatewayTotal
        FailedRefreshes = $gatewayFailures
        FailureRate     = $failureRate
        HealthStatus    = if ($failureRate -gt 30) { "Poor" } 
                         elseif ($failureRate -gt 10) { "Fair" } 
                         else { "Good" }
    }
} | Sort-Object FailureRate -Descending

if ($gatewayUsage) {
    $gatewayUsage | Export-Csv "GatewayHealth_$timestamp.csv" -NoTypeInformation
}

Write-Host "✓ Gateway analysis complete`n" -ForegroundColor Green

# ============================================================================
# 4. GET CAPACITY INFORMATION
# ============================================================================

Write-Host "Retrieving capacity information..." -ForegroundColor Yellow

try {
    # Get capacities
    $capacitiesUri = "https://api.powerbi.com/v1.0/myorg/capacities"
    $capacitiesResponse = Invoke-RestMethod -Uri $capacitiesUri -Headers $headers -Method Get
    $capacities = $capacitiesResponse.value
    
   
Write-host "✓ Found $($capacities.Count) capacities`n" -ForegroundColor Green
}
catch {
   
Write-host "⚠️  Could not retrieve capacity list: $($_.Exception.Message)`n" -ForegroundColor Yellow
    $capacities = @()
}

# Get capacity details
$capacityDetails = foreach ($capacity in $capacities) {
    [PSCustomObject]@{
        CapacityId          = $capacity.id
        DisplayName         = $capacity.displayName
        Admins              = ($capacity.admins -join "; ")
        Sku                 = $capacity.sku
        State               = $capacity.state
        CapacityUserAccessRight = $capacity.capacityUserAccessRight
        Region              = $capacity.region
    }
}

if ($capacityDetails) {
    $capacityDetails | Export-Csv "Capacities_$timestamp.csv" -NoTypeInformation
}

# ============================================================================
# 5. ANALYZE CAPACITY USAGE & STRESS
# ============================================================================

Write-Host "Analyzing capacity stress..." -ForegroundColor Yellow

# Map workspaces to capacities
$capacityUsage = $workspaces | Where-Object { $_.CapacityId } | 
    Group-Object CapacityId | ForEach-Object {
    
    $capacityId = $_.Name
    $capacityInfo = $capacityDetails | Where-Object { $_.CapacityId -eq $capacityId }
    
    # Get all datasets in this capacity
    $capacityWorkspaceIds = $_.Group | Select-Object -ExpandProperty WorkspaceId
    $capacityDatasets = $datasets | Where-Object { $capacityWorkspaceIds -contains $_.WorkspaceId }
    
    # Count refreshable datasets
    $refreshableCount = ($capacityDatasets | Where-Object { $_.IsRefreshable -eq "True" }).Count
    
    # Get refresh history for this capacity
    $capacityDatasetIds = $capacityDatasets | Select-Object -ExpandProperty DatasetId
    $capacityRefreshes = $refreshHistory | Where-Object { $capacityDatasetIds -contains $_.DatasetId }
    
    # Calculate refresh load (refreshes per day)
    $recentRefreshes = $capacityRefreshes | Where-Object { 
        [DateTime]$_.StartTime -gt (Get-Date).AddDays(-7) 
    }
    $refreshesPerDay = [Math]::Round(($recentRefreshes.Count / 7), 1)
    
    # Calculate average concurrent refreshes (approximation)
    # Group by 15-minute windows
    $concurrentRefreshes = $recentRefreshes | ForEach-Object {
        [PSCustomObject]@{
            TimeWindow = ([DateTime]$_.StartTime).ToString("yyyy-MM-dd HH:mm")
            DatasetId = $_.DatasetId
        }
    } | Group-Object TimeWindow | ForEach-Object { $_.Count } | Measure-Object -Average -Maximum
    
    $avgConcurrent = [Math]::Round($concurrentRefreshes.Average, 1)
    $maxConcurrent = $concurrentRefreshes.Maximum
    
    # Calculate total refresh duration per day
    $totalRefreshMinutes = ($recentRefreshes | Where-Object { $_.Status -eq "Completed" } | ForEach-Object {
        if ($_.StartTime -and $_.EndTime) {
            ([DateTime]$_.EndTime - [DateTime]$_.StartTime).TotalMinutes
        }
    } | Measure-Object -Sum).Sum
    
    $avgRefreshMinutesPerDay = [Math]::Round(($totalRefreshMinutes / 7), 1)
    
    # Stress level calculation
    $stressLevel = if ($refreshesPerDay -gt 100 -or $maxConcurrent -gt 5) {
        "High"
    } elseif ($refreshesPerDay -gt 50 -or $maxConcurrent -gt 3) {
        "Medium"
    } else {
        "Low"
    }
    
    [PSCustomObject]@{
        CapacityId               = $capacityId
        CapacityName             = $capacityInfo.DisplayName
        CapacitySku              = $capacityInfo.Sku
        State                    = $capacityInfo.State
        WorkspaceCount           = $_.Count
        TotalDatasets            = $capacityDatasets.Count
        RefreshableDatasets      = $refreshableCount
        RefreshesPerDay          = $refreshesPerDay
        AvgConcurrentRefreshes   = $avgConcurrent
        MaxConcurrentRefreshes   = $maxConcurrent
        AvgRefreshMinutesPerDay  = $avgRefreshMinutesPerDay
        StressLevel              = $stressLevel
    }
} | Sort-Object StressLevel -Descending

# Add shared capacity (workspaces without capacity)
$sharedCapacityWorkspaces = $workspaces | Where-Object { -not $_.CapacityId -or $_.CapacityId -eq "" }
if ($sharedCapacityWorkspaces) {
    $sharedDatasets = $datasets | Where-Object { 
        ($sharedCapacityWorkspaces | Select-Object -ExpandProperty WorkspaceId) -contains $_.WorkspaceId 
    }
    
    $sharedRefreshableCount = ($sharedDatasets | Where-Object { $_.IsRefreshable -eq "True" }).Count
    
    $capacityUsage += [PSCustomObject]@{
        CapacityId               = "Shared"
        CapacityName             = "Shared Capacity (No Premium)"
        CapacitySku              = "Shared"
        State                    = "Active"
        WorkspaceCount           = $sharedCapacityWorkspaces.Count
        TotalDatasets            = $sharedDatasets.Count
        RefreshableDatasets      = $sharedRefreshableCount
        RefreshesPerDay          = 0
        AvgConcurrentRefreshes   = 0
        MaxConcurrentRefreshes   = 0
        AvgRefreshMinutesPerDay  = 0
        StressLevel              = "N/A"
    }
}

if ($capacityUsage) {
    $capacityUsage | Export-Csv "CapacityUsage_$timestamp.csv" -NoTypeInformation
}

Write-Host "✓ Capacity analysis complete`n" -ForegroundColor Green

# ============================================================================
# 6. ANALYZE DATASET HEALTH
# ============================================================================

Write-Host "Analyzing dataset health..." -ForegroundColor Yellow

$datasetHealth = foreach ($dataset in ($datasets | Where-Object { $_.IsRefreshable -eq "True" })) {
    
    $datasetRefreshes = $refreshHistory | Where-Object { $_.DatasetId -eq $dataset.DatasetId } | 
        Sort-Object StartTime -Descending
    
    if ($datasetRefreshes) {
        $lastRefresh = $datasetRefreshes | Select-Object -First 1
        
        $totalAttempts = $datasetRefreshes.Count
        $successfulRefreshes = ($datasetRefreshes | Where-Object { $_.Status -eq "Completed" }).Count
        $failedRefreshes = ($datasetRefreshes | Where-Object { $_.Status -eq "Failed" }).Count
        $successRate = if ($totalAttempts -gt 0) { 
            [Math]::Round(($successfulRefreshes / $totalAttempts) * 100, 1) 
        } else { 
            0 
        }
        
        $successfulDurations = $datasetRefreshes | Where-Object { 
            $_.Status -eq "Completed" -and $_.StartTime -and $_.EndTime 
        } | ForEach-Object {
            ([DateTime]$_.EndTime - [DateTime]$_.StartTime).TotalMinutes
        }
        
        $avgDuration = if ($successfulDurations) {
            [Math]::Round(($successfulDurations | Measure-Object -Average).Average, 1)
        } else {
            0
        }
        
        $lastSuccessful = $datasetRefreshes | Where-Object { $_.Status -eq "Completed" } | 
            Select-Object -First 1
        
        $daysSinceSuccess = if ($lastSuccessful) {
            [Math]::Round(((Get-Date) - [DateTime]$lastSuccessful.EndTime).TotalDays, 1)
        } else {
            999
        }
        
        $healthStatus = if ($successRate -eq 0) {
            "Critical"
        } elseif ($successRate -lt 50) {
            "Poor"
        } elseif ($successRate -lt 80) {
            "Fair"
        } elseif ($daysSinceSuccess -gt 7) {
            "Stale"
        } else {
            "Healthy"
        }
        
        # Get gateway info for this dataset
        $datasetGateway = $datasources | Where-Object { $_.DatasetId -eq $dataset.DatasetId -and $_.GatewayId } | 
            Select-Object -First 1
        
        $gatewayName = if ($datasetGateway) {
            ($gatewayDetails | Where-Object { $_.GatewayId -eq $datasetGateway.GatewayId }).GatewayName
        } else {
            ""
        }
        
        [PSCustomObject]@{
            WorkspaceId             = $dataset.WorkspaceId
            WorkspaceName           = $dataset.WorkspaceName
            CapacityId              = ($workspaces | Where-Object { $_.WorkspaceId -eq $dataset.WorkspaceId }).CapacityId
            DatasetId               = $dataset.DatasetId
            DatasetName             = $dataset.DatasetName
            Endorsement             = $dataset.Endorsement
            Owner                   = $dataset.ConfiguredBy
            LastRefreshStatus       = $lastRefresh.Status
            LastRefreshTime         = $lastRefresh.EndTime
            DaysSinceLastSuccess    = $daysSinceSuccess
            TotalRefreshAttempts    = $totalAttempts
            SuccessfulRefreshes     = $successfulRefreshes
            FailedRefreshes         = $failedRefreshes
            SuccessRate             = $successRate
            AvgRefreshDuration      = $avgDuration
            HealthStatus            = $healthStatus
            IsGatewayRequired       = $dataset.IsOnPremGatewayRequired
            GatewayName             = $gatewayName
        }
    }
    else {
        $gatewayName = if ($dataset.IsOnPremGatewayRequired -eq "True") {
            $datasetGateway = $datasources | Where-Object { $_.DatasetId -eq $dataset.DatasetId -and $_.GatewayId } | 
                Select-Object -First 1
            if ($datasetGateway) {
                ($gatewayDetails | Where-Object { $_.GatewayId -eq $datasetGateway.GatewayId }).GatewayName
            } else { "" }
        } else { "" }
        
        [PSCustomObject]@{
            WorkspaceId             = $dataset.WorkspaceId
            WorkspaceName           = $dataset.WorkspaceName
            CapacityId              = ($workspaces | Where-Object { $_.WorkspaceId -eq $dataset.WorkspaceId }).CapacityId
            DatasetId               = $dataset.DatasetId
            DatasetName             = $dataset.DatasetName
            Endorsement             = $dataset.Endorsement
            Owner                   = $dataset.ConfiguredBy
            LastRefreshStatus       = "Never Refreshed"
            LastRefreshTime         = $null
            DaysSinceLastSuccess    = 999
            TotalRefreshAttempts    = 0
            SuccessfulRefreshes     = 0
            FailedRefreshes         = 0
            SuccessRate             = 0
            AvgRefreshDuration      = 0
            HealthStatus            = "Never Refreshed"
            IsGatewayRequired       = $dataset.IsOnPremGatewayRequired
            GatewayName             = $gatewayName
        }
    }
}

$datasetHealth | Export-Csv "DatasetHealth_$timestamp.csv" -NoTypeInformation

Write-Host "✓ Dataset health analysis complete`n" -ForegroundColor Green

# ============================================================================
# 7. IDENTIFY CRITICAL ISSUES
# ============================================================================

$criticalDatasets = $datasetHealth | Where-Object { $_.HealthStatus -eq "Critical" }
$poorHealthDatasets = $datasetHealth | Where-Object { $_.HealthStatus -eq "Poor" }
$staleDatasets = $datasetHealth | Where-Object { $_.HealthStatus -eq "Stale" }
$neverRefreshed = $datasetHealth | Where-Object { $_.HealthStatus -eq "Never Refreshed" }
$healthyDatasets = $datasetHealth | Where-Object { $_.HealthStatus -eq "Healthy" }

$recentFailures = $refreshHistory | Where-Object { 
    $_.Status -eq "Failed" -and 
    [DateTime]$_.StartTime -gt (Get-Date).AddDays(-1) 
}

$longRunningThreshold = 60
$longRunning = $datasetHealth | Where-Object { $_.AvgRefreshDuration -gt $longRunningThreshold }

# Gateway issues
$problematicGateways = $gatewayUsage | Where-Object { 
    $_.HealthStatus -eq "Poor" -or $_.GatewayStatus -ne "Live"
}

# Capacity stress
$stressedCapacities = $capacityUsage | Where-Object { $_.StressLevel -eq "High" }

# ============================================================================
# 8. BUILD COMPREHENSIVE HTML REPORT
# ============================================================================

Write-Host "Building HTML report..." -ForegroundColor Yellow

$sb = New-Object System.Text.StringBuilder

# Summary metrics
$healthMetricsHtml = @"
<div class='metric-card' style='border-left-color: #28a745;'>
    <h4>Healthy Datasets</h4>
    <div class='value' style='color: #28a745;'>$($healthyDatasets.Count)</div>
    <div class='subvalue'>$([Math]::Round(($healthyDatasets.Count / $datasetHealth.Count) * 100, 1))% functioning well</div>
</div>
<div class='metric-card' style='border-left-color: #dc3545;'>
    <h4>Critical Issues</h4>
    <div class='value' style='color: #dc3545;'>$($criticalDatasets.Count)</div>
    <div class='subvalue'>100% failure rate</div>
</div>
<div class='metric-card' style='border-left-color: #ffc107;'>
    <h4>Gateways</h4>
    <div class='value' style='color: $(if ($problematicGateways) { "#ffc107" } else { "#28a745" })'>$($gateways.Count)</div>
    <div class='subvalue'>$(if ($problematicGateways) { "$($problematicGateways.Count) with issues" } else { "All healthy" })</div>
</div>
<div class='metric-card' style='border-left-color: #17a2b8;'>
    <h4>Capacities</h4>
    <div class='value' style='color: $(if ($stressedCapacities) { "#ffc107" } else { "#28a745" })'>$($capacities.Count)</div>
    <div class='subvalue'>$(if ($stressedCapacities) { "$($stressedCapacities.Count) under stress" } else { "Load normal" })</div>
</div>
"@

# Gateway health table
$gatewayHealthHtml = if ($gatewayUsage) {
    "<table><thead><tr><th>Gateway Name</th><th>Status</th><th>Datasets</th><th>Refresh Attempts</th><th>Failure Rate</th><th>Health</th></tr></thead><tbody>" +
    (($gatewayUsage | ForEach-Object {
        $healthBadge = switch ($_.HealthStatus) {
            "Good" { "<span style='background:#28a745;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>✓ Good</span>" }
            "Fair" { "<span style='background:#ffc107;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>⚠️ Fair</span>" }
            "Poor" { "<span style='background:#dc3545;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>🚨 Poor</span>" }
        }
        
        $statusBadge = if ($_.GatewayStatus -eq "Live") {
            "<span style='background:#28a745;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>Live</span>"
        } else {
            "<span style='background:#dc3545;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>$($_.GatewayStatus)</span>"
        }
        
        "<tr>
            <td><strong>$($_.GatewayName)</strong></td>
            <td style='text-align:center;'>$statusBadge</td>
            <td style='text-align:center;'>$($_.DatasetCount)</td>
            <td style='text-align:center;'>$($_.RefreshAttempts)</td>
            <td style='text-align:center;'>$($_.FailureRate)%</td>
            <td style='text-align:center;'>$healthBadge</td>
        </tr>"
    }) -join "`n") +
    "</tbody></table>"
} else {
    "<div class='alert alert-info'><strong>ℹ️ No gateways configured or accessible</strong></div>"
}

# Capacity stress table
$capacityStressHtml = if ($capacityUsage) {
    "<table><thead><tr><th>Capacity</th><th>SKU</th><th>Workspaces</th><th>Datasets</th><th>Refreshes/Day</th><th>Max Concurrent</th><th>Stress Level</th></tr></thead><tbody>" +
    (($capacityUsage | ForEach-Object {
        $stressBadge = switch ($_.StressLevel) {
            "Low" { "<span style='background:#28a745;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>✓ Low</span>" }
            "Medium" { "<span style='background:#ffc107;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>⚠️ Medium</span>" }
            "High" { "<span style='background:#dc3545;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>🚨 High</span>" }
            "N/A" { "<span style='background:#6c757d;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>N/A</span>" }
        }
        
        "<tr>
            <td><strong>$($_.CapacityName)</strong></td>
            <td>$($_.CapacitySku)</td>
            <td style='text-align:center;'>$($_.WorkspaceCount)</td>
            <td style='text-align:center;'>$($_.RefreshableDatasets)</td>
            <td style='text-align:center;'>$($_.RefreshesPerDay)</td>
            <td style='text-align:center;'>$($_.MaxConcurrentRefreshes)</td>
            <td style='text-align:center;'>$stressBadge</td>
        </tr>"
    }) -join "`n") +
    "</tbody></table>"
} else {
    "<div class='alert alert-info'><strong>ℹ️ No Premium/Fabric capacities found</strong></div>"
}

# Critical datasets table
$criticalDatasetsHtml = if ($criticalDatasets) {
    "<table><thead><tr><th>Workspace</th><th>Dataset</th><th>Gateway</th><th>Owner</th><th>Last Status</th><th>Days Since Success</th></tr></thead><tbody>" +
    (($criticalDatasets | Select-Object -First 20 | ForEach-Object {
        "<tr>
            <td>$($_.WorkspaceName)</td>
            <td><strong>$($_.DatasetName)</strong></td>
            <td style='font-size:0.85em;'>$(if ($_.GatewayName) { $_.GatewayName } else { "—" })</td>
            <td>$($_.Owner)</td>
            <td><span style='background:#dc3545;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>$($_.LastRefreshStatus)</span></td>
            <td style='text-align:center;'>$($_.DaysSinceLastSuccess)</td>
        </tr>"
    }) -join "`n") +
    "</tbody></table>"
} else {
    "<div class='alert alert-success'><strong>✓ No critical datasets</strong></div>"
}

# Recent failures
$recentFailuresHtml = if ($recentFailures) {
    "<table><thead><tr><th>Time</th><th>Dataset</th><th>Workspace</th><th>Error</th></tr></thead><tbody>" +
    (($recentFailures | Select-Object -First 10 | ForEach-Object {
        $errorMsg = if ($_.ServiceExceptionJson) {
            try {
                $Myerror = $_.ServiceExceptionJson | ConvertFrom-Json
                $Myerror.errorDescription
            } catch {
                "Error parsing exception"
            }
        } else {
            "No error details"
        }
        
        "<tr>
            <td>$([DateTime]$_.StartTime | Get-Date -Format 'yyyy-MM-dd HH:mm')</td>
            <td>$($_.DatasetName)</td>
            <td>$($_.WorkspaceName)</td>
            <td style='font-size:0.85em;color:#666;max-width:400px;'>$errorMsg</td>
        </tr>"
    }) -join "`n") +
    "</tbody></table>"
} else {
    "<div class='alert alert-success'><strong>✓ No failures in last 24 hours</strong></div>"
}

# All datasets health table
$allHealthTableHtml = ($datasetHealth | Sort-Object HealthStatus, DatasetName | ForEach-Object {
    $healthBadge = switch ($_.HealthStatus) {
        "Healthy" { "<span style='background:#28a745;colorContinue14:28:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>✓ Healthy</span>" }
"Critical" { "<span style='background:#dc3545;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>🚨 Critical</span>" }
"Poor" { "<span style='background:#ffc107;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>⚠️ Poor</span>" }
"Fair" { "<span style='background:#17a2b8;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>ℹ️ Fair</span>" }
"Stale" { "<span style='background:#17a2b8;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>ℹ️ Stale</span>" }
"Never Refreshed" { "<span style='background:#6c757d;color:white;padding:3px 8px;border-radius:10px;font-size:0.75em;'>Never Refreshed</span>" }
default { $_.HealthStatus }
}
$lastRefreshDisplay = if ($_.LastRefreshTime) {
    [DateTime]$_.LastRefreshTime | Get-Date -Format 'yyyy-MM-dd HH:mm'
} else {
    "—"
}

"<tr>
    <td>$($_.WorkspaceName)</td>
    <td><strong>$($_.DatasetName)</strong></td>
    <td style='font-size:0.85em;'>$(if ($_.GatewayName) { $_.GatewayName } else { "—" })</td>
    <td style='text-align:center;'>$healthBadge</td>
    <td style='text-align:center;'>$($_.SuccessRate)%</td>
    <td>$lastRefreshDisplay</td>
    <td style='text-align:center;'>$($_.AvgRefreshDuration) min</td>
</tr>"
}) -join "`n"
[void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power BI Health Report - $reportDate</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; color: #333; line-height: 1.6; }
        .container { max-width: 1800px; margin: 0 auto; padding: 20px; }
        header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; border-radius: 10px; margin-bottom: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        header h1 { font-size: 2.5em; margin-bottom: 10px; }
        header p { font-size: 1.1em; opacity: 0.9; }
        .section { background: white; padding: 30px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .section h2 { color: #667eea; border-bottom: 3px solid #667eea; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.8em; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #667eea; color: white; padding: 12px; text-align: left; font-weight: 600; }
        td { padding: 10px 12px; border-bottom: 1px solid #e0e0e0; }
        tr:hover { background: #f8f9fa; }
        .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin: 20px 0; }
        .metric-card { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 20px; border-radius: 8px; border-left: 4px solid #667eea; }
        .metric-card h4 { color: #555; font-size: 0.9em; margin-bottom: 10px; text-transform: uppercase; }
        .metric-card .value { font-size: 2em; font-weight: bold; color: #333; }
        .metric-card .subvalue { color: #666; font-size: 0.85em; margin-top: 5px; }
        .alert { padding: 15px 20px; border-radius: 8px; margin: 15px 0; border-left: 4px solid; }
        .alert-danger { background: #f8d7da; border-color: #dc3545; color: #721c24; }
        .alert-success { background: #d4edda; border-color: #28a745; color: #155724; }
        .alert-info { background: #d1ecf1; border-color: #17a2b8; color: #0c5460; }
        footer { text-align: center; padding: 30px; color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>💚 Power BI Health & Infrastructure Report</h1>
            <p>Generated: $reportDate</p>
            <p>Datasets: $($datasetHealth.Count) | Gateways: $($gateways.Count) | Capacities: $($capacities.Count + 1)</p>
        </header>
    <div class="section">
        <h2>📊 Health Overview</h2>
        <div class="metric-grid">
            $healthMetricsHtml
        </div>
    </div>
    
    <div class="section">
        <h2>🌉 Gateway Health</h2>
        <p>On-premises data gateway status and performance.</p>
        $gatewayHealthHtml
    </div>
    
    <div class="section">
        <h2>⚡ Capacity Stress Analysis</h2>
        <p>Premium/Fabric capacity usage and concurrent refresh load.</p>
        $capacityStressHtml
    </div>
    
    <div class="section">
        <h2>🚨 Critical Datasets (Immediate Attention Required)</h2>
        <p>Datasets with 100% failure rate or never successfully refreshed.</p>
        $criticalDatasetsHtml
    </div>
    
    <div class="section">
        <h2>⏱️ Recent Failures (Last 24 Hours)</h2>
        $recentFailuresHtml
    </div>
    
    <div class="section">
        <h2>📋 All Refreshable Datasets</h2>
        <table>
            <thead>
                <tr>
                    <th>Workspace</th>
                    <th>Dataset</th>
                    <th>Gateway</th>
                    <th style="text-align:center;">Health Status</th>
                    <th style="text-align:center;">Success Rate</th>
                    <th>Last Refresh</th>
                    <th style="text-align:center;">Avg Duration</th>
                </tr>
            </thead>
            <tbody>
                $allHealthTableHtml
            </tbody>
        </table>
    </div>
    
    <footer>
        <p>Power BI Health & Infrastructure Report | Generated: $reportDate</p>
    </footer>
</div>
</body>
</html>
"@)
$htmlPath = "PowerBI_Health_Report_$timestamp.html"
$sb.ToString() | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "n========================================" -ForegroundColor Green
Write-host "✓ HEALTH ANALYSIS COMPLETE" -ForegroundColor Green
Write-host "========================================" -ForegroundColor Green
Write-host "Datasets Analyzed: $($datasetHealth.Count)" -ForegroundColor White
Write-host "  Healthy: $($healthyDatasets.Count)" -ForegroundColor Green
Write-host "  Critical: $($criticalDatasets.Count)" -ForegroundColor Red
Write-host "  Poor: $($poorHealthDatasets.Count)" -ForegroundColor Yellow
Write-host "Gateways: $($gateways.Count)" -ForegroundColor White
Write-host "  Problematic: $($problematicGateways.Count)" -ForegroundColor $(if ($problematicGateways) { "Red" } else { "Green" })
Write-host "Capacities: $($capacities.Count)" -ForegroundColor White
Write-host "  Under Stress: $($stressedCapacities.Count)" -ForegroundColor $(if ($stressedCapacities) { "Yellow" } else { "Green" })
Write-host "nFiles created:" -ForegroundColor Cyan
Write-Host "  RefreshHistory_Raw_$timestamp.csv" -ForegroundColor White
Write-Host "  DatasetHealth_$timestamp.csv" -ForegroundColor White
Write-Host "  Gateways_$timestamp.csv" -ForegroundColor White
Write-Host "  GatewayHealth_$timestamp.csv" -ForegroundColor White
Write-Host "  Capacities_$timestamp.csv" -ForegroundColor White
Write-Host "  CapacityUsage_$timestamp.csv" -ForegroundColor White
Write-Host "  $htmlPath" -ForegroundColor White
Start-Process $htmlPath
Write-Host "`n✓ Done! Ready for usage analysis next." -ForegroundColor Green
