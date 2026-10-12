<#
.SYNOPSIS
    Azure Utilization, Performance Monitoring & Cost Optimization Assessment (v2) - runs Phase 1 + Phase 2 sequentially.

.DESCRIPTION
    Wrapper that runs both phases end to end into a single timestamped folder:
      - Phase 1: Phase1-CollectUtilization.ps1  (read-only collection: every resource with ALL properties and tags,
                 hourly Azure Monitor utilization via the Metrics Batch API, auto-discovered metrics for every other
                 billed resource type, Log Analytics / Sentinel daily ingestion per workspace and table, cost history,
                 Advisor cost + performance, orphaned resources, retail / reservation / savings plan pricing)
      - Phase 2: Phase2-AnalyzeAndReport.ps1    (provisioned vs. utilized analysis, utilization % of the current SKU,
                 right-sizing, rule-based suggested recommendations with criteria, efficiency score, savings,
                 master Excel workbook, one workbook per resource type, interactive HTML dashboard)

    Purpose: validate - with measured average / P95 / peak utilization and usage patterns - whether everything that is
    provisioned and paid for is actually being consumed, and where right-sizing or commitment changes apply
    (e.g. as a follow-up to a previous cost optimization exercise).

.PARAMETER CustomerName
    Required. Customer name used in report titles and file names.

.PARAMETER TenantId
    Optional. Tenant to assess (defaults to the current Az PowerShell / Azure CLI tenant).

.PARAMETER SubscriptionIds
    Optional. Specific subscription IDs to assess (default: all enabled subscriptions in the tenant).

.PARAMETER LookbackDays
    Utilization window in days (7-93). Default 30.

.PARAMETER CostType
    AmortizedCost (default) or ActualCost.

.PARAMETER CostHistoryMonths
    Months of monthly cost history (trend / impact of previous optimization). Default 6.

.PARAMETER TimeZoneId / WeekendDays / BusinessHoursStart / BusinessHoursEnd
    Usage-pattern settings. Defaults: Arab Standard Time (UTC+3), Friday & Saturday weekend, 07:00-18:00.

.PARAMETER ConfigPath
    Optional JSON configuration (default: CostOptimization.Config.json next to the scripts) - thresholds, property / tag
    column limits, auto-discovered metric settings, excluded types, metric overrides, environment patterns.

.PARAMETER CpuIdleThreshold / CpuLowThreshold / CpuHighThreshold / MemLowThreshold / MemHighThreshold / TargetUtilization
    Right-sizing thresholds passed to Phase 2 only when specified (otherwise the configuration file / defaults apply:
    5 / 40 / 85 / 60 / 90 / 70).

.PARAMETER IncludeResourceTypes
    Optional. Limit the dedicated analysis to specific catalog keys (e.g. VirtualMachines, SqlDatabases).

.PARAMETER ThrottleLimit
    Parallel Azure Monitor requests (default 12).

.PARAMETER MetricsBatchSize
    Resources per Azure Monitor Metrics Batch API call (default 20, max 50).

.PARAMETER NoBatchMetrics
    Use per-resource Azure Monitor calls instead of the Metrics Batch API.

.PARAMETER SkipCostData / SkipLogAnalytics / SkipPricing / SkipGenericTypes
    Skip Cost Management, Log Analytics, retail pricing or auto-discovered (generic) resource types.

.PARAMETER OutputRoot
    Folder where the timestamped assessment folder is created (default: this script's folder).

.PARAMETER NoOpen
    Do not open the HTML dashboard at the end.

.EXAMPLE
    .\Start-AzureCostOptimizationAssessment.ps1 -CustomerName "Contoso"

.EXAMPLE
    .\Start-AzureCostOptimizationAssessment.ps1 -CustomerName "Contoso" -SubscriptionIds @("sub-id-1","sub-id-2") -LookbackDays 60 -CostHistoryMonths 12

.NOTES
    Author  : Microsoft Solution Engineering Team
    Version : 2.0
    Requires: PowerShell 7+, Az.Accounts (or Azure CLI sign-in), ImportExcel. Reader + Monitoring Reader + Log Analytics Reader
              (+ Cost Management Reader for cost data) on the subscriptions in scope. Read-only - no changes are made.

    Output (CostOptimizationAssessment_<timestamp>):
      01_RawData\                                         Phase 1 raw exports (CSV) - input for Phase 2
        ResourceTypes\<Type>.csv                          Every property (Prop.*), every tag (Tag.*) and utilization statistics per resource
        DailyUtilization\<Type>_Daily.csv                 Daily average / max / min / total per metric and resource
        UsagePatterns\<Type>_HourOfWeek.csv               Average utilization by local day-of-week and hour
        LogAnalytics_DailyIngestion.csv / _DailyByTable   Daily billable GB per workspace (and table) vs. LA / Sentinel commitment
        Cost_*.csv (incl. monthly trend), Advisor_*.csv, Orphaned_Resources.csv, Pricing_*.csv, Reference_*.csv, Collection_*
      <Customer> - Utilization & Cost Optimization Report.xlsx    Master workbook (dashboard, suggested recommendations,
                                                                  action tracker, scorecards, commitments, cost trend,
                                                                  tags, usage patterns, one sheet per resource type, rules)
      ResourceTypes\<Type>.xlsx                           One workbook per resource type
      <Customer> - Utilization & Optimization Dashboard.html     Interactive dashboard (offline, single file)
      Recommendations.csv, Assessment_Log.txt
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomerName,

    [string]$TenantId,

    [string[]]$SubscriptionIds,

    [ValidateRange(7, 93)]
    [int]$LookbackDays = 30,

    [ValidateSet('AmortizedCost', 'ActualCost')]
    [string]$CostType = 'AmortizedCost',

    [ValidateRange(0, 36)]
    [int]$CostHistoryMonths = 6,

    [string]$TimeZoneId = 'Arab Standard Time',

    [ValidateSet('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')]
    [string[]]$WeekendDays = @('Friday', 'Saturday'),

    [ValidateRange(0, 23)]
    [int]$BusinessHoursStart = 7,

    [ValidateRange(1, 24)]
    [int]$BusinessHoursEnd = 18,

    [string]$ConfigPath,

    [double]$CpuIdleThreshold,
    [double]$CpuLowThreshold,
    [double]$CpuHighThreshold,
    [double]$MemLowThreshold,
    [double]$MemHighThreshold,
    [double]$TargetUtilization,

    [string[]]$IncludeResourceTypes,

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 12,

    [ValidateRange(1, 50)]
    [int]$MetricsBatchSize = 20,

    [switch]$NoBatchMetrics,
    [switch]$SkipCostData,
    [switch]$SkipLogAnalytics,
    [switch]$SkipPricing,
    [switch]$SkipGenericTypes,

    [string]$OutputRoot = $PSScriptRoot,

    [switch]$NoOpen
)

# --- Bypass execution policy for this process so downloaded/unsigned scripts can run ---
$currentPolicy = Get-ExecutionPolicy -Scope Process
if ($currentPolicy -ne 'Bypass' -and $currentPolicy -ne 'Unrestricted') {
    try { Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force; Write-Host '[INFO] Execution policy set to Bypass for this session.' -ForegroundColor Yellow }
    catch { Write-Warning "Could not set execution policy. Run: Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process" }
}

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "`nThis assessment requires PowerShell 7 or later (parallel metric collection)." -ForegroundColor Red
    Write-Host "Install it with:  winget install --id Microsoft.PowerShell --source winget" -ForegroundColor Yellow
    Write-Host "Then run this script from 'pwsh'.`n" -ForegroundColor Yellow
    exit 1
}

try { Clear-Host } catch { Write-Verbose 'Non-interactive host - screen not cleared.' }
Write-Host ''
Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '  AZURE UTILIZATION, PERFORMANCE MONITORING & COST OPTIMIZATION  (v2)' -ForegroundColor Yellow
Write-Host '  Phase 1: Collect utilization  →  Phase 2: Analyze & report' -ForegroundColor Yellow
Write-Host '  Provisioned vs. utilized · usage patterns · right-sizing · savings' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''
Write-Host '  This script will automatically:' -ForegroundColor White
Write-Host '  1. Inventory every resource with all properties & tags; collect hourly utilization (avg / P95 / peak)' -ForegroundColor Gray
Write-Host '  2. Auto-discover metrics for every other billed resource type (IaaS, PaaS, consumption services)' -ForegroundColor Gray
Write-Host '  3. Measure Log Analytics & Sentinel daily ingestion (per workspace and table) vs. commitment tiers' -ForegroundColor Gray
Write-Host '  4. Collect cost history, Advisor (cost + performance), orphaned resources and retail / RI / SP pricing' -ForegroundColor Gray
Write-Host '  5. Analyze utilization % vs. current SKU, right-size, score efficiency and suggest rule-based actions' -ForegroundColor Gray
Write-Host '  6. Export the master workbook, per-type workbooks and the interactive HTML dashboard' -ForegroundColor Gray
Write-Host ''
Write-Host '  Assessment for: ' -NoNewline -ForegroundColor White; Write-Host $CustomerName -ForegroundColor Cyan
Write-Host "  Window        : last $LookbackDays full days ($TimeZoneId, weekend $($WeekendDays -join '/'), business hours $($BusinessHoursStart):00-$($BusinessHoursEnd):00)" -ForegroundColor White
Write-Host ''

$phase1Script = Join-Path $PSScriptRoot 'Phase1-CollectUtilization.ps1'
$phase2Script = Join-Path $PSScriptRoot 'Phase2-AnalyzeAndReport.ps1'
foreach ($s in @($phase1Script, $phase2Script)) { if (-not (Test-Path $s)) { Write-Host "ERROR: Script not found: $s" -ForegroundColor Red; exit 1 } }

################################################################################
# PREREQUISITES: MODULES & SIGN-IN
################################################################################

foreach ($m in @('Az.Accounts', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing module $m (current user)..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}
Import-Module Az.Accounts -WarningAction SilentlyContinue | Out-Null

$authOk = $false
$ctx = Get-AzContext -ErrorAction SilentlyContinue
if ($ctx) {
    try {
        $tokParams = @{ ResourceUrl = 'https://management.azure.com/'; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
        if ($TenantId) { $tokParams.TenantId = $TenantId }
        $null = Get-AzAccessToken @tokParams
        $authOk = $true
        Write-Host "Signed in (Az PowerShell): $($ctx.Account.Id)" -ForegroundColor Green
    }
    catch { Write-Host 'Az PowerShell session found but the token could not be refreshed (MFA / expiry).' -ForegroundColor Yellow }
}
if (-not $authOk -and (Get-Command az -ErrorAction SilentlyContinue)) {
    $cliAccount = & az account show -o json 2>$null | Out-String
    if ($LASTEXITCODE -eq 0 -and $cliAccount) {
        $cli = $cliAccount | ConvertFrom-Json
        if (-not $TenantId -or $cli.tenantId -eq $TenantId) { $authOk = $true; Write-Host "Signed in (Azure CLI): $($cli.user.name) - tenant $($cli.tenantId)" -ForegroundColor Green }
    }
}
if (-not $authOk) {
    Write-Host 'Signing in to Azure...' -ForegroundColor Yellow
    $connect = @{ ErrorAction = 'Stop' }
    if ($TenantId) { $connect.TenantId = $TenantId }
    Connect-AzAccount @connect | Out-Null
    Write-Host "Signed in: $((Get-AzContext).Account.Id)" -ForegroundColor Green
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputPath = Join-Path $OutputRoot "CostOptimizationAssessment_$timestamp"
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
Write-Host "Output directory: $outputPath`n" -ForegroundColor Green

################################################################################
# PHASE 1: DATA COLLECTION
################################################################################

Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Magenta
Write-Host '  PHASE 1: Utilization, ingestion, cost & inventory collection' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Magenta
try {
    $p1 = @{
        OutputPath = $outputPath; LookbackDays = $LookbackDays; ThrottleLimit = $ThrottleLimit; CostType = $CostType; CostHistoryMonths = $CostHistoryMonths
        TimeZoneId = $TimeZoneId; WeekendDays = $WeekendDays; BusinessHoursStart = $BusinessHoursStart; BusinessHoursEnd = $BusinessHoursEnd; MetricsBatchSize = $MetricsBatchSize
    }
    if ($TenantId) { $p1.TenantId = $TenantId }
    if ($SubscriptionIds) { $p1.SubscriptionIds = $SubscriptionIds }
    if ($ConfigPath) { $p1.ConfigPath = $ConfigPath }
    if ($IncludeResourceTypes) { $p1.IncludeResourceTypes = $IncludeResourceTypes }
    foreach ($sw in @('NoBatchMetrics', 'SkipCostData', 'SkipLogAnalytics', 'SkipPricing', 'SkipGenericTypes')) { if ($PSBoundParameters.ContainsKey($sw) -and $PSBoundParameters[$sw]) { $p1[$sw] = $true } }
    & $phase1Script @p1
    if (-not (Test-Path (Join-Path $outputPath '01_RawData\Collection_Metadata.json'))) { throw 'Phase 1 did not produce 01_RawData\Collection_Metadata.json' }
    Write-Host "`n  ✓✓✓ PHASE 1 COMPLETE ✓✓✓ - raw data in 01_RawData. Proceeding to Phase 2...`n" -ForegroundColor Green
}
catch {
    Write-Host "`n  ✗✗✗ PHASE 1 FAILED ✗✗✗" -ForegroundColor Red
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host '  Assessment halted. Phase 2 will not run.' -ForegroundColor Yellow
    exit 1
}

################################################################################
# PHASE 2: ANALYSIS & REPORTING
################################################################################

Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Magenta
Write-Host '  PHASE 2: Provisioned vs. utilized analysis, savings & reports' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Magenta
try {
    $p2 = @{ InputPath = $outputPath; CustomerName = $CustomerName }
    if ($ConfigPath) { $p2.ConfigPath = $ConfigPath }
    foreach ($t in @('CpuIdleThreshold', 'CpuLowThreshold', 'CpuHighThreshold', 'MemLowThreshold', 'MemHighThreshold', 'TargetUtilization')) { if ($PSBoundParameters.ContainsKey($t)) { $p2[$t] = $PSBoundParameters[$t] } }
    if ($NoOpen) { $p2.NoOpen = $true }
    & $phase2Script @p2
    Write-Host "`n  ✓✓✓ PHASE 2 COMPLETE ✓✓✓`n" -ForegroundColor Green
}
catch {
    Write-Host "`n  ✗✗✗ PHASE 2 FAILED ✗✗✗" -ForegroundColor Red
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host '  Phase 1 data is intact. Re-run Phase 2 only (no Azure calls needed):' -ForegroundColor Yellow
    Write-Host "  .\Phase2-AnalyzeAndReport.ps1 -InputPath `"$outputPath`" -CustomerName `"$CustomerName`"" -ForegroundColor Cyan
    exit 1
}

################################################################################
# FINAL SUMMARY
################################################################################

$xlsx = Get-ChildItem -LiteralPath $outputPath -Filter '*Utilization & Cost Optimization Report.xlsx' | Select-Object -First 1
$html = Get-ChildItem -LiteralPath $outputPath -Filter '*Utilization & Optimization Dashboard.html' | Select-Object -First 1
$typeFiles = @(Get-ChildItem -LiteralPath (Join-Path $outputPath 'ResourceTypes') -Filter '*.xlsx' -ErrorAction SilentlyContinue)

Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Green
Write-Host '  UTILIZATION & COST OPTIMIZATION ASSESSMENT FINISHED SUCCESSFULLY!' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════════════════' -ForegroundColor Green
Write-Host "Customer        : $CustomerName" -ForegroundColor White
Write-Host "Output location : $outputPath" -ForegroundColor White
Write-Host ''
Write-Host '🎯 Deliverables:' -ForegroundColor Green
if ($xlsx) {
    Write-Host "   ✓ $($xlsx.Name)" -ForegroundColor White
    Write-Host '       Executive Overview (what is collected & coverage), Executive Dashboard (charts), Suggested Recommendations (Rule ID + criteria), Action Tracker,' -ForegroundColor Gray
    Write-Host '       Quick Wins, Efficiency Scorecard, Commitments, LA / Sentinel ingestion, Cost Trend, Tags, Usage Patterns,' -ForegroundColor Gray
    Write-Host '       one sheet per resource type (all properties + tags + utilization), Recommendation Rules, Data Dictionary' -ForegroundColor Gray
}
Write-Host "   ✓ ResourceTypes\*.xlsx ($($typeFiles.Count) per-type workbooks: resources, recommendations, metric statistics, daily pivots, hour-of-week)" -ForegroundColor White
if ($html) { Write-Host "   ✓ $($html.Name)" -ForegroundColor White; Write-Host '       Interactive dashboard: global filters, service dashboards, usage patterns, commitments, cost, tags, explorer' -ForegroundColor Gray }
Write-Host '   ✓ Recommendations.csv and 01_RawData\ (raw CSV exports, daily utilization per resource type)' -ForegroundColor White
Write-Host ''
Write-Host "Assessment complete! Review the outputs in: $outputPath" -ForegroundColor Green
