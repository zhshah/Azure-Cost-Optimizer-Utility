<#
.SYNOPSIS
    Phase 2 - Azure Utilization, Performance & Cost Optimization: ANALYSIS & REPORTING (v2).

.DESCRIPTION
    Consumes the raw exports produced by Phase1-CollectUtilization.ps1 (<InputPath>\01_RawData) and produces:

      <Customer> - Utilization & Cost Optimization Report.xlsx   Master workbook:
            Executive Dashboard (charts), Suggested Recommendations (rule-based, with Rule ID + criteria + evidence),
            Action Tracker, Quick Wins / Rightsize / Delete / Configure / Performance Risk, Efficiency Scorecard,
            Commitment Utilization (Log Analytics / Sentinel / Reservations / Savings Plans / VM commitment candidates),
            Cost Trend (monthly history), Daily Cost, Cost by Tag, Tag Coverage, Usage Patterns, Daily Utilization Trend,
            Log Analytics daily ingestion (per workspace & table), ONE SHEET PER RESOURCE TYPE (every property, every tag,
            utilization statistics vs. current SKU, suggested recommendation), Orphaned Resources, Other Paid Resources,
            Advisor (Cost + Performance), Inventory, Recommendation Rules, Data Dictionary, Methodology, Scope.
      ResourceTypes\<Type>.xlsx    One workbook per resource type (resources + recommendations + daily utilization pivots
                                   + hour-of-week pattern)
      Recommendations.csv          Flat list of every suggested recommendation
      <Customer> - Utilization & Optimization Dashboard.html   Interactive single-file dashboard (global filters,
                                   service dashboards, usage patterns, commitments, cost trends, tags, resource explorer)

    Phase 2 makes NO Azure calls - it can be re-run offline with different thresholds.

.PARAMETER InputPath
    Assessment folder created by Phase 1 (contains 01_RawData).

.PARAMETER CustomerName
    Customer name used in report titles and file names.

.PARAMETER OutputPath
    Output folder (defaults to InputPath).

.PARAMETER ConfigPath
    Optional JSON configuration (default: CostOptimization.Config.json next to this script; built-in defaults when absent). Explicit parameters win.

.PARAMETER CpuIdleThreshold / CpuLowThreshold / CpuHighThreshold / MemLowThreshold / MemHighThreshold
    Right-sizing thresholds (percent). Defaults: idle avg < 5, over-provisioned P95 < 40, under-provisioned P95 > 85,
    memory over-provisioned max < 60, memory pressure P95 > 90.

.PARAMETER TargetUtilization
    Target P95 CPU after a resize (default 70%).

.PARAMETER NoOpen
    Do not open the HTML report when finished.

.NOTES
    Requires: PowerShell 7+, ImportExcel module.
    Version : 2.0
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$CustomerName,
    [string]$OutputPath,
    [string]$ConfigPath,
    [double]$CpuIdleThreshold = 5,
    [double]$CpuLowThreshold = 40,
    [double]$CpuHighThreshold = 85,
    [double]$MemLowThreshold = 60,
    [double]$MemHighThreshold = 90,
    [double]$TargetUtilization = 70,
    [switch]$NoOpen
)

#region ============================== BOOTSTRAP ==============================

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
$script:RunStart = Get-Date
if (-not $OutputPath) { $OutputPath = $InputPath }
$RawPath = Join-Path $InputPath '01_RawData'
if (-not (Test-Path (Join-Path $RawPath 'Collection_Metadata.json'))) { throw "Phase 1 output not found: $RawPath\Collection_Metadata.json" }
$TypeOutPath = Join-Path $OutputPath 'ResourceTypes'
foreach ($p in @($OutputPath, $TypeOutPath)) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }
$script:LogFile = Join-Path $OutputPath 'Assessment_Log.txt'

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host 'ImportExcel module not found - installing for current user...' -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
}
Import-Module ImportExcel -ErrorAction Stop

function Write-Log {
    param([string]$Message, [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Section')][string]$Level = 'Info')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) { 'Info' { 'Gray' } 'Warning' { 'Yellow' } 'Error' { 'Red' } 'Success' { 'Green' } 'Section' { 'Cyan' } }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
    try { "[$ts] [$Level] $Message" | Out-File -FilePath $script:LogFile -Append -Encoding utf8 } catch { Write-Verbose $_ }
}

function N($v) {
    if ($null -eq $v) { return $null }
    $s = [string]$v
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $d = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}
function B($v) { return ([string]$v) -match '^(?i)true$' }
function Rnd($v, [int]$d = 1) { if ($null -eq $v) { return $null }; return [Math]::Round([double]$v, $d) }
function Fmt($v, [int]$d = 1, [string]$suffix = '') { if ($null -eq $v) { return 'n/a' }; return ('{0:N' + $d + '}{1}') -f [double]$v, $suffix }
function Money($v) { if ($null -eq $v) { return 'n/a' }; return ('${0:N0}' -f [double]$v) }

# ---- Configuration (explicit parameters win over the config file)
if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'CostOptimization.Config.json' }
$script:Config = $null
if (Test-Path $ConfigPath) { try { $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable } catch { Write-Warning "Config not parsed: $($_.Exception.Message)" } }
function Get-Cfg([string]$Path, $Default) {
    $node = $script:Config
    foreach ($p in ($Path -split '\.')) { if ($node -is [System.Collections.IDictionary] -and $node.Contains($p)) { $node = $node[$p] } else { return $Default } }
    if ($null -eq $node) { return $Default }; return $node
}
$cfgMap = @{ CpuIdleThreshold = 'Thresholds.CpuIdle'; CpuLowThreshold = 'Thresholds.CpuLow'; CpuHighThreshold = 'Thresholds.CpuHigh'; MemLowThreshold = 'Thresholds.MemLow'; MemHighThreshold = 'Thresholds.MemHigh'; TargetUtilization = 'Thresholds.TargetUtilization' }
foreach ($k in $cfgMap.Keys) { if (-not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value ([double](Get-Cfg $cfgMap[$k] (Get-Variable -Name $k -ValueOnly))) } }
$GenericLow = [double](Get-Cfg 'Thresholds.GenericLowUtilization' 30)
$GenericHigh = [double](Get-Cfg 'Thresholds.GenericHighUtilization' 85)
$SnapshotMaxAge = [double](Get-Cfg 'Thresholds.SnapshotMaxAgeDays' 90)
$SteadyHoursPct = [double](Get-Cfg 'Thresholds.SteadyStateHoursPct' 95)
$GenericMinRecCost = [double](Get-Cfg 'Thresholds.MinMonthlyCostForGenericRecommendation' 1)
$EnvProdTag = [string](Get-Cfg 'Environment.ProductionTagPattern' '(?i)^(prod|production|prd|live|prod-.*)$')
$EnvNonProdTag = [string](Get-Cfg 'Environment.NonProductionTagPattern' '(?i)(dev|test|uat|qa|stag|sandbox|sbx|poc|demo|non-?prod|pre-?prod|lab|train|trial)')
$EnvNonProdName = [string](Get-Cfg 'Environment.NonProductionNamePattern' '(?<![a-z])(dev|devtest|test|tst|uat|qa|sit|stg|stage|staging|sandbox|sbx|poc|demo|nonprod|non-prod|np|preprod|pre-prod|lab|training)(?![a-z])')
$EnvProdName = [string](Get-Cfg 'Environment.ProductionNamePattern' '(?<![a-z])(prod|prd|production|live)(?![a-z])')

function Import-RawCsv([string]$Name) {
    $p = Join-Path $RawPath $Name
    if (Test-Path $p) { return @(Import-Csv -Path $p -Encoding utf8) }
    return @()
}

#endregion

#region ============================== LOAD PHASE 1 DATA ==============================

Write-Log ('=' * 78) Section
Write-Log 'PHASE 2 - ANALYSIS: PROVISIONED vs UTILIZED, RIGHT-SIZING & SAVINGS' Section
Write-Log ('=' * 78) Section

$Meta = Get-Content (Join-Path $RawPath 'Collection_Metadata.json') -Raw | ConvertFrom-Json
$LookbackDays = [int]$Meta.LookbackDays
$MonthFactor = 30.4375 / [Math]::Max(1, $LookbackDays)
$TotalHours = [int]$Meta.TotalHours
$BizHoursPerWeek = (7 - @($Meta.WeekendDays).Count) * ([int]$Meta.BusinessHoursEnd - [int]$Meta.BusinessHoursStart)
Write-Log "Window: $($Meta.WindowStartLocal) -> $($Meta.WindowEndLocal) ($LookbackDays days, $($Meta.TimeZoneId)); subscriptions: $($Meta.SubscriptionCount)"

$Subscriptions = Import-RawCsv 'Subscriptions.csv'
$Inventory = Import-RawCsv 'Inventory_AllResources.csv'
$CollectionSummary = Import-RawCsv 'Collection_Summary.csv'
$Orphans = Import-RawCsv 'Orphaned_Resources.csv'
$Advisor = Import-RawCsv 'Advisor_CostRecommendations.csv'
$Defender = Import-RawCsv 'Defender_Plans.csv'
$Reservations = Import-RawCsv 'Commitments_Reservations.csv'
$SavingsPlans = Import-RawCsv 'Commitments_SavingsPlans.csv'
$LaDaily = Import-RawCsv 'LogAnalytics_DailyIngestion.csv'
$LaTopTables = Import-RawCsv 'LogAnalytics_TopTables.csv'
$CostByResource = Import-RawCsv 'Cost_ByResource.csv'
$CostDaily = Import-RawCsv 'Cost_DailyByService.csv'
$VmSizes = Import-RawCsv 'Reference_VmSizes.csv'
$PriceVm = Import-RawCsv 'Pricing_VirtualMachines.csv'
$PriceDisk = Import-RawCsv 'Pricing_ManagedDisks.csv'
$PriceApp = Import-RawCsv 'Pricing_AppService.csv'
$PriceLa = Import-RawCsv 'Pricing_LogAnalyticsSentinel.csv'
$PriceFiles = Import-RawCsv 'Pricing_Files.csv'
$PriceSql = Import-RawCsv 'Pricing_SqlVCore.csv'

# Per-type data
$TypeData = [ordered]@{}
foreach ($f in (Get-ChildItem -Path (Join-Path $RawPath 'ResourceTypes') -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)) {
    $TypeData[$f.BaseName] = @(Import-Csv -Path $f.FullName -Encoding utf8)
}
Write-Log "Loaded $($TypeData.Count) resource-type dataset(s), $($Inventory.Count) inventory rows, $($CostByResource.Count) cost rows"

# ---- Cost per resource (monthly run-rate)
$script:HasCost = $CostByResource.Count -gt 0
$script:Currency = 'USD'
$CostMap = @{}
if ($script:HasCost) {
    $useUsd = @($CostByResource | Where-Object { $_.CostUSD -ne '' -and $null -ne $_.CostUSD }).Count -gt 0
    if (-not $useUsd) { $script:Currency = [string](@($CostByResource | Where-Object Currency | Select-Object -First 1).Currency) }
    foreach ($c in $CostByResource) {
        $id = ([string]$c.ResourceId).ToLower()
        $v = if ($useUsd) { N $c.CostUSD } else { N $c.Cost }
        $CostMap[$id] = [double]($CostMap[$id] ?? 0) + [double]($v ?? 0)
    }
}
function Get-MonthlyCost([string]$Id) {
    if (-not $script:HasCost -or [string]::IsNullOrEmpty($Id)) { return $null }
    $v = $CostMap[$Id.ToLower()]
    if ($null -eq $v) { return 0.0 }
    return [Math]::Round([double]$v * $MonthFactor, 2)
}
$TotalMonthlyCost = if ($script:HasCost) { [Math]::Round((($CostMap.Values | Measure-Object -Sum).Sum) * $MonthFactor, 0) } else { $null }

#endregion

#region ============================== PRICING LOOKUPS ==============================

$VmPrice = @{}; $VmCommit = @{}
foreach ($p in $PriceVm) {
    $k = "$($p.Region)|$($p.VmSize)|$($p.OS)".ToLower()
    $VmPrice[$k] = N $p.HourlyPrice
    $VmCommit[$k] = @{ RI1Y = N $p.RI1YHourly; RI3Y = N $p.RI3YHourly; SP1Y = N $p.SP1YHourly; SP3Y = N $p.SP3YHourly }
}
function Get-VmHourly([string]$Region, [string]$Size, [string]$Os) {
    $os2 = if ($Os -match '(?i)windows') { 'Windows' } else { 'Linux' }
    return $VmPrice["$Region|$Size|$os2".ToLower()]
}

$DiskPrice = @{}
foreach ($p in $PriceDisk) { if ($p.MeterName -match '^([PES]\d+) (LRS|ZRS) Disk$') { $DiskPrice["$($p.Region)|$($Matches[1])|$($Matches[2])".ToLower()] = N $p.Price } }
function Get-DiskMonthly([string]$Region, [string]$Tier, [string]$Redundancy) { return $DiskPrice["$Region|$Tier|$Redundancy".ToLower()] }

$AppPrice = @{}
foreach ($p in $PriceApp) {
    $os = if ($p.ProductName -match 'Linux') { 'Linux' } else { 'Windows' }
    $sku = ([string]$p.SkuName -replace '\s', '').ToLower()
    $k = "$($p.Region)|$sku|$os".ToLower()
    $v = N $p.Price
    if ($null -ne $v -and $v -gt 0 -and ($null -eq $AppPrice[$k] -or $v -lt $AppPrice[$k])) { $AppPrice[$k] = $v }
}
function Get-AppHourly([string]$Region, [string]$Sku, [string]$Os) { return $AppPrice["$Region|$(($Sku -replace '\s','').ToLower())|$Os".ToLower()] }

function Get-FilesPremiumGiBMonth([string]$Region, [string]$Redundancy) {
    $p = $PriceFiles | Where-Object { $_.Region -eq $Region -and $_.MeterName -eq "Premium $Redundancy Provisioned" } | Select-Object -First 1
    if ($p) { return N $p.Price }
    return 0.16
}

# Log Analytics / Sentinel price book per region.
# Retail 'Sentinel' meters are the simplified (unified) prices that already include Log Analytics ingestion;
# classic Sentinel-only prices are derived as unified - Log Analytics for the same tier.
function Get-LaPriceBook([string]$Region) {
    $rows = @($PriceLa | Where-Object { $_.Region -eq $Region })
    $book = @{ LaPayg = $null; LaTiers = @{}; UniPayg = $null; UniTiers = @{}; SentClassicPayg = $null; SentClassicTiers = @{}; Retention = $null; Basic = $null; Source = 'Retail API' }
    foreach ($p in $rows) {
        $price = N $p.Price
        if ($p.ServiceName -eq 'Log Analytics' -and $p.MeterName -eq 'Analytics Logs Data Ingestion' -and $price -gt 0) { $book.LaPayg = $price }
        elseif ($p.ServiceName -eq 'Log Analytics' -and $p.MeterName -eq 'Analytics Logs Data Retention' -and $price -gt 0) { $book.Retention = $price }
        elseif ($p.ServiceName -eq 'Log Analytics' -and $p.MeterName -match 'Basic Logs Data Ingestion' -and $price -gt 0) { $book.Basic = $price }
        elseif ($p.ServiceName -eq 'Azure Monitor' -and $p.MeterName -match '^(\d+) GB Commitment Tier Capacity Reservation$') { $book.LaTiers[[int]$Matches[1]] = $price }
        elseif ($p.ServiceName -eq 'Sentinel' -and $p.MeterName -eq 'Pay-as-you-go Analysis' -and $price -gt 0) { $book.UniPayg = $price }
        elseif ($p.ServiceName -eq 'Sentinel' -and $p.MeterName -match '^(\d+) GB Commitment Tier Capacity Reservation$') { $book.UniTiers[[int]$Matches[1]] = $price }
    }
    if (-not $book.LaPayg -or $book.LaTiers.Count -eq 0) {
        # Fallback: public list prices (USD, East US) - indicative only
        $book.Source = 'Fallback list prices (US)'
        $book.LaPayg = 2.30; $book.Retention = 0.10; $book.Basic = 0.50
        $book.LaTiers = @{ 100 = 196; 200 = 368; 300 = 540; 400 = 704; 500 = 865; 1000 = 1700; 2000 = 3320; 5000 = 8050 }
    }
    if (-not $book.UniPayg) { $book.UniPayg = $book.LaPayg + 2.00; $book.Source = "$($book.Source); Sentinel fallback" }
    if ($book.UniTiers.Count -eq 0) { foreach ($t in $book.LaTiers.Keys) { $book.UniTiers[$t] = [double]$book.LaTiers[$t] + [double]$t } }
    $book.SentClassicPayg = [Math]::Max(0.01, $book.UniPayg - $book.LaPayg)
    foreach ($t in $book.UniTiers.Keys) { if ($book.LaTiers.ContainsKey($t) -and ([double]$book.UniTiers[$t] - [double]$book.LaTiers[$t]) -gt 0) { $book.SentClassicTiers[$t] = [double]$book.UniTiers[$t] - [double]$book.LaTiers[$t] } }
    return $book
}

function Get-TierPlan([double[]]$Daily, [double]$Payg, [hashtable]$Tiers) {
    <# Monthly cost for every tier option given the observed daily series (overage billed at the tier's effective per-GB rate). #>
    $days = [Math]::Max(1, $Daily.Count)
    $plans = [System.Collections.Generic.List[object]]::new()
    $plans.Add([pscustomobject]@{ Tier = 0; Name = 'Pay-as-you-go'; Monthly = [Math]::Round((($Daily | Measure-Object -Sum).Sum) * $Payg / $days * 30.4375, 0) })
    foreach ($t in ($Tiers.Keys | Sort-Object)) {
        $dayPrice = [double]$Tiers[$t]; $rate = $dayPrice / $t; $sum = 0.0
        foreach ($g in $Daily) { $sum += $dayPrice + [Math]::Max(0.0, $g - $t) * $rate }
        $plans.Add([pscustomobject]@{ Tier = [int]$t; Name = "$t GB/day commitment"; Monthly = [Math]::Round($sum / $days * 30.4375, 0) })
    }
    return $plans
}

# ---- VM size catalog (for right-sizing within the same series)
$VmSizeInfo = @{}
foreach ($v in $VmSizes) { $VmSizeInfo["$($v.Location)|$($v.VmSize)".ToLower()] = $v }

function Get-VmSeriesKey([string]$Size) {
    # Standard_D8s_v5 -> D|s|_v5 ; Standard_DS3_v2 -> DS|#|_v2 ; constrained (E16-4s_v3) -> unique
    if ($Size -notmatch '^(?i)(Standard|Basic)_([A-Z]+)(\d+)(-\d+)?([a-z]*)(_v\d+)?(_.*)?$') { return $null }
    if ($Matches[4]) { return $null }
    return ('{0}|{1}|{2}|{3}' -f $Matches[2].ToUpper(), $Matches[5].ToLower(), ([string]$Matches[6]).ToLower(), ([string]$Matches[7]).ToLower())
}

function Get-VmDownsizeTarget {
    param([string]$Size, [string]$Region, [double]$vCPUs, [double]$MemGB, [double]$CpuP95, $MemMaxPct, [int]$DataDisks, [string]$Os)
    $key = Get-VmSeriesKey $Size
    if (-not $key -or $vCPUs -le 1) { return $null }
    $reqCpu = [Math]::Max(1, [Math]::Ceiling($vCPUs * $CpuP95 / $TargetUtilization))
    if ($null -ne $MemMaxPct) { $reqMem = $MemGB * [double]$MemMaxPct / 80.0; $minCpu = 1 }
    else { $reqMem = $MemGB / 2; $minCpu = [Math]::Floor($vCPUs / 2) }   # memory not measured: at most one step down
    $reqCpu = [Math]::Max($reqCpu, [Math]::Max($minCpu, 2))
    $curPrice = Get-VmHourly $Region $Size $Os
    $cands = foreach ($v in ($VmSizes | Where-Object { $_.Location -eq $Region })) {
        if ((Get-VmSeriesKey $v.VmSize) -ne $key) { continue }
        $vc = N $v.vCPUs; $mg = N $v.MemoryGB; $dd = N $v.MaxDataDisks
        if ($null -eq $vc -or $null -eq $mg) { continue }
        if ($vc -ge $vCPUs -or $vc -lt $reqCpu -or $mg -lt $reqMem -or ($null -ne $dd -and $dd -lt $DataDisks)) { continue }
        $pr = Get-VmHourly $Region $v.VmSize $Os
        [pscustomobject]@{ Size = $v.VmSize; vCPUs = $vc; MemoryGB = $mg; Hourly = $pr }
    }
    $best = @($cands | Sort-Object @{ e = { if ($null -ne $_.Hourly) { $_.Hourly } else { $_.vCPUs * 1000 } } }, vCPUs) | Select-Object -First 1
    if (-not $best) { return $null }
    if ($null -ne $curPrice -and $null -ne $best.Hourly -and $best.Hourly -ge $curPrice) { return $null }
    return [pscustomobject]@{ Size = $best.Size; vCPUs = $best.vCPUs; MemoryGB = $best.MemoryGB; Hourly = $best.Hourly; CurrentHourly = $curPrice }
}

function Get-Savings {
    <# Prefer actual cost x price ratio (keeps customer discounts); fall back to retail delta. #>
    param($MonthlyCost, $CurrentUnit, $TargetUnit, [double]$HoursPerMonth = 730, [double]$Ratio = -1, [double]$Share = 1.0)
    if ($Ratio -ge 0) {
        if ($null -ne $MonthlyCost -and $MonthlyCost -gt 0) { return [Math]::Round($MonthlyCost * (1 - $Ratio) * $Share, 2) }
        return $null
    }
    if ($null -ne $CurrentUnit -and $null -ne $TargetUnit -and $CurrentUnit -gt 0) {
        if ($null -ne $MonthlyCost -and $MonthlyCost -gt 0) { return [Math]::Round($MonthlyCost * (1 - $TargetUnit / $CurrentUnit) * $Share, 2) }
        return [Math]::Round(($CurrentUnit - $TargetUnit) * $HoursPerMonth * $Share, 2)
    }
    return $null
}

#endregion

#region ============================== ANALYSIS FRAMEWORK ==============================

$script:Recs = [System.Collections.Generic.List[object]]::new()
$script:SheetNames = [ordered]@{
    VirtualMachines = 'Virtual Machines'; VMScaleSets = 'VM Scale Sets'; AKSNodePools = 'AKS Node Pools'; AzureVirtualDesktop = 'AVD Host Pools'
    ContainerApps = 'Container Apps'; ContainerAppEnvironments = 'Container App Envs'; ContainerInstances = 'Container Instances'; ContainerRegistries = 'Container Registries'
    AppServicePlans = 'App Service Plans'; AppServices = 'Web & Function Apps'
    SqlDatabases = 'SQL Databases'; SqlElasticPools = 'SQL Elastic Pools'; SqlManagedInstances = 'SQL Managed Instances'; SqlVirtualMachines = 'SQL Server on VMs'
    MySQLFlexibleServers = 'MySQL Flexible'; PostgreSQLFlexibleServers = 'PostgreSQL Flexible'; LegacyDatabaseServers = 'Legacy DB Servers'
    CosmosDB = 'Cosmos DB'; RedisCaches = 'Redis Cache'; ManagedRedis = 'Managed Redis'
    ManagedDisks = 'Managed Disks'; Snapshots = 'Snapshots'; StorageAccounts = 'Storage Accounts'; PremiumFileShares = 'Provisioned File Shares'
    NetAppCapacityPools = 'NetApp Pools'; RecoveryServicesVaults = 'Recovery Vaults'
    LogAnalyticsWorkspaces = 'Log Analytics & Sentinel'; LogAnalyticsClusters = 'LA Dedicated Clusters'; ApplicationInsights = 'App Insights'; ManagedHSMs = 'Managed HSM'
    EventHubNamespaces = 'Event Hubs'; ServiceBusNamespaces = 'Service Bus'; LogicApps = 'Logic Apps'; APIManagement = 'API Management'; IoTHubs = 'IoT Hub'; StreamAnalyticsJobs = 'Stream Analytics'
    ApplicationGateways = 'Application Gateways'; LoadBalancers = 'Load Balancers'; PublicIPAddresses = 'Public IPs'; NatGateways = 'NAT Gateways'; AzureFirewalls = 'Azure Firewalls'
    VirtualNetworkGateways = 'VPN & ER Gateways'; ExpressRouteCircuits = 'ExpressRoute Circuits'; BastionHosts = 'Bastion'; PrivateEndpoints = 'Private Endpoints'
    FrontDoorCdn = 'Front Door & CDN'; DdosProtectionPlans = 'DDoS Plans'; VirtualHubs = 'vWAN Hubs'
    DataFactories = 'Data Factory'; SynapseSqlPools = 'Synapse SQL Pools'; SynapseSparkPools = 'Synapse Spark Pools'; DataExplorerClusters = 'Data Explorer'
    Databricks = 'Databricks'; HDInsightClusters = 'HDInsight'; AnalysisServices = 'Analysis Services'; PowerBIFabricCapacities = 'PBI & Fabric Capacity'
    CognitiveServices = 'AI Services'; AIModelDeployments = 'AI Model Deployments'; SearchServices = 'AI Search'; MLWorkspaces = 'ML Workspaces'; MLComputes = 'ML Computes'
    BatchAccounts = 'Batch Accounts'; DedicatedHosts = 'Dedicated Hosts'; CapacityReservations = 'Capacity Reservations'; ElasticSans = 'Elastic SAN'; SignalRWebPubSub = 'SignalR & Web PubSub'
}
$script:TypeCategory = @{}
foreach ($c in $CollectionSummary) {
    $script:TypeCategory[$c.Key] = $c.Category
    if (-not $script:SheetNames.Contains($c.Key)) { $script:SheetNames[$c.Key] = [string]$c.Display }
}
$script:GenericKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($c in $CollectionSummary) { if ((B $c.Generic) -or ([string]$c.Key).StartsWith('Gen_')) { [void]$script:GenericKeys.Add($c.Key) } }

# Short prefixes used to build Rule IDs (e.g. VM-RS-SKU)
$script:TypePrefix = @{
    VirtualMachines = 'VM'; VMScaleSets = 'VMSS'; AKSNodePools = 'AKS'; AzureVirtualDesktop = 'AVD'; ContainerApps = 'ACA'; ContainerAppEnvironments = 'ACAENV'; ContainerInstances = 'ACI'
    ContainerRegistries = 'ACR'; AppServicePlans = 'ASP'; AppServices = 'APP'; SqlDatabases = 'SQLDB'; SqlElasticPools = 'SQLEP'; SqlManagedInstances = 'SQLMI'; SqlVirtualMachines = 'SQLVM'
    MySQLFlexibleServers = 'MYSQL'; PostgreSQLFlexibleServers = 'PGSQL'; LegacyDatabaseServers = 'DBSS'; CosmosDB = 'COSMOS'; RedisCaches = 'REDIS'; ManagedRedis = 'AMR'
    ManagedDisks = 'DISK'; Snapshots = 'SNAP'; StorageAccounts = 'STG'; PremiumFileShares = 'FILES'; NetAppCapacityPools = 'ANF'; RecoveryServicesVaults = 'RSV'
    LogAnalyticsWorkspaces = 'LAW'; LogAnalyticsClusters = 'LAC'; ApplicationInsights = 'APPI'; ManagedHSMs = 'HSM'; EventHubNamespaces = 'EVH'; ServiceBusNamespaces = 'SB'
    LogicApps = 'LOGIC'; APIManagement = 'APIM'; IoTHubs = 'IOT'; StreamAnalyticsJobs = 'ASA'; ApplicationGateways = 'AGW'; LoadBalancers = 'LB'; PublicIPAddresses = 'PIP'
    NatGateways = 'NAT'; AzureFirewalls = 'AFW'; VirtualNetworkGateways = 'VNG'; ExpressRouteCircuits = 'ER'; BastionHosts = 'BAS'; PrivateEndpoints = 'PE'; FrontDoorCdn = 'AFD'
    DdosProtectionPlans = 'DDOS'; VirtualHubs = 'VHUB'; DataFactories = 'ADF'; SynapseSqlPools = 'SYNSQL'; SynapseSparkPools = 'SYNSPK'; DataExplorerClusters = 'ADX'
    Databricks = 'DBX'; HDInsightClusters = 'HDI'; AnalysisServices = 'AAS'; PowerBIFabricCapacities = 'FAB'; CognitiveServices = 'AI'; AIModelDeployments = 'AIDEP'
    SearchServices = 'SRCH'; MLWorkspaces = 'MLW'; MLComputes = 'MLC'; BatchAccounts = 'BATCH'; DedicatedHosts = 'DHOST'; CapacityReservations = 'CRES'; ElasticSans = 'ESAN'; SignalRWebPubSub = 'SIGR'
}

# Recommendation rule catalog - every suggestion references one of these rules (criteria are rendered with the active thresholds)
$script:RuleCatalog = [ordered]@{
    'NOACT'        = @{ Name = 'No usage activity'; Criteria = 'All activity metrics (requests, transactions, messages, calls, connections, sessions, traffic) are zero for the whole {Days}-day window.'; Rationale = 'Capacity is paid for but not consumed - remove it or consolidate the workload.' }
    'DEALLOC'      = @{ Name = 'Deallocated for the entire window'; Criteria = 'VM power state is deallocated and it ran 0% of the hours in the window.'; Rationale = 'Compute is free while deallocated, but disks, IPs and backups keep billing.' }
    'STOPPED'      = @{ Name = 'Stopped but still billed'; Criteria = "VM power state is 'stopped' (shut down from inside the OS without deallocation)."; Rationale = 'Compute charges continue until the VM is deallocated.' }
    'ORPHAN'       = @{ Name = 'Orphaned / unattached'; Criteria = 'Resource has no attachment, backend, connection, subnet, VNet or child workload (unattached disk/IP, empty plan/pool, gateway without connections, host without VMs...).'; Rationale = 'Billed without serving any workload.' }
    'IDLE'         = @{ Name = 'Idle'; Criteria = 'Average utilization < {Idle}% and P95 < {Idle2}% (memory < 50% where measured).'; Rationale = 'Capacity is provisioned but essentially unused.' }
    'RS-SKU'       = @{ Name = 'Right-size SKU / size'; Criteria = 'P95 utilization < {Low}% (and peak memory < {MemLow}% where measured); the smaller SKU keeps P95 at or below {Target}%.'; Rationale = 'Match the provisioned SKU to measured demand.' }
    'RS-SCALE'     = @{ Name = 'Reduce instances / nodes / units'; Criteria = 'Measured P95 demand fits on fewer instances, nodes, units, replicas or throughput while keeping P95 at or below {Target}%.'; Rationale = 'Scale units are billed per hour regardless of load.' }
    'RS-STORAGE'   = @{ Name = 'Reduce provisioned capacity'; Criteria = 'Provisioned capacity (share quota, max DB size, reserved storage, pool / SAN size, provisioned IOPS) is far above the capacity used.'; Rationale = 'Provisioned storage is billed on size, not usage.' }
    'DISK-TIER'    = @{ Name = 'Disk performance tier'; Criteria = 'Premium SSD whose P95 IOPS and throughput fit within the Standard SSD limits of the same size.'; Rationale = 'Standard SSD costs significantly less for low-IO workloads.' }
    'AUTOSCALE'    = @{ Name = 'Autoscale / minimum capacity'; Criteria = 'Autoscale minimum, fixed capacity or always-on replicas/nodes exceed what the measured demand requires.'; Rationale = 'Minimum capacity is billed 24x7.' }
    'SCHEDULE'     = @{ Name = 'Off-hours schedule'; Criteria = 'Non-production (or bursty) resource running >= 90% of hours, or idle outside business hours ({Biz}, weekend {Weekend}).'; Rationale = 'Stop / pause outside business hours.' }
    'TIER-NONPROD' = @{ Name = 'Premium option in non-production'; Criteria = 'Non-production resource uses a premium tier, geo replication, HA or multi-region option.'; Rationale = 'Premium options are rarely required outside production.' }
    'LICENSE'      = @{ Name = 'License optimization'; Criteria = 'Windows / SQL Server license is included in the price, or a premium edition runs a low-utilized workload.'; Rationale = 'Azure Hybrid Benefit or a lower edition removes license cost.' }
    'COMMIT-TIER'  = @{ Name = 'Commitment tier vs. actual usage'; Criteria = 'Cheapest tier recalculated from the actual daily series (pay-as-you-go vs every commitment tier, overage at the tier rate).'; Rationale = 'Commitment tiers are billed in full every day regardless of usage.' }
    'COMMIT-RI'    = @{ Name = 'Reservation / savings plan candidate'; Criteria = 'Running >= {Steady}% of hours, right-sized, and billed close to pay-as-you-go rates (no reservation / savings plan benefit detected).'; Rationale = 'Steady 24x7 workloads save 30-65% with 1/3-year commitments.' }
    'RETENTION'    = @{ Name = 'Data retention'; Criteria = 'Interactive retention configured beyond the included free period.'; Rationale = 'Move older data to low-cost long-term retention.' }
    'INGESTION'    = @{ Name = 'Log ingestion volume'; Criteria = 'One table contributes >= 30% of billable ingestion and >= 10 GB/day.'; Rationale = 'Filter, transform or move verbose tables to Basic / Auxiliary logs.' }
    'LIFECYCLE'    = @{ Name = 'Storage tiering'; Criteria = 'More than 1 TB in the Hot tier with a low transaction rate.'; Rationale = 'Lifecycle policies move cold data to Cool / Cold / Archive.' }
    'MODERNIZE'    = @{ Name = 'Retiring / previous-generation SKU'; Criteria = 'Resource uses a retired or previous-generation SKU (Basic LB/IP, App Gateway v1, classic services, Single Server, older App Service tiers...).'; Rationale = 'Newer SKUs have better price/performance and avoid forced migrations.' }
    'SNAPSHOT'     = @{ Name = 'Aged / orphaned snapshot'; Criteria = 'Snapshot older than {SnapAge} days or its source disk no longer exists.'; Rationale = 'Snapshots are billed per GB-month indefinitely.' }
    'BACKUP'       = @{ Name = 'Backups of deleted resources'; Criteria = 'Backup items whose source resource no longer exists.'; Rationale = 'Retained backup data keeps billing per instance and per GB.' }
    'SERVERLESS'   = @{ Name = 'Serverless / consumption model'; Criteria = 'Very low sustained utilization on a provisioned capacity model.'; Rationale = 'Serverless billing charges only for actual use.' }
    'CONSOLIDATE'  = @{ Name = 'Consolidate'; Criteria = 'A pool, plan or environment hosts a single workload.'; Rationale = 'Consolidation removes fixed per-container overhead.' }
    'PERF-RISK'    = @{ Name = 'Performance risk (under-provisioned)'; Criteria = 'P95 utilization > {High}% or throttling / near-limit signals (429s, queueing, quota).'; Rationale = 'Scale or tune before optimizing cost - protects user experience.' }
    'GEN-UTIL'     = @{ Name = 'Low utilization (auto-discovered metric)'; Criteria = 'P95 of the auto-discovered capacity utilization metric (CPU / memory / load / units, not service-limit headroom) < {GenLow}% of the current SKU capacity, and the resource has scale units or costs at least 5 per month.'; Rationale = 'Reduce units / SKU after validating workload requirements.' }
    'GEN-NOACT'    = @{ Name = 'No activity (auto-discovered metrics)'; Criteria = 'All auto-discovered usage metrics (requests, queries, messages, calls, events - configuration counts excluded) are zero in the window while the resource is billed at least {GenMinCost} per month.'; Rationale = 'Paid but unused.' }
    'GEN-HOT'      = @{ Name = 'High utilization / near service limit (auto-discovered metric)'; Criteria = 'P95 of the auto-discovered utilization metric > {GenHigh}%, or a service-limit metric (quota / saturation) above {GenHigh}% of the limit.'; Rationale = 'Performance / throttling risk - scale, split the workload or request a limit increase.' }
    'GENERAL'      = @{ Name = 'General optimization'; Criteria = 'Configuration-based finding.'; Rationale = 'Cost optimization best practice.' }
}
function Format-RuleText([string]$t) {
    $subs = @{
        '{Days}' = $LookbackDays; '{Idle}' = $CpuIdleThreshold; '{Idle2}' = (2 * $CpuIdleThreshold); '{Low}' = $CpuLowThreshold; '{MemLow}' = $MemLowThreshold
        '{Target}' = $TargetUtilization; '{High}' = $CpuHighThreshold; '{Steady}' = $SteadyHoursPct; '{SnapAge}' = $SnapshotMaxAge; '{GenLow}' = $GenericLow; '{GenHigh}' = $GenericHigh; '{GenMinCost}' = $GenericMinRecCost
        '{Biz}' = "$($Meta.BusinessHoursStart):00-$($Meta.BusinessHoursEnd):00"; '{Weekend}' = (@($Meta.WeekendDays) -join '/')
    }
    foreach ($k in $subs.Keys) { $t = $t.Replace($k, [string]$subs[$k]) }
    return $t
}

$script:HowTo = @{
    VirtualMachines = 'How to resize a virtual machine: https://learn.microsoft.com/azure/virtual-machines/sizes/resize-vm'
    VMScaleSets = 'Scale sets autoscale: https://learn.microsoft.com/azure/virtual-machine-scale-sets/virtual-machine-scale-sets-autoscale-overview'
    AKSNodePools = 'AKS cluster autoscaler: https://learn.microsoft.com/azure/aks/cluster-autoscaler'
    ManagedDisks = 'Change disk type/performance: https://learn.microsoft.com/azure/virtual-machines/disks-convert-types'
    Snapshots = 'Manage snapshots: https://learn.microsoft.com/azure/virtual-machines/snapshot-copy-managed-disk'
    StorageAccounts = 'Storage cost optimization: https://learn.microsoft.com/azure/storage/common/storage-plan-manage-costs'
    PremiumFileShares = 'Resize a file share: https://learn.microsoft.com/azure/storage/files/storage-how-to-create-file-share'
    AppServicePlans = 'Scale up an App Service plan: https://learn.microsoft.com/azure/app-service/manage-scale-up'
    AppServices = 'App Service cost: https://learn.microsoft.com/azure/app-service/overview-manage-costs'
    SqlDatabases = 'Scale single database: https://learn.microsoft.com/azure/azure-sql/database/single-database-scale'
    SqlElasticPools = 'Scale elastic pool: https://learn.microsoft.com/azure/azure-sql/database/elastic-pool-scale'
    SqlManagedInstances = 'Manage MI resources: https://learn.microsoft.com/azure/azure-sql/managed-instance/management-operations-overview'
    MySQLFlexibleServers = 'Scale MySQL flexible server: https://learn.microsoft.com/azure/mysql/flexible-server/how-to-manage-server-portal'
    PostgreSQLFlexibleServers = 'Scale PostgreSQL flexible server: https://learn.microsoft.com/azure/postgresql/flexible-server/how-to-scale-compute-storage-portal'
    CosmosDB = 'Optimize provisioned throughput: https://learn.microsoft.com/azure/cosmos-db/optimize-cost-throughput'
    RedisCaches = 'Scale Azure Cache for Redis: https://learn.microsoft.com/azure/azure-cache-for-redis/cache-how-to-scale'
    LogAnalyticsWorkspaces = 'Change pricing tier: https://learn.microsoft.com/azure/azure-monitor/logs/change-pricing-tier'
    EventHubNamespaces = 'Event Hubs scalability: https://learn.microsoft.com/azure/event-hubs/event-hubs-scalability'
    ServiceBusNamespaces = 'Service Bus premium scaling: https://learn.microsoft.com/azure/service-bus-messaging/service-bus-premium-messaging'
    ApplicationGateways = 'App Gateway autoscaling: https://learn.microsoft.com/azure/application-gateway/application-gateway-autoscaling-zone-redundant'
    VirtualNetworkGateways = 'Resize VPN gateway: https://learn.microsoft.com/azure/vpn-gateway/gateway-sku-resize'
    AzureFirewalls = 'Azure Firewall SKUs: https://learn.microsoft.com/azure/firewall/choose-firewall-sku'
    DdosProtectionPlans = 'DDoS IP Protection: https://learn.microsoft.com/azure/ddos-protection/ddos-protection-sku-comparison'
    CognitiveServices = 'Provisioned throughput: https://learn.microsoft.com/azure/ai-services/openai/concepts/provisioned-throughput'
    AIModelDeployments = 'Provisioned throughput: https://learn.microsoft.com/azure/ai-services/openai/concepts/provisioned-throughput'
    SynapseSqlPools = 'Pause/resume dedicated SQL pool: https://learn.microsoft.com/azure/synapse-analytics/sql-data-warehouse/pause-and-resume-compute-portal'
    Default = 'Azure cost optimization: https://learn.microsoft.com/azure/cost-management-billing/costs/cost-mgt-best-practices'
}

function Get-Environment($row) {
    $tag = [string]$row.EnvironmentTag
    if ($tag -match $EnvProdTag) { return 'Production' }
    if ($tag -match $EnvNonProdTag) { return 'Non-Production' }
    $text = " $($row.SubscriptionName) $($row.ResourceGroup) $($row.ResourceName) ".ToLower()
    if ($text -match $EnvNonProdName) { return 'Non-Production' }
    if ($text -match $EnvProdName) { return 'Production' }
    return 'Unclassified'
}

function Set-Analysis {
    param($row, [string]$Status, [string]$Provisioned, [string]$Utilized, $UtilPct, [string]$Note)
    $row | Add-Member -NotePropertyName UtilizationStatus -NotePropertyValue $Status -Force
    $row | Add-Member -NotePropertyName ProvisionedSummary -NotePropertyValue $Provisioned -Force
    $row | Add-Member -NotePropertyName UtilizedSummary -NotePropertyValue $Utilized -Force
    $row | Add-Member -NotePropertyName UtilizationPct -NotePropertyValue (Rnd $UtilPct 1) -Force
    if ($Note) { $row | Add-Member -NotePropertyName AnalysisNote -NotePropertyValue $Note -Force }
}

function Add-Rec {
    param($row, [string]$Key, [ValidateSet('Rightsize', 'Delete', 'Configure', 'Performance Risk')][string]$Category, [string]$Action,
        [string]$Target, $MonthlySavings, [ValidateSet('High', 'Medium', 'Low')][string]$Confidence = 'Medium', [string]$Evidence, [string]$Effort = '', [string]$Comment, [string]$Rule = 'GENERAL')
    if (-not $Effort) { $Effort = switch ($Category) { 'Delete' { 'Low' } 'Configure' { 'Low' } default { 'Medium' } } }
    $ms = if ($null -ne $MonthlySavings) { [Math]::Max(0.0, [Math]::Round([double]$MonthlySavings, 2)) } else { $null }
    $name = [string]$row.ResourceName
    if ($row.PSObject.Properties['NodePool'] -and $row.NodePool) { $name = "$name / $($row.NodePool)" }
    elseif ($row.PSObject.Properties['ShareName'] -and $row.ShareName) { $name = "$name / $($row.ShareName)" }
    elseif ($row.PSObject.Properties['Deployment'] -and $row.Deployment) { $name = "$name / $($row.Deployment)" }
    elseif ($row.PSObject.Properties['ComputeName'] -and $row.ComputeName) { $name = "$name / $($row.ComputeName)" }
    $rc = $script:RuleCatalog[$Rule] ?? $script:RuleCatalog['GENERAL']
    $prefix = $script:TypePrefix[$Key] ?? $(if ($Key -like 'Gen_*') { 'GEN' } else { ($Key -replace '[^A-Z]', '') })
    $script:Recs.Add([pscustomobject][ordered]@{
            Priority = ''; SubscriptionName = $row.SubscriptionName; ResourceGroup = $row.ResourceGroup; ResourceName = $name
            ResourceType = $script:SheetNames[$Key] ?? $row.ResourceType; Location = $row.Location; Environment = $row.Environment
            Category = $Category; RuleId = $(if ($prefix -eq 'GEN' -and $Rule.StartsWith('GEN-')) { $Rule } else { "$prefix-$Rule" }); RuleName = $rc.Name; Criteria = (Format-RuleText $rc.Criteria)
            CurrentConfiguration = $row.ProvisionedSummary; ObservedUtilization = $row.UtilizedSummary
            RecommendedAction = $Action; TargetConfiguration = $Target; CurrentMonthlyCost = $row.MonthlyCost
            EstMonthlySavings = $ms; EstAnnualSavings = if ($null -ne $ms) { [Math]::Round($ms * 12, 0) } else { $null }
            Confidence = $Confidence; Effort = $Effort; Evidence = $Evidence; Rationale = $rc.Rationale
            Comment = if ($Comment) { $Comment } else { $script:HowTo[$Key] ?? $script:HowTo.Default }
            CustomerComment = ''; ResourceId = $row.ResourceId; TypeKey = $Key; RuleCode = $Rule
        })
}

#endregion

#region ============================== ANALYZERS: COMPUTE ==============================

$Analyzers = @{}
$script:DiskCostByVm = @{}; $script:AppsByPlan = @{}; $script:VmById = @{}

function Get-StatusFromCpu($avg, $p95, $memP95, $memMax, [double]$IdleMemCap = 50) {
    if ($null -eq $p95) { return 'No data' }
    if ($avg -lt $CpuIdleThreshold -and $p95 -lt (2 * $CpuIdleThreshold) -and ($null -eq $memMax -or $memMax -lt $IdleMemCap)) { return 'Idle' }
    if ($p95 -gt $CpuHighThreshold -or ($null -ne $memP95 -and $memP95 -gt $MemHighThreshold)) { return 'Under-provisioned' }
    if ($p95 -lt $CpuLowThreshold -and ($null -eq $memMax -or $memMax -lt $MemLowThreshold)) { return 'Over-provisioned' }
    return 'Right-sized'
}

function Get-LadderTarget([double[]]$Ladder, [double]$Current, [double]$Required) {
    $t = @($Ladder | Where-Object { $_ -ge $Required -and $_ -lt $Current } | Sort-Object) | Select-Object -First 1
    return $t
}

$Analyzers.VirtualMachines = { param($row)
    $cpuAvg = N $row.CPU_Avg; $cpuP95 = N $row.CPU_P95; $cpuMax = N $row.CPU_Max
    $memMax = N $row.MemUsedPct_Max; $memP95 = N $row.MemUsedPct_P95; $active = N $row.Pattern_ActiveHoursPct
    $vc = N $row.vCPUs; $mem = N $row.MemoryGB; $size = [string]$row.VmSize
    $prov = "$size ($(Fmt $vc 0) vCPU, $(Fmt $mem 0) GB) $($row.OsType)"
    $util = if ($null -ne $cpuP95) { "CPU avg $(Fmt $cpuAvg)% | P95 $(Fmt $cpuP95)% | max $(Fmt $cpuMax)%; memory P95 $(Fmt $memP95)% | max $(Fmt $memMax)%; running $(Fmt $active 0)% of hours" } else { 'No CPU data in window' }
    $evidence = "CPU avg $(Fmt $cpuAvg)%, P95 $(Fmt $cpuP95)%, max $(Fmt $cpuMax)%; memory max $(if ($null -ne $memMax) { "$(Fmt $memMax)%" } else { 'not measured' }) (past $LookbackDays days)"
    if ($row.ManagedByService -in @('Databricks', 'AKS')) { Set-Analysis $row 'Managed by service' $prov $util $cpuP95 "Lifecycle managed by $($row.ManagedByService)"; return }
    if ($row.PowerState -eq 'deallocated') {
        if (($active ?? 0) -lt 1) {
            Set-Analysis $row 'Deallocated' $prov 'Deallocated for the entire window' $null
            $diskCost = $script:DiskCostByVm[([string]$row.ResourceId).ToLower()]
            Add-Rec $row 'VirtualMachines' 'Delete' 'Deallocated for the entire window - confirm with owner; delete the VM and its disks (disks are still billed) or snapshot and delete' 'Delete VM + disks' $diskCost 'High' "Power state deallocated; 0 running hours in $LookbackDays days; attached disks cost $(Money $diskCost)/month" -Rule 'DEALLOC'
        }
        else { Set-Analysis $row 'Intermittent' $prov $util $cpuP95 "Currently deallocated; ran $(Fmt $active 0)% of hours" }
        return
    }
    if ($row.PowerState -eq 'stopped') {
        Set-Analysis $row 'Stopped (billed)' $prov $util $cpuP95
        Add-Rec $row 'VirtualMachines' 'Delete' 'VM is stopped but NOT deallocated - compute is still billed. Deallocate (Stop in portal) or delete' 'Deallocate' $row.MonthlyCost 'High' "Power state 'stopped' (shut down from inside the OS without deallocation)" -Rule 'STOPPED'
        return
    }
    $status = Get-StatusFromCpu $cpuAvg $cpuP95 $memP95 $memMax
    Set-Analysis $row $status $prov $util $cpuP95
    if ($status -in @('Idle', 'Over-provisioned') -and $row.Priority -ne 'Spot') {
        $t = Get-VmDownsizeTarget -Size $size -Region $row.Location -vCPUs ($vc ?? 0) -MemGB ($mem ?? 0) -CpuP95 $cpuP95 -MemMaxPct $memMax -DataDisks ([int]((N $row.DataDiskCount) ?? 0)) -Os $row.OsType
        if ($t) {
            $sav = Get-Savings -MonthlyCost $row.MonthlyCost -CurrentUnit $t.CurrentHourly -TargetUnit $t.Hourly -HoursPerMonth (730 * [Math]::Max(0.05, ($active ?? 100) / 100))
            $act = if ($status -eq 'Idle') { "Idle - validate with owner; delete/deallocate if unused, otherwise change from $size to $($t.Size)" } else { "Change from $size to $($t.Size)" }
            Add-Rec $row 'VirtualMachines' 'Rightsize' $act "$($t.Size) ($($t.vCPUs) vCPU, $($t.MemoryGB) GB)" $sav $(if ($null -ne $memMax) { 'High' } else { 'Medium' }) $evidence -Rule 'RS-SKU'
        }
        elseif ($status -eq 'Idle') {
            Add-Rec $row 'VirtualMachines' 'Delete' 'Idle VM (no smaller size in series) - validate with owner and deallocate / delete if no longer required' 'Deallocate or delete' $row.MonthlyCost 'Medium' $evidence -Rule 'IDLE'
        }
    }
    elseif ($status -eq 'Under-provisioned') {
        Add-Rec $row 'VirtualMachines' 'Performance Risk' "Sustained high utilization - consider scaling up or out (P95 CPU $(Fmt $cpuP95)%, P95 memory $(Fmt $memP95)%)" 'Larger size / scale-out' $null 'Medium' $evidence -Rule 'PERF-RISK'
    }
    if ($row.Environment -eq 'Non-Production' -and ($active ?? 0) -ge 90 -and $status -ne 'Idle') {
        $offSave = if ($row.MonthlyCost) { [Math]::Round($row.MonthlyCost * (1 - ($BizHoursPerWeek + 10) / 168.0), 2) } else { $null }
        Add-Rec $row 'VirtualMachines' 'Configure' 'Non-production VM running 24x7 - configure auto-shutdown / start-stop schedule outside business hours' 'Business-hours schedule' $offSave 'Medium' "Running $(Fmt $active 0)% of hours; CPU business-hours avg $(Fmt (N $row.Pattern_BizHoursAvg))% vs off-hours $(Fmt (N $row.Pattern_OffHoursAvg))%" -Rule 'SCHEDULE'
    }
    if ($row.OsType -eq 'Windows' -and [string]::IsNullOrEmpty([string]$row.LicenseType) -and [string]$row.ImagePublisher -notmatch 'microsoftsqlserver|MicrosoftWindowsDesktop') {
        $w = Get-VmHourly $row.Location $size 'Windows'; $l = Get-VmHourly $row.Location $size 'Linux'
        $sav = if ($w -and $l) { [Math]::Round(($w - $l) * 730 * [Math]::Max(0.05, ($active ?? 100) / 100), 2) } else { $null }
        Add-Rec $row 'VirtualMachines' 'Configure' 'Apply Azure Hybrid Benefit for Windows Server (requires eligible licenses with Software Assurance)' 'LicenseType = Windows_Server' $sav 'Low' 'Windows VM billed with the Windows license (LicenseType not set)' 'Low' 'Azure Hybrid Benefit: https://learn.microsoft.com/azure/virtual-machines/windows/hybrid-use-benefit-licensing' -Rule 'LICENSE'
    }
    # Effective discount vs. retail pay-as-you-go and reservation / savings plan candidates for steady workloads.
    # Azure Hybrid Benefit VMs are billed at the base (Linux) compute rate - no Windows license uplift applies.
    $ahub = $row.OsType -match '(?i)windows' -and [string]$row.LicenseType -match '(?i)^windows_(server|client)$'
    $payg = Get-VmHourly $row.Location $size $(if ($ahub) { 'Linux' } else { $row.OsType })
    $hoursRun = 730 * [Math]::Max(0.01, ($active ?? 0) / 100)
    if ($payg -and $row.MonthlyCost -gt 0 -and ($active ?? 0) -gt 0) {
        $row | Add-Member -NotePropertyName RetailPaygMonthly -NotePropertyValue ([Math]::Round($payg * $hoursRun, 2)) -Force
        $row | Add-Member -NotePropertyName EffectiveDiscountPct -NotePropertyValue ([Math]::Round(100 * (1 - [double]$row.MonthlyCost / ($payg * $hoursRun)), 1)) -Force
    }
    $cm = $VmCommit["$($row.Location)|$size|$(if ($row.OsType -match '(?i)windows') { 'Windows' } else { 'Linux' })".ToLower()]
    $cmLinux = $VmCommit["$($row.Location)|$size|linux".ToLower()]
    if ($cm -and $payg -and ($active ?? 0) -ge $SteadyHoursPct -and $status -in @('Right-sized', 'Under-provisioned') -and $row.Priority -ne 'Spot') {
        $disc = N $row.EffectiveDiscountPct
        if ($null -eq $disc -or $disc -lt 15) {
            # Reservation and savings plan rates cover compute only - the Windows license stays at pay-as-you-go
            $licAdj = if ($row.OsType -match '(?i)windows' -and -not $ahub) { [Math]::Max(0.0, $payg - [double]((Get-VmHourly $row.Location $size 'Linux') ?? $payg)) } else { 0.0 }
            $best = @(@{ N = '3-year reservation'; P = $cm.RI3Y }, @{ N = '1-year reservation'; P = $cm.RI1Y }, @{ N = '3-year savings plan'; P = $cm.SP3Y ?? $cmLinux.SP3Y }, @{ N = '1-year savings plan'; P = $cm.SP1Y ?? $cmLinux.SP1Y }) | Where-Object { $_.P } | ForEach-Object { $_.P2 = [double]$_.P + $licAdj; $_ }
            $oneYear = @($best | Where-Object { $_.N -like '1-year*' } | Sort-Object { $_.P2 } | Select-Object -First 1)
            if ($oneYear.Count) {
                $o = $oneYear[0]; $ratio = $o.P2 / $payg
                $sav = Get-Savings -MonthlyCost $row.MonthlyCost -Ratio $ratio
                $three = @($best | Where-Object { $_.N -like '3-year*' } | Sort-Object { $_.P2 } | Select-Object -First 1)
                Add-Rec $row 'VirtualMachines' 'Configure' "Steady 24x7 VM ($(Fmt $active 0)% of hours) billed at pay-as-you-go - cover with a $($o.N) (~$(Fmt (100 * (1 - $ratio)) 0)% saving$(if ($three.Count) { "; $($three[0].N) ~$(Fmt (100 * (1 - $three[0].P2 / $payg)) 0)%" }))" $o.N $sav 'Medium' "Running $(Fmt $active 0)% of hours; effective discount vs retail $(if ($null -ne $disc) { "$(Fmt $disc 0)%" } else { 'n/a' }); P95 CPU $(Fmt $cpuP95)%" 'Low' 'Reservations & savings plans: https://learn.microsoft.com/azure/cost-management-billing/savings-plan/decide-between-savings-plan-reservation' -Rule 'COMMIT-RI'
            }
        }
    }
}

$Analyzers.VMScaleSets = { param($row)
    $cpuAvg = N $row.CPU_Avg; $cpuP95 = N $row.CPU_P95; $cpuMax = N $row.CPU_Max; $memMax = N $row.MemUsedPct_Max; $memP95 = N $row.MemUsedPct_P95
    $n = [int]((N $row.Instances) ?? 0); $auto = B $row.AutoscaleEnabled
    $prov = "$($row.VmSize) x $n instance(s) ($(Fmt (N $row.vCPUsPerInstance) 0) vCPU each)$(if ($auto) { ", autoscale $($row.AutoscaleMin)-$($row.AutoscaleMax)" })"
    $util = if ($null -ne $cpuP95) { "CPU avg $(Fmt $cpuAvg)% | P95 $(Fmt $cpuP95)% | max $(Fmt $cpuMax)%; memory max $(Fmt $memMax)%" } else { 'No CPU data' }
    if ($row.ManagedByService -eq 'AKS') { Set-Analysis $row 'Managed by service' $prov 'Analyzed on the AKS Node Pools sheet' $null; return }
    if ($row.ManagedByService -eq 'Databricks') { Set-Analysis $row 'Managed by service' $prov $util $cpuP95; return }
    if ($n -eq 0) { Set-Analysis $row 'Stopped' $prov 'Scaled to zero instances' $null; return }
    $status = Get-StatusFromCpu $cpuAvg $cpuP95 $memP95 $memMax
    Set-Analysis $row $status $prov $util $cpuP95
    $evidence = "CPU avg $(Fmt $cpuAvg)%, P95 $(Fmt $cpuP95)%, max $(Fmt $cpuMax)% across $n instance(s)"
    if ($status -in @('Idle', 'Over-provisioned')) {
        $target = [int][Math]::Max(1, [Math]::Ceiling($n * [Math]::Max($cpuP95 / $TargetUtilization, [double]($memMax ?? 0) / 80.0)))
        if ($auto -and (N $row.AutoscaleMin) -gt $target) {
            Add-Rec $row 'VMScaleSets' 'Rightsize' "Lower autoscale minimum from $($row.AutoscaleMin) to $target instance(s)" "Autoscale min $target" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($target / [double](N $row.AutoscaleMin))) 'Medium' $evidence -Rule 'AUTOSCALE'
        }
        elseif (-not $auto -and $target -lt $n) {
            Add-Rec $row 'VMScaleSets' 'Rightsize' "Reduce instance count from $n to $target (and enable autoscale)" "$target instance(s) + autoscale" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($target / [double]$n)) 'Medium' $evidence -Rule 'RS-SCALE'
        }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'VMScaleSets' 'Performance Risk' 'Sustained high utilization - add instances or enable autoscale' 'Scale out' $null 'Medium' $evidence -Rule 'PERF-RISK' }
}

$Analyzers.AKSNodePools = { param($row)
    $cpuAvg = N $row.NodeCPU_Avg; $cpuP95 = N $row.NodeCPU_P95; $cpuMax = N $row.NodeCPU_Max; $memP95 = N $row.NodeMemory_P95; $memMax = N $row.NodeMemory_Max
    $n = [int]((N $row.NodeCount) ?? 0); $auto = B $row.AutoscaleEnabled
    $prov = "Pool '$($row.NodePool)' ($($row.PoolMode)): $($row.VmSize) x $n node(s)$(if ($auto) { ", autoscale $($row.MinCount)-$($row.MaxCount)" })"
    $util = if ($null -ne $cpuP95) { "Node CPU avg $(Fmt $cpuAvg)% | P95 $(Fmt $cpuP95)% | max $(Fmt $cpuMax)%; memory working set P95 $(Fmt $memP95)% | max $(Fmt $memMax)%" } else { 'No node metrics' }
    if ($row.ClusterPowerState -eq 'Stopped') { Set-Analysis $row 'Stopped' $prov 'Cluster stopped' $null; return }
    $status = Get-StatusFromCpu $cpuAvg $cpuP95 $memP95 $memMax 40
    Set-Analysis $row $status $prov $util $cpuP95
    $evidence = "Node CPU P95 $(Fmt $cpuP95)%, memory P95 $(Fmt $memP95)% on $n node(s) (past $LookbackDays days)"
    $floor = if ($row.PoolMode -eq 'System') { if ($row.Environment -eq 'Production') { 2 } else { 1 } } else { if ($auto) { 0 } else { 1 } }
    $nodeHourly = Get-VmHourly $row.Location $row.VmSize $(if ($row.OsType -eq 'Windows') { 'Windows' } else { 'Linux' })
    if ($status -in @('Idle', 'Over-provisioned') -and $null -ne $cpuP95) {
        $target = [int][Math]::Max($floor, [Math]::Ceiling($n * [Math]::Max($cpuP95 / $TargetUtilization, [double]($memP95 ?? 0) / 80.0)))
        $minCount = [int]((N $row.MinCount) ?? $n)
        if ($auto -and $minCount -gt $target) {
            $sav = if ($row.MonthlyCost) { Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($target / [double][Math]::Max(1, $n)) } elseif ($nodeHourly) { [Math]::Round(($minCount - $target) * $nodeHourly * 730, 2) } else { $null }
            Add-Rec $row 'AKSNodePools' 'Rightsize' "Lower cluster-autoscaler minimum for pool '$($row.NodePool)' from $minCount to $target node(s)" "min $target / max $($row.MaxCount)" $sav 'Medium' $evidence -Rule 'AUTOSCALE'
        }
        elseif (-not $auto -and $target -lt $n) {
            $sav = if ($row.MonthlyCost) { Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($target / [double]$n) } elseif ($nodeHourly) { [Math]::Round(($n - $target) * $nodeHourly * 730, 2) } else { $null }
            Add-Rec $row 'AKSNodePools' 'Rightsize' "Decrease node pool '$($row.NodePool)' size from $n to $target node(s) and enable cluster autoscaler" "$target node(s), autoscale $target-$n" $sav 'Medium' $evidence -Rule 'RS-SCALE'
        }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'AKSNodePools' 'Performance Risk' "Node pool '$($row.NodePool)' running hot - enable/raise autoscaler maximum or use larger nodes" 'Scale out' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ($row.Environment -eq 'Non-Production' -and $row.PoolMode -eq 'System') {
        Add-Rec $row 'AKSNodePools' 'Configure' 'Non-production cluster - use AKS start/stop outside business hours' 'Scheduled stop/start' $(if ($row.MonthlyCost) { [Math]::Round($row.MonthlyCost * (1 - ($BizHoursPerWeek + 10) / 168.0), 2) } else { $null }) 'Medium' 'Non-production cluster running 24x7' 'Low' 'Stop/start AKS: https://learn.microsoft.com/azure/aks/start-stop-cluster' -Rule 'SCHEDULE'
    }
}

#endregion

#region ============================== ANALYZERS: CONTAINERS & APP SERVICE ==============================

$Analyzers.ContainerApps = { param($row)
    $cpu = N $row.CpuUtilPct_P95; $memU = N $row.MemUtilPct_P95; $req = N $row.Requests_Total; $rep = N $row.Replicas_Avg; $minR = N $row.MinReplicas
    $prov = "$($row.CpuCoresPerReplica) vCPU / $($row.MemoryGiBPerReplica) GiB per replica; replicas $($row.MinReplicas)-$($row.MaxReplicas) ($(if ($row.WorkloadProfile) { $row.WorkloadProfile } else { 'Consumption' }))"
    $util = "CPU P95 $(Fmt $cpu)% of allocation; memory P95 $(Fmt $memU)%; requests $(Fmt $req 0); avg replicas $(Fmt $rep)"
    if ($null -eq $cpu -and $null -eq $req) { Set-Analysis $row 'No data' $prov $util $null; return }
    $ingress = B $row.IngressEnabled
    $noTraffic = if ($ingress) { ($req ?? 0) -eq 0 } else { $null -ne $cpu -and $cpu -lt 1 }
    if (($minR ?? 0) -ge 1 -and $noTraffic) {
        Set-Analysis $row 'Idle' $prov $util $cpu
        $why = if ($ingress) { "No requests in $LookbackDays days but min replicas = $minR" } else { "Background app (no ingress) with near-zero CPU (P95 $(Fmt $cpu)% of allocation) and min replicas = $minR" }
        Add-Rec $row 'ContainerApps' 'Configure' "$why - scale to zero (minReplicas 0, event-driven KEDA scaling) or delete" 'minReplicas = 0' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.2) 'Medium' $util -Rule 'AUTOSCALE'
        return
    }
    $status = if ($null -eq $cpu) { 'In use' } elseif ($cpu -lt 20 -and ($memU ?? 0) -lt 40) { 'Over-provisioned' } elseif ($cpu -gt $CpuHighThreshold -or ($memU ?? 0) -gt $MemHighThreshold) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $cpu
    if ($status -eq 'Over-provisioned') { Add-Rec $row 'ContainerApps' 'Rightsize' 'Reduce CPU/memory allocation per replica (P95 usage well below allocation)' 'Smaller container resources' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Medium' $util -Rule 'RS-SKU' }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'ContainerApps' 'Performance Risk' 'Replicas running close to allocation - raise CPU/memory or max replicas' 'Scale up/out' $null 'Medium' $util -Rule 'PERF-RISK' }
}

$Analyzers.ContainerAppEnvironments = { param($row)
    $ded = [int]((N $row.DedicatedProfiles) ?? 0); $minN = [int]((N $row.DedicatedMinNodes) ?? 0)
    Set-Analysis $row 'Not assessed' "Workload profiles: $($row.WorkloadProfiles)" 'Usage measured on member container apps' $null
    if ($ded -gt 0 -and $minN -gt 0) { Add-Rec $row 'ContainerAppEnvironments' 'Configure' "Dedicated workload profiles keep $minN node(s) always on - review minimum node counts vs. app usage" 'Lower minimum nodes / Consumption profile' $null 'Low' "Dedicated profiles: $ded, minimum nodes: $minN" -Rule 'AUTOSCALE' }
}

$Analyzers.ContainerInstances = { param($row)
    $cpu = N $row.CpuUtilPct_P95; $mem = N $row.MemUtilPct_P95
    $prov = "$($row.CpuCores) vCPU / $($row.MemoryGB) GB, $($row.ContainerCount) container(s), $($row.OsType)"
    $util = "CPU P95 $(Fmt $cpu)% of requested; memory P95 $(Fmt $mem)%; state $($row.State)"
    if ($row.State -and $row.State -ne 'Running') { Set-Analysis $row 'Stopped' $prov $util $null 'Not billed while stopped'; return }
    if ($null -eq $cpu) { Set-Analysis $row 'No data' $prov $util $null; return }
    $status = if ($cpu -lt 10 -and ($mem ?? 0) -lt 30) { 'Over-provisioned' } elseif ($cpu -gt $CpuHighThreshold) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $cpu
    if ($status -eq 'Over-provisioned') { Add-Rec $row 'ContainerInstances' 'Rightsize' 'Reduce requested CPU/memory for the container group' 'Smaller resource requests' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Medium' $util -Rule 'RS-SKU' }
}

$Analyzers.ContainerRegistries = { param($row)
    $pulls = N $row.Pulls_Total; $pushes = N $row.Pushes_Total; $used = N $row.StorageUsedGB_Max; $inc = N $row.IncludedStorageGB
    $prov = "$($row.RegistrySku) (includes $(Fmt $inc 0) GB), geo-replications $($row.GeoReplications)"
    $util = "Storage used $(Fmt $used) GB ($(Fmt (Get-Pct $used $inc))% of included); pulls $(Fmt $pulls 0), pushes $(Fmt $pushes 0)"
    if (($pulls ?? 0) -eq 0 -and ($pushes ?? 0) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial', 'NoData')) {
        Set-Analysis $row 'Unused' $prov $util (Get-Pct $used $inc)
        Add-Rec $row 'ContainerRegistries' 'Delete' "No image pulls or pushes in $LookbackDays days - validate and delete the registry" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    if ($row.RegistrySku -eq 'Premium' -and [int]((N $row.GeoReplications) ?? 0) -le 1 -and ($used ?? 0) -lt 400) {
        Set-Analysis $row 'Over-provisioned' $prov $util (Get-Pct $used $inc)
        Add-Rec $row 'ContainerRegistries' 'Configure' 'Premium tier without geo-replication and < 400 GB stored - move to Standard if private link/content trust are not required' 'Standard SKU' $(if ($row.MonthlyCost) { [Math]::Round($row.MonthlyCost * 0.6, 2) } else { 30 }) 'Low' $util -Rule 'RS-SKU'
        return
    }
    Set-Analysis $row 'Right-sized' $prov $util (Get-Pct $used $inc)
}

function Get-Pct($a, $b) { if ($null -eq $a -or $null -eq $b -or [double]$b -eq 0) { return $null }; return [Math]::Round(100.0 * [double]$a / [double]$b, 1) }

$AspLadder = @{ 'P3v3' = 'P2v3'; 'P2v3' = 'P1v3'; 'P1v3' = 'P0v3'; 'P5mv3' = 'P4mv3'; 'P4mv3' = 'P3mv3'; 'P3mv3' = 'P2mv3'; 'P2mv3' = 'P1mv3'
    'P3v2' = 'P2v2'; 'P2v2' = 'P1v2'; 'S3' = 'S2'; 'S2' = 'S1'; 'B3' = 'B2'; 'B2' = 'B1'; 'I6v2' = 'I5v2'; 'I5v2' = 'I4v2'; 'I4v2' = 'I3v2'; 'I3v2' = 'I2v2'; 'I2v2' = 'I1v2'
    'EP3' = 'EP2'; 'EP2' = 'EP1'; 'WS3' = 'WS2'; 'WS2' = 'WS1'
}
$AspModernize = @{ 'P1v2' = 'P0v3'; 'P2v2' = 'P1v3'; 'P3v2' = 'P2v3'; 'S1' = 'P0v3'; 'S2' = 'P1v3'; 'S3' = 'P2v3' }

$Analyzers.AppServicePlans = { param($row)
    $tier = [string]$row.PlanTier; $sku = [string]$row.PlanSku; $n = [int]((N $row.Instances) ?? 1); $sites = [int]((N $row.NumberOfSites) ?? 0)
    $cpuAvg = N $row.CPU_Avg; $cpuP95 = N $row.CPU_P95; $cpuMax = N $row.CPU_Max; $memP95 = N $row.Memory_P95; $memMax = N $row.Memory_Max
    $prov = "$sku ($tier) x $n instance(s), $sites app(s), $($row.OsType)"
    $util = if ($null -ne $cpuP95) { "CPU avg $(Fmt $cpuAvg)% | P95 $(Fmt $cpuP95)% | max $(Fmt $cpuMax)%; memory P95 $(Fmt $memP95)% | max $(Fmt $memMax)%" } else { 'No metrics' }
    if ($tier -in @('Free', 'Shared', 'Dynamic', 'FlexConsumption')) { Set-Analysis $row 'Not assessed' $prov 'Consumption / free plan (billed per use)' $null; return }
    $hourly = Get-AppHourly $row.Location $sku $row.OsType
    $retailMonthly = if ($hourly) { [Math]::Round($hourly * $n * 730, 2) } else { $null }
    $mc = if ($row.MonthlyCost) { $row.MonthlyCost } else { $retailMonthly }
    $apps = $script:AppsByPlan[([string]$row.ResourceId).ToLower()]
    if ($sites -eq 0) {
        Set-Analysis $row 'Unused' $prov 'Plan hosts no apps' 0
        Add-Rec $row 'AppServicePlans' 'Delete' 'Empty App Service plan (no apps) - delete' 'Delete' $mc 'High' "0 apps on plan; $n instance(s) billed" -Rule 'ORPHAN'
        return
    }
    if ($apps -and $apps.Total -gt 0 -and $apps.Stopped -eq $apps.Total) {
        Set-Analysis $row 'Unused' $prov "All $($apps.Total) app(s) on the plan are stopped" $cpuP95
        Add-Rec $row 'AppServicePlans' 'Delete' "All apps on the plan are stopped ($($apps.Names)) - delete apps and plan, or consolidate" 'Delete' $mc 'Medium' 'Stopped apps still incur the plan cost' -Rule 'ORPHAN'
        return
    }
    $status = Get-StatusFromCpu $cpuAvg $cpuP95 $memP95 $memMax 85
    Set-Analysis $row $status $prov $util ([Math]::Max([double]($cpuP95 ?? 0), [double]($memP95 ?? 0)))
    $evidence = "CPU P95 $(Fmt $cpuP95)%, memory P95 $(Fmt $memP95)% on $n instance(s) (past $LookbackDays days)"
    if ($status -in @('Idle', 'Over-provisioned')) {
        if ($n -gt 1) {
            $floor = if (B $row.'Prop.zoneRedundant') { 2 } else { 1 }
            $target = [int][Math]::Max($floor, [Math]::Ceiling($n * [Math]::Max([double]$cpuP95 / $TargetUtilization, [double]($memP95 ?? 0) / 80.0)))
            if ($target -lt $n) { Add-Rec $row 'AppServicePlans' 'Rightsize' "Scale in from $n to $target instance(s)$(if (-not (B $row.AutoscaleEnabled)) { ' and configure autoscale' })" "$sku x $target" (Get-Savings -MonthlyCost $mc -Ratio ($target / [double]$n)) 'Medium' $evidence -Rule 'RS-SCALE' }
        }
        elseif ($AspLadder[$sku] -and $cpuP95 -lt 35 -and ($memP95 ?? 0) -lt 50) {
            $tSku = $AspLadder[$sku]; $tHourly = Get-AppHourly $row.Location $tSku $row.OsType
            Add-Rec $row 'AppServicePlans' 'Rightsize' "Scale down from $sku to $tSku" $tSku (Get-Savings -MonthlyCost $mc -CurrentUnit $hourly -TargetUnit $tHourly -HoursPerMonth (730 * $n)) 'Medium' $evidence -Rule 'RS-SKU'
        }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'AppServicePlans' 'Performance Risk' 'Plan running hot - scale out (autoscale) or scale up' 'Scale out/up' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ($AspModernize[$sku]) {
        $mSku = $AspModernize[$sku]; $mHourly = Get-AppHourly $row.Location $mSku $row.OsType
        $sav = Get-Savings -MonthlyCost $mc -CurrentUnit $hourly -TargetUnit $mHourly -HoursPerMonth (730 * $n)
        if ($null -eq $sav -or $sav -gt 0) { Add-Rec $row 'AppServicePlans' 'Configure' "Move from $sku to Premium v3 ($mSku) - better price/performance and eligible for reservations/savings plan" $mSku $sav 'Low' "Current SKU $sku (previous generation)" -Rule 'MODERNIZE' }
    }
}

$Analyzers.AppServices = { param($row)
    $req = N $row.Requests_Total; $exec = N $row.Executions_Total; $e5 = N $row.Http5xx_Total; $rt = N $row.ResponseTimeSec_Avg
    $prov = "$($row.AppType) on plan '$($row.AppServicePlan)'"
    $util = "Requests $(Fmt $req 0) ($(Fmt ($(if ($null -ne $req) { $req / $LookbackDays } else { $null })) 0)/day); HTTP 5xx $(Fmt $e5 0); avg response $(Fmt $rt 2)s$(if ($null -ne $exec) { "; executions $(Fmt $exec 0)" })"
    if ($row.State -eq 'Stopped') {
        Set-Analysis $row 'Stopped' $prov 'App is stopped' $null
        Add-Rec $row 'AppServices' 'Delete' 'App is stopped - delete if no longer required (frees plan capacity)' 'Delete' $null 'Medium' 'State = Stopped' -Rule 'NOACT'
        return
    }
    if ($row.MetricsStatus -notin @('OK', 'Partial', 'NoData')) { Set-Analysis $row 'No data' $prov $util $null; return }
    if (($req ?? 0) -eq 0 -and ($exec ?? 0) -eq 0) {
        Set-Analysis $row 'Unused' $prov $util 0
        Add-Rec $row 'AppServices' 'Delete' "No requests or executions in $LookbackDays days - validate with owner and remove" 'Delete' $null 'Medium' $util -Rule 'NOACT'
        return
    }
    Set-Analysis $row 'In use' $prov $util $null 'Sizing is assessed at App Service plan level'
}

#endregion

#region ============================== ANALYZERS: DATABASES ==============================

$DtuLadder = @{
    Basic    = @(@{ N = 'Basic'; D = 5 })
    Standard = @(@{ N = 'S0'; D = 10 }, @{ N = 'S1'; D = 20 }, @{ N = 'S2'; D = 50 }, @{ N = 'S3'; D = 100 }, @{ N = 'S4'; D = 200 }, @{ N = 'S6'; D = 400 }, @{ N = 'S7'; D = 800 }, @{ N = 'S9'; D = 1600 }, @{ N = 'S12'; D = 3000 })
    Premium  = @(@{ N = 'P1'; D = 125 }, @{ N = 'P2'; D = 250 }, @{ N = 'P4'; D = 500 }, @{ N = 'P6'; D = 1000 }, @{ N = 'P11'; D = 1750 }, @{ N = 'P15'; D = 4000 })
}
# Indicative monthly list prices (USD) used only for the price ratio between DTU tiers
$DtuPrice = @{ Basic = 4.99; S0 = 14.72; S1 = 29.43; S2 = 73.58; S3 = 147.17; S4 = 294.34; S6 = 588.68; S7 = 1177.36; S9 = 2354.72; S12 = 4415.10; P1 = 456.25; P2 = 912.50; P4 = 1825; P6 = 3650; P11 = 6862; P15 = 15695 }
$VCoreLadder = [double[]]@(2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 24, 32, 40, 80, 128)
$MiVCoreLadder = [double[]]@(4, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 128)
$FlexVCoreLadder = [double[]]@(2, 4, 8, 16, 20, 32, 48, 64, 96)

$Analyzers.SqlDatabases = { param($row)
    $model = [string]$row.PurchaseModel; $cap = N $row.Capacity; $conns = N $row.Connections_Total
    $p95 = if ($model -eq 'DTU') { N $row.DTU_P95 } elseif ($model -eq 'DWU') { N $row.DWU_P95 } else { N $row.CPU_P95 }
    $avg = if ($model -eq 'DTU') { N $row.DTU_Avg } elseif ($model -eq 'DWU') { N $row.DWU_Avg } else { N $row.CPU_Avg }
    $mx = if ($model -eq 'DTU') { N $row.DTU_Max } elseif ($model -eq 'DWU') { N $row.DWU_Max } else { N $row.CPU_Max }
    $memP95 = N $row.Memory_P95; $used = N $row.StorageUsedGB_Max; $maxGb = N $row.MaxSizeGB
    $unit = switch ($model) { 'DTU' { 'DTU' } 'DWU' { 'DWU' } default { 'vCores' } }
    $prov = if ($row.ElasticPool) { "In elastic pool '$($row.ElasticPool)'" } else { "$($row.ServiceObjective) ($($row.Edition), $(Fmt $cap 0) $unit$(if ($row.ComputeTier -eq 'Serverless') { ' max, serverless' })), max size $(Fmt $maxGb 0) GB" }
    $util = "$(if ($model -eq 'DTU') { 'DTU' } else { 'CPU' }) avg $(Fmt $avg)% | P95 $(Fmt $p95)% | max $(Fmt $mx)%; storage used $(Fmt $used) GB ($(Fmt (Get-Pct $used $maxGb))%); connections $(Fmt $conns 0)"
    $evidence = "$util (past $LookbackDays days)"
    $isSecondary = -not [string]::IsNullOrEmpty([string]$row.SecondaryType)
    if (($conns ?? -1) -eq 0 -and -not $isSecondary -and $row.MetricsStatus -in @('OK', 'Partial')) {
        Set-Analysis $row 'Unused' $prov $util $p95
        Add-Rec $row 'SqlDatabases' 'Delete' "No successful connections in $LookbackDays days - validate, export (BACPAC) and delete$(if ($row.ElasticPool) { ' (frees elastic pool capacity)' })" 'Delete' $(if ($row.ElasticPool) { $null } else { $row.MonthlyCost }) 'Medium' $evidence -Rule 'NOACT'
        return
    }
    if ($row.ElasticPool) { Set-Analysis $row 'In use' $prov $util $p95 'Sized at elastic pool level'; return }
    $status = Get-StatusFromCpu $avg $p95 $memP95 $null
    Set-Analysis $row $status $prov $util $p95
    if ($status -in @('Idle', 'Over-provisioned') -and $null -ne $cap) {
        if ($model -eq 'DTU' -and $DtuLadder[[string]$row.Edition]) {
            $req = $cap * $p95 / $TargetUtilization
            $t = @($DtuLadder[[string]$row.Edition] | Where-Object { $_.D -ge $req -and $_.D -lt $cap }) | Select-Object -First 1
            if ($t -and $DtuPrice[$row.ServiceObjective]) {
                Add-Rec $row 'SqlDatabases' 'Rightsize' "Move from $($row.ServiceObjective) ($(Fmt $cap 0) DTU) to $($t.N) ($($t.D) DTU)" $t.N (Get-Savings -MonthlyCost $row.MonthlyCost -CurrentUnit $DtuPrice[$row.ServiceObjective] -TargetUnit $DtuPrice[$t.N] -HoursPerMonth 1) 'High' $evidence -Rule 'RS-SKU'
            }
            elseif (-not $t) { $row | Add-Member -NotePropertyName AnalysisNote -NotePropertyValue 'Already at the smallest size of its tier - consider serverless or consolidating into an elastic pool' -Force }
        }
        elseif ($model -eq 'vCore' -and $row.ComputeTier -eq 'Provisioned') {
            $t = Get-LadderTarget $VCoreLadder $cap ([Math]::Max(2, [Math]::Ceiling($cap * $p95 / $TargetUtilization)))
            if ($t) { Add-Rec $row 'SqlDatabases' 'Rightsize' "Decrease vCores from $(Fmt $cap 0) to $t" "$t vCores" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $cap) -Share 0.85) 'High' $evidence -Rule 'RS-SKU' }
        }
        elseif ($row.ComputeTier -eq 'Serverless') {
            Add-Rec $row 'SqlDatabases' 'Configure' "Serverless max vCores ($(Fmt $cap 0)) far above usage - lower max vCores / min vCores and keep auto-pause enabled" 'Lower max vCores' $null 'Low' $evidence -Rule 'RS-SKU'
        }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'SqlDatabases' 'Performance Risk' "Sustained high $(if ($model -eq 'DTU') { 'DTU' } else { 'CPU' }) utilization - scale up or tune top queries" 'Scale up / query tuning' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ((N $row.Sessions_Max) -ge 90) { Add-Rec $row 'SqlDatabases' 'Performance Risk' "Concurrent sessions peaked at $(Fmt (N $row.Sessions_Max) 0)% of the service-tier limit - risk of connection failures" 'Higher tier / connection pooling' $null 'Medium' "sessions_percent max $(Fmt (N $row.Sessions_Max) 0)%; failed connections $(Fmt (N $row.ConnectionsFailed_Total) 0)" -Rule 'PERF-RISK' }
    if ((N $row.Deadlocks_Total) -gt 0) { Add-Rec $row 'SqlDatabases' 'Performance Risk' "$(Fmt (N $row.Deadlocks_Total) 0) deadlock(s) in $LookbackDays days - review indexing / transaction design" 'Query & index tuning' $null 'Low' 'deadlock metric > 0' -Rule 'PERF-RISK' }
    if ($model -eq 'vCore' -and $row.Edition -in @('GeneralPurpose', 'BusinessCritical') -and $maxGb -gt 64 -and $null -ne $used -and $used / $maxGb -lt 0.3) {
        $tGb = [Math]::Max(32, [Math]::Ceiling($used * 1.5 / 32) * 32)
        $perGb = if ($row.Edition -eq 'BusinessCritical') { 0.25 } else { 0.115 }
        Add-Rec $row 'SqlDatabases' 'Configure' "Reduce max data size from $(Fmt $maxGb 0) GB to $tGb GB (vCore storage is billed on max size)" "$tGb GB max size" ([Math]::Round(($maxGb - $tGb) * $perGb, 2)) 'Medium' "Storage used $(Fmt $used) GB of $(Fmt $maxGb 0) GB" -Rule 'RS-STORAGE'
    }
    if ($model -eq 'vCore' -and $row.LicenseType -eq 'LicenseIncluded' -and $row.ComputeTier -eq 'Provisioned' -and $row.Edition -ne 'Hyperscale') {
        Add-Rec $row 'SqlDatabases' 'Configure' 'Apply Azure Hybrid Benefit for SQL Server (requires SQL licenses with Software Assurance)' 'LicenseType = BasePrice' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.7) 'Low' 'License included in the vCore price' 'Low' 'Azure Hybrid Benefit for SQL: https://learn.microsoft.com/azure/azure-sql/azure-hybrid-benefit' -Rule 'LICENSE'
    }
    if ($row.Environment -eq 'Non-Production' -and $row.Edition -eq 'BusinessCritical') {
        Add-Rec $row 'SqlDatabases' 'Configure' 'Non-production database on Business Critical - use General Purpose' 'General Purpose' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.4) 'Medium' 'Business Critical tier in non-production' -Rule 'TIER-NONPROD'
    }
}

$Analyzers.SqlElasticPools = { param($row)
    $model = [string]$row.PurchaseModel; $cap = N $row.Capacity; $dbs = [int]((N $row.DatabaseCount) ?? 0)
    $p95 = if ($model -eq 'DTU') { N $row.DTU_P95 } else { N $row.CPU_P95 }; $avg = if ($model -eq 'DTU') { N $row.DTU_Avg } else { N $row.CPU_Avg }; $mx = if ($model -eq 'DTU') { N $row.DTU_Max } else { N $row.CPU_Max }
    $prov = "$($row.PoolSku) ($($row.Edition), $(Fmt $cap 0) $(if ($model -eq 'DTU') { 'eDTU' } else { 'vCores' })), $dbs database(s)"
    $util = "$(if ($model -eq 'DTU') { 'eDTU' } else { 'CPU' }) avg $(Fmt $avg)% | P95 $(Fmt $p95)% | max $(Fmt $mx)%; storage $(Fmt (N $row.StoragePct_Max))% used"
    if ($dbs -eq 0) { Set-Analysis $row 'Unused' $prov 'Pool contains no databases' 0; Add-Rec $row 'SqlElasticPools' 'Delete' 'Elastic pool without databases - delete' 'Delete' $row.MonthlyCost 'High' '0 databases in pool' -Rule 'ORPHAN'; return }
    $status = Get-StatusFromCpu $avg $p95 $null $null
    Set-Analysis $row $status $prov $util $p95
    $evidence = "$util (past $LookbackDays days)"
    if ($status -in @('Idle', 'Over-provisioned') -and $null -ne $cap) {
        $ladder = if ($model -eq 'DTU') { if ($row.Edition -eq 'Premium') { [double[]]@(125, 250, 500, 1000, 1500, 2000, 2500, 3000, 3500, 4000) } else { [double[]]@(50, 100, 200, 300, 400, 800, 1200, 1600, 2000, 2500, 3000) } } else { $VCoreLadder }
        $t = Get-LadderTarget $ladder $cap ([Math]::Max($ladder[0], [Math]::Ceiling($cap * $p95 / $TargetUtilization)))
        if ($t) { Add-Rec $row 'SqlElasticPools' 'Rightsize' "Decrease pool from $(Fmt $cap 0) to $t $(if ($model -eq 'DTU') { 'eDTU' } else { 'vCores' })" "$t" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $cap) -Share 0.9) 'High' $evidence -Rule 'RS-SKU' }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'SqlElasticPools' 'Performance Risk' 'Pool running hot - increase pool capacity or rebalance databases' 'Scale up' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ($dbs -eq 1) { Add-Rec $row 'SqlElasticPools' 'Configure' 'Pool hosts a single database - a single database is usually cheaper' 'Single database' $null 'Low' '1 database in pool' -Rule 'CONSOLIDATE' }
}

$Analyzers.SqlManagedInstances = { param($row)
    $vc = N $row.vCores; $p95 = N $row.CPU_P95; $avg = N $row.CPU_Avg; $mx = N $row.CPU_Max; $used = N $row.StorageUsedGB_Max; $res = N $row.ReservedStorageGB_Max; if ($null -eq $res) { $res = N $row.StorageSizeGB }
    $prov = "$($row.MiSku) ($($row.Tier), $(Fmt $vc 0) vCores, $(Fmt $res 0) GB storage, $($row.LicenseType))"
    $util = "CPU avg $(Fmt $avg)% | P95 $(Fmt $p95)% | max $(Fmt $mx)%; storage used $(Fmt $used) GB ($(Fmt (Get-Pct $used $res))%)"
    if ($row.State -eq 'Stopped') { Set-Analysis $row 'Stopped' $prov 'Instance stopped' $null; return }
    $status = Get-StatusFromCpu $avg $p95 $null $null
    Set-Analysis $row $status $prov $util $p95
    $evidence = "$util (past $LookbackDays days)"
    if ($status -in @('Idle', 'Over-provisioned') -and $vc) {
        $t = Get-LadderTarget $MiVCoreLadder $vc ([Math]::Max(4, [Math]::Ceiling($vc * $p95 / $TargetUtilization)))
        if ($t) { Add-Rec $row 'SqlManagedInstances' 'Rightsize' "Decrease vCores from $(Fmt $vc 0) to $t" "$t vCores" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $vc) -Share 0.85) 'High' $evidence -Rule 'RS-SKU' }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'SqlManagedInstances' 'Performance Risk' 'Managed instance running hot - add vCores or tune workload' 'Scale up' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ($null -ne $used -and $res -gt 256 -and $used / $res -lt 0.4) {
        $t = [Math]::Max(32, [Math]::Ceiling($used * 1.4 / 32) * 32)
        Add-Rec $row 'SqlManagedInstances' 'Configure' "Reduce reserved storage from $(Fmt $res 0) GB to $t GB" "$t GB" ([Math]::Round(($res - $t) * 0.115, 2)) 'Medium' $util -Rule 'RS-STORAGE'
    }
    if ($row.LicenseType -eq 'LicenseIncluded') { Add-Rec $row 'SqlManagedInstances' 'Configure' 'Apply Azure Hybrid Benefit for SQL Server (requires licenses with SA)' 'BasePrice' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.7) 'Low' 'License included' -Rule 'LICENSE' }
    if ($row.Environment -eq 'Non-Production') { Add-Rec $row 'SqlManagedInstances' 'Configure' 'Non-production managed instance running 24x7 - use the stop/start schedule' 'Scheduled stop/start' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5 -Share 0.85) 'Medium' 'Non-production instance' 'Low' 'Stop/start MI: https://learn.microsoft.com/azure/azure-sql/managed-instance/instance-stop-start-how-to' -Rule 'SCHEDULE' }
}

$Analyzers.SqlVirtualMachines = { param($row)
    $vm = $script:VmById[([string]$row.VirtualMachineId).ToLower()]
    $p95 = if ($vm) { N $vm.CPU_P95 } else { $null }
    $prov = "SQL Server $($row.SqlEdition) ($($row.SqlImageOffer)), license $($row.SqlLicenseType) on VM $($row.VirtualMachine)"
    Set-Analysis $row $(if ($vm) { $vm.UtilizationStatus } else { 'Not assessed' }) $prov $(if ($vm) { $vm.UtilizedSummary } else { 'VM not in scope' }) $p95
    if ($row.SqlLicenseType -eq 'PAYG' -and $row.SqlEdition -in @('Enterprise', 'Standard')) {
        Add-Rec $row 'SqlVirtualMachines' 'Configure' "SQL Server $($row.SqlEdition) billed pay-as-you-go - apply Azure Hybrid Benefit (AHUB) or the free DR license for passive replicas" 'AHUB / DR license' $null 'Low' 'SQL license included in VM price' 'Low' 'SQL VM licensing: https://learn.microsoft.com/azure/azure-sql/virtual-machines/windows/licensing-model-azure-hybrid-benefit-ahb-change' -Rule 'LICENSE'
    }
    if ($row.SqlEdition -eq 'Enterprise' -and $null -ne $p95 -and $p95 -lt $CpuLowThreshold) {
        Add-Rec $row 'SqlVirtualMachines' 'Configure' 'Low-utilized SQL Enterprise - validate whether Enterprise-only features are needed (Standard edition costs ~70% less)' 'Standard edition' $null 'Low' "VM CPU P95 $(Fmt $p95)%" -Rule 'LICENSE'
    }
}

$FlexAnalyzer = { param($row, [string]$Key)
    $vc = N $row.vCores; $avg = N $row.CPU_Avg; $p95 = N $row.CPU_P95; $mx = N $row.CPU_Max; $memP95 = N $row.Memory_P95; $memMax = N $row.Memory_Max; $conn = N $row.Connections_Max
    $prov = "$($row.ServerSku) ($($row.Tier), $(Fmt $vc 0) vCores), $($row.StorageGB) GB storage, HA $($row.HAMode)"
    $util = "CPU avg $(Fmt $avg)% | P95 $(Fmt $p95)% | max $(Fmt $mx)%; memory P95 $(Fmt $memP95)%; storage $(Fmt (N $row.StoragePct_Max))% used; max connections $(Fmt $conn 0)"
    if ($row.State -eq 'Stopped') { Set-Analysis $row 'Stopped' $prov 'Server stopped (compute not billed; restarts automatically after 7 days)' $null; return }
    if (($conn ?? -1) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial')) {
        Set-Analysis $row 'Unused' $prov $util $p95
        Add-Rec $row $Key 'Delete' "No client connections in $LookbackDays days - validate and delete (or stop)" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    $status = Get-StatusFromCpu $avg $p95 $memP95 $memMax 70
    Set-Analysis $row $status $prov $util $p95
    $evidence = "$util (past $LookbackDays days)"
    if ($status -in @('Idle', 'Over-provisioned') -and $vc -gt 2 -and $row.Tier -ne 'Burstable') {
        $t = Get-LadderTarget $FlexVCoreLadder $vc ([Math]::Max(2, [Math]::Ceiling($vc * [Math]::Max($p95 / $TargetUtilization, [double]($memP95 ?? 0) / 80.0))))
        if ($t) { Add-Rec $row $Key 'Rightsize' "Decrease compute from $(Fmt $vc 0) to $t vCores" "$t vCores ($($row.Tier))" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $vc) -Share 0.85) 'High' $evidence -Rule 'RS-SKU' }
    }
    if ($status -eq 'Idle' -and $row.Environment -eq 'Non-Production' -and $row.Tier -eq 'GeneralPurpose') {
        Add-Rec $row $Key 'Configure' 'Idle non-production server - move to the Burstable tier (or stop when not in use)' 'Burstable B2s/B2ms' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.3) 'Medium' $evidence -Rule 'TIER-NONPROD'
    }
    if ($status -eq 'Under-provisioned') { Add-Rec $row $Key 'Performance Risk' 'Server running hot - scale up compute' 'Scale up' $null 'Medium' $evidence -Rule 'PERF-RISK' }
    if ($row.Environment -eq 'Non-Production' -and $row.HAMode -in @('ZoneRedundant', 'SameZone')) {
        Add-Rec $row $Key 'Configure' 'High availability enabled on a non-production server - disable HA' 'HA disabled' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Medium' "HA mode $($row.HAMode)" -Rule 'TIER-NONPROD'
    }
}
$Analyzers.MySQLFlexibleServers = { param($row) & $FlexAnalyzer $row 'MySQLFlexibleServers' }
$Analyzers.PostgreSQLFlexibleServers = { param($row) & $FlexAnalyzer $row 'PostgreSQLFlexibleServers' }

$Analyzers.LegacyDatabaseServers = { param($row)
    $p95 = N $row.CPU_P95
    Set-Analysis $row (Get-StatusFromCpu (N $row.CPU_Avg) $p95 (N $row.Memory_P95) $null) "$($row.ServerSku) ($($row.Tier), $($row.vCores) vCores)" "CPU P95 $(Fmt $p95)%" $p95
    Add-Rec $row 'LegacyDatabaseServers' 'Configure' 'Single Server is retired - migrate to Flexible Server (right-size during migration)' 'Flexible Server' $null 'High' "Engine: $($row.ArmType)" 'Medium' -Rule 'MODERNIZE'
}

$Analyzers.CosmosDB = { param($row)
    $nru = N $row.NormalizedRUPct_P95; $nruMax = N $row.NormalizedRUPct_Max; $prov = N $row.ProvisionedRUps; $avgRu = N $row.AvgConsumedRUps; $req = N $row.Requests_Total
    $provTxt = "$($row.ApiType), $($row.CapacityMode)$(if ($null -ne $prov) { ", $(Fmt $prov 0) RU/s provisioned" }), $($row.RegionCount) region(s)$(if (B $row.MultiRegionWrites) { ', multi-region writes' })"
    $util = "Normalized RU P95 $(Fmt $nru)% | max $(Fmt $nruMax)%; avg consumed $(Fmt $avgRu 0) RU/s$(if ($null -ne $prov) { " ($(Fmt (Get-Pct $avgRu $prov))% of provisioned)" }); requests $(Fmt $req 0); data $(Fmt (N $row.DataGB_Max)) GB"
    if (($req ?? -1) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial')) {
        Set-Analysis $row 'Unused' $provTxt $util 0
        Add-Rec $row 'CosmosDB' 'Delete' "No requests in $LookbackDays days - validate and delete the account" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    if ($row.CapacityMode -eq 'Serverless') { Set-Analysis $row 'In use' $provTxt $util $null 'Serverless - billed per request'; return }
    if ($null -eq $nru) { Set-Analysis $row 'No data' $provTxt $util $null; return }
    $status = if ($nru -lt 10) { 'Idle' } elseif ($nru -lt 30) { 'Over-provisioned' } elseif ($nru -gt 90) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $provTxt $util $nru
    if ($status -in @('Idle', 'Over-provisioned') -and $prov -gt 400) {
        $t = [Math]::Max(400, [Math]::Ceiling($prov * $nru / $TargetUtilization / 100) * 100)
        if ($t -lt $prov) { Add-Rec $row 'CosmosDB' 'Rightsize' "Reduce provisioned throughput from $(Fmt $prov 0) to ~$t RU/s (or switch to autoscale)" "$t RU/s / autoscale" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $prov) -Share 0.8) 'Medium' $util -Rule 'RS-SCALE' }
        if (($avgRu ?? 0) -lt 50) { Add-Rec $row 'CosmosDB' 'Configure' 'Very low sustained RU consumption - evaluate serverless capacity mode' 'Serverless' $null 'Low' $util -Rule 'SERVERLESS' }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'CosmosDB' 'Performance Risk' 'Normalized RU consumption near 100% - risk of throttling (429); increase RU/s or review partition key' 'More RU/s / autoscale' $null 'Medium' $util -Rule 'PERF-RISK' }
    if ((N $row.ThrottledPct_P95) -ge 1 -or (N $row.ThrottledPct_Max) -ge 5) { Add-Rec $row 'CosmosDB' 'Performance Risk' "Requests throttled (HTTP 429): P95 $(Fmt (N $row.ThrottledPct_P95))% / max $(Fmt (N $row.ThrottledPct_Max))% of requests - hot partitions or insufficient RU/s" 'Partition review / autoscale' $null 'Medium' "ThrottledRequestPercentage P95 $(Fmt (N $row.ThrottledPct_P95))%, max $(Fmt (N $row.ThrottledPct_Max))%" -Rule 'PERF-RISK' }
    if ($row.Environment -eq 'Non-Production' -and (B $row.MultiRegionWrites)) { Add-Rec $row 'CosmosDB' 'Configure' 'Multi-region writes on non-production - use a single write region' 'Single region' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Medium' 'Multi-region writes enabled' -Rule 'TIER-NONPROD' }
}

$RedisLadder = @{ 'C6' = 'C5'; 'C5' = 'C4'; 'C4' = 'C3'; 'C3' = 'C2'; 'C2' = 'C1'; 'P5' = 'P4'; 'P4' = 'P3'; 'P3' = 'P2'; 'P2' = 'P1' }
$RedisHourly = @{ 'Standard|C0' = 0.055; 'Standard|C1' = 0.137; 'Standard|C2' = 0.226; 'Standard|C3' = 0.454; 'Standard|C4' = 0.535; 'Standard|C5' = 1.07; 'Standard|C6' = 2.14
    'Basic|C0' = 0.022; 'Basic|C1' = 0.055; 'Basic|C2' = 0.09; 'Basic|C3' = 0.18; 'Basic|C4' = 0.21; 'Basic|C5' = 0.42; 'Basic|C6' = 0.84
    'Premium|P1' = 0.554; 'Premium|P2' = 1.11; 'Premium|P3' = 2.22; 'Premium|P4' = 4.44; 'Premium|P5' = 8.33 }

$Analyzers.RedisCaches = { param($row)
    $load = N $row.ServerLoad_P95; $mem = N $row.MemoryUsedPct_Max; $cmd = N $row.Commands_Total; $cli = N $row.ConnectedClients_Max
    $prov = "$($row.CacheSku) $($row.CacheSize) ($(Fmt (N $row.MemoryGBPerShard)) GB$(if ((N $row.ShardCount) -gt 1) { " x $($row.ShardCount) shards" }))"
    $util = "Server load P95 $(Fmt $load)%; memory used max $(Fmt $mem)%; max clients $(Fmt $cli 0); commands $(Fmt $cmd 0)"
    if (($cmd ?? -1) -eq 0 -and ($cli ?? 0) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial')) {
        Set-Analysis $row 'Unused' $prov $util 0
        Add-Rec $row 'RedisCaches' 'Delete' "No commands or client connections in $LookbackDays days - delete the cache" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    if ($null -eq $load) { Set-Analysis $row 'No data' $prov $util $null; return }
    $status = if ($load -lt 30 -and ($mem ?? 0) -lt 40) { 'Over-provisioned' } elseif ($load -gt 80 -or ($mem ?? 0) -gt 90) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util ([Math]::Max($load, [double]($mem ?? 0)))
    if ($status -eq 'Over-provisioned' -and $RedisLadder[$row.CacheSize]) {
        $t = $RedisLadder[$row.CacheSize]; $c = $RedisHourly["$($row.CacheSku)|$($row.CacheSize)"]; $tp = $RedisHourly["$($row.CacheSku)|$t"]
        Add-Rec $row 'RedisCaches' 'Rightsize' "Scale down from $($row.CacheSize) to $t" "$($row.CacheSku) $t" (Get-Savings -MonthlyCost $row.MonthlyCost -CurrentUnit $c -TargetUnit $tp) 'Medium' $util -Rule 'RS-SKU'
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'RedisCaches' 'Performance Risk' 'Cache under pressure (server load / memory) - scale up or out' 'Scale up/out' $null 'Medium' $util -Rule 'PERF-RISK' }
    if ($row.CacheSku -eq 'Premium' -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'RedisCaches' 'Configure' 'Premium cache in non-production - use Standard tier' 'Standard tier' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.4) 'Medium' 'Premium tier in non-production' -Rule 'TIER-NONPROD' }
}

$Analyzers.ManagedRedis = { param($row) Set-Analysis $row 'Not assessed' "$($row.CacheSku) (capacity $($row.Capacity))" 'Configuration inventory only' $null }

#endregion

#region ============================== ANALYZERS: STORAGE ==============================

$StdSsdLimits = @{ 1 = @(500, 60); 2 = @(500, 60); 3 = @(500, 60); 4 = @(500, 60); 6 = @(500, 60); 10 = @(500, 60); 15 = @(500, 60); 20 = @(500, 60); 30 = @(500, 60); 40 = @(500, 60); 50 = @(500, 60); 60 = @(2000, 400); 70 = @(4000, 600); 80 = @(6000, 750) }

$Analyzers.ManagedDisks = { param($row)
    $tier = [string]$row.DiskTier; $sku = [string]$row.DiskSku; $red = ($sku -split '_')[-1]; $size = N $row.DiskSizeGB
    $piops = N $row.ProvisionedIOPS; $iopsP95 = N $row.TotalIOPS_P95; $iopsMax = N $row.TotalIOPS_Max; $mbP95 = N $row.TotalMBps_P95; $util = N $row.IOPSUtilPct_P95
    $prov = "$sku $tier, $(Fmt $size 0) GB, $(Fmt $piops 0) IOPS / $(Fmt (N $row.ProvisionedMBps) 0) MBps"
    $utilTxt = if ($null -ne $iopsP95) { "IOPS P95 $(Fmt $iopsP95 0) ($(Fmt $util)% of provisioned) | max $(Fmt $iopsMax 0); throughput P95 $(Fmt $mbP95) MBps" } else { "State $($row.DiskState)" }
    if ($row.DiskState -eq 'Unattached') {
        Set-Analysis $row 'Orphaned' $prov "Unattached$(if ($row.DaysUnattached) { " for $($row.DaysUnattached) days" })" 0
        Add-Rec $row 'ManagedDisks' 'Delete' "Unattached disk$(if ($row.DaysUnattached) { " (detached $($row.DaysUnattached) days ago)" }) - snapshot if needed and delete" 'Delete' $row.MonthlyCost 'High' "Disk state Unattached; $prov" -Rule 'ORPHAN'
        return
    }
    if ($row.DiskState -eq 'Reserved') { Set-Analysis $row 'Deallocated' $prov 'Attached to a deallocated VM (still billed)' $null; return }
    if ($null -eq $iopsP95) { Set-Analysis $row 'No data' $prov $utilTxt $null; return }
    $status = if ($null -ne $util -and $util -lt 10) { 'Over-provisioned' } elseif ($null -ne (N $row.IOPSUtilPct_Max) -and (N $row.IOPSUtilPct_Max) -ge 95) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $utilTxt $util
    if ($sku -like 'Premium_*' -and $tier -match '^P(\d+)$') {
        $n = [int]$Matches[1]; $lim = $StdSsdLimits[$n]
        if ($lim -and $iopsP95 -lt $lim[0] * 0.8 -and ($mbP95 ?? 0) -lt $lim[1] * 0.8 -and $status -ne 'Under-provisioned') {
            $pP = Get-DiskMonthly $row.Location $tier $red; $pE = Get-DiskMonthly $row.Location "E$n" $red
            $isSqlVm = $script:SqlVmIds.Contains(([string]$row.AttachedToId).ToLower())
            Add-Rec $row 'ManagedDisks' 'Rightsize' "Change Premium SSD $tier to Standard SSD E$n (P95 $(Fmt $iopsP95 0) IOPS fits the E$n limit of $($lim[0]) IOPS)" "StandardSSD_$red E$n" (Get-Savings -MonthlyCost $row.MonthlyCost -CurrentUnit $pP -TargetUnit $pE -HoursPerMonth 1) $(if ($isSqlVm -or (B $row.IsOsDisk)) { 'Low' } else { 'Medium' }) "$utilTxt$(if ($isSqlVm) { '; attached to a SQL Server VM - validate latency requirements' })" -Rule 'DISK-TIER'
        }
    }
    if ($sku -match 'UltraSSD|PremiumV2' -and $piops -gt 3000 -and ($iopsMax ?? 0) -lt $piops * 0.3) {
        $t = [Math]::Max(3000, [Math]::Ceiling($iopsMax * 1.3 / 100) * 100)
        Add-Rec $row 'ManagedDisks' 'Rightsize' "Reduce provisioned IOPS from $(Fmt $piops 0) to $t" "$t IOPS" $null 'Medium' $utilTxt -Rule 'RS-STORAGE'
    }
    if ($status -eq 'Under-provisioned') { Add-Rec $row 'ManagedDisks' 'Performance Risk' 'Disk reaching its IOPS limit - consider a higher tier / bursting' 'Higher performance tier' $null 'Medium' $utilTxt -Rule 'PERF-RISK' }
}

$Analyzers.Snapshots = { param($row)
    $age = N $row.AgeDays; $orph = ([string]$row.SourceExists) -eq 'False'
    $prov = "$($row.SnapshotSku), $(Fmt (N $row.SizeGB) 0) GB, $(if (B $row.Incremental) { 'incremental' } else { 'full' })"
    $status = if ($orph) { 'Orphaned' } elseif ($age -gt $SnapshotMaxAge) { 'Aged' } else { 'In use' }
    Set-Analysis $row $status $prov "Age $(Fmt $age 0) days; source disk $(if ($orph) { 'deleted' } else { 'exists' })" $null
    if ($orph -or $age -gt $SnapshotMaxAge) {
        $sav = if ($row.MonthlyCost) { $row.MonthlyCost } elseif (-not (B $row.Incremental)) { [Math]::Round((N $row.SizeGB) * 0.05, 2) } else { $null }
        Add-Rec $row 'Snapshots' 'Delete' "$(if ($orph) { 'Snapshot of a deleted disk' } else { "Snapshot older than $(Fmt $SnapshotMaxAge 0) days ($(Fmt $age 0) days)" }) - validate retention requirement and delete" 'Delete' $sav 'Medium' $prov -Rule 'SNAPSHOT'
    }
}

$Analyzers.StorageAccounts = { param($row)
    $used = N $row.UsedCapacityGB_Max; $tx = N $row.Transactions_Total; $red = [string]$row.Redundancy; $hot = N $row.BlobHotGB
    $prov = "$($row.StorageKind) $($row.StorageSku)$(if ($row.AccessTier) { ", default tier $($row.AccessTier)" })"
    $util = "Used $(Fmt $used) GB; transactions $(Fmt $tx 0) ($(Fmt ($(if ($null -ne $tx) { $tx / $LookbackDays } else { $null })) 0)/day); egress $(Fmt (N $row.EgressGB_Total)) GB$(if ($row.BlobTierBreakdown) { "; blob tiers: $($row.BlobTierBreakdown)" })"
    if ($row.MetricsStatus -notin @('OK', 'Partial', 'NoData')) { Set-Analysis $row 'No data' $prov $util $null; return }
    if (($tx ?? 0) -eq 0 -and $row.ManagedByService -notin @('Databricks', 'AKS')) {
        Set-Analysis $row 'Unused' $prov $util $null
        Add-Rec $row 'StorageAccounts' 'Delete' "No transactions in $LookbackDays days ($(Fmt $used) GB stored) - validate, archive if required and delete" 'Delete / archive' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    Set-Analysis $row 'In use' $prov $util (N $row.FileQuotaUsedPct)
    if ($red -match '^(GRS|RAGRS|GZRS|RAGZRS)$' -and $row.Environment -eq 'Non-Production') {
        Add-Rec $row 'StorageAccounts' 'Configure' "Geo-redundant ($red) storage in non-production - change to LRS/ZRS" "$(if ($red -match 'Z') { 'ZRS' } else { 'LRS' })" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.55) 'Medium' $util -Rule 'TIER-NONPROD'
    }
    if ($hot -gt 1000 -and $null -ne $tx -and ($tx / [Math]::Max(1, $hot)) -lt 50) {
        Add-Rec $row 'StorageAccounts' 'Configure' "$(Fmt $hot 0) GB in Hot tier with low access - add lifecycle management to move cold data to Cool/Cold/Archive" 'Lifecycle policy' ([Math]::Round($hot * 0.4 * 0.008, 2)) 'Low' $util 'Low' 'Lifecycle management: https://learn.microsoft.com/azure/storage/blobs/lifecycle-management-overview' -Rule 'LIFECYCLE'
    }
    if ($row.Performance -eq 'Premium' -and $row.StorageKind -eq 'BlockBlobStorage' -and ($tx ?? 0) / $LookbackDays -lt 10000) {
        Add-Rec $row 'StorageAccounts' 'Configure' 'Premium block blob account with low transaction volume - Standard (Hot) is cheaper per GB' 'Standard StorageV2' $null 'Low' $util -Rule 'RS-SKU'
    }
}

$Analyzers.PremiumFileShares = { param($row)
    $q = N $row.ProvisionedGiB; $u = N $row.UsedGiB; $pct = N $row.UsedPct
    $prov = "$(Fmt $q 0) GiB provisioned ($($row.BillingModel), $($row.StorageSku))"
    $utilTxt = if ($null -ne $u) { "$(Fmt $u) GiB used ($(Fmt $pct)%)" } else { 'Usage not available' }
    if ($null -eq $u) { Set-Analysis $row 'No data' $prov $utilTxt $null; return }
    $status = if ($u -lt 1) { 'Unused' } elseif ($pct -lt 50) { 'Over-provisioned' } elseif ($pct -gt 90) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $utilTxt $pct
    if ($status -in @('Unused', 'Over-provisioned') -and $q -gt 100) {
        $t = [Math]::Max(100, [Math]::Ceiling($u * 1.3 / 10) * 10)
        if ($t -lt $q) {
            $red = (([string]$row.StorageSku) -split '_')[-1]
            $price = Get-FilesPremiumGiBMonth $row.Location $red
            $sav = if ($row.BillingModel -like 'Provisioned v1*') { [Math]::Round(($q - $t) * $price, 2) } else { Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $q) -Share 0.8 }
            $qTxt = if ($q -ge 1024) { "$(Fmt ($q / 1024) 1) TiB" } else { "$(Fmt $q 0) GiB" }
            $tTxt = if ($t -ge 1024) { "$(Fmt ($t / 1024) 1) TiB" } else { "$t GiB" }
            Add-Rec $row 'PremiumFileShares' 'Rightsize' "$qTxt provisioned on share '$($row.ShareName)', $(Fmt $u) GiB used - resize quota to $tTxt (validate IOPS/throughput needs)" "Quota $tTxt" $sav 'High' $utilTxt -Rule 'RS-STORAGE'
        }
    }
}

$Analyzers.NetAppCapacityPools = { param($row)
    $size = N $row.PoolSizeTiB; $used = N $row.LogicalSizeTiB_Max; $pct = N $row.PoolUsedPct
    $prov = "$($row.ServiceLevel) pool, $(Fmt $size 1) TiB"; $util = "Logical used $(Fmt $used 2) TiB ($(Fmt $pct)%)"
    if ($null -eq $pct) { Set-Analysis $row 'No data' $prov $util $null; return }
    $status = if ($pct -lt 50) { 'Over-provisioned' } elseif ($pct -gt 90) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $pct
    if ($status -eq 'Over-provisioned' -and $size -gt 1) {
        $t = [Math]::Max(1, [Math]::Ceiling($used * 1.25))
        if ($t -lt $size) { Add-Rec $row 'NetAppCapacityPools' 'Rightsize' "Shrink capacity pool from $(Fmt $size 1) TiB to $t TiB" "$t TiB" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $size)) 'Medium' $util -Rule 'RS-STORAGE' }
    }
}

$Analyzers.RecoveryServicesVaults = { param($row)
    $items = [int]((N $row.ProtectedItems) ?? 0); $orph = [int]((N $row.ItemsOrphanedSource) ?? 0)
    $prov = "$($row.StorageRedundancy) backup storage; $items protected item(s) ($($row.ProtectedByWorkload))"
    $status = if ($items -eq 0) { 'Unused' } elseif ($orph -gt 0) { 'Orphaned' } else { 'In use' }
    Set-Analysis $row $status $prov "$items item(s); $orph item(s) protecting deleted resources; $($row.ItemsProtectionStopped) with protection stopped" $null
    if ($orph -gt 0) { Add-Rec $row 'RecoveryServicesVaults' 'Delete' "$orph backup item(s) still retained for source resources that no longer exist - stop protection and delete backup data if retention is not required" 'Delete backup data' $null 'Medium' "See Orphaned Resources sheet (Backup item for deleted source resource)" -Rule 'BACKUP' }
    if ($row.StorageRedundancy -eq 'GeoRedundant' -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'RecoveryServicesVaults' 'Configure' 'Non-production vault using geo-redundant backup storage - use LRS (only possible before first backup / new vault)' 'LRS' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.6) 'Low' 'GeoRedundant backup storage' -Rule 'TIER-NONPROD' }
}

#endregion

#region ============================== ANALYZERS: MONITORING & SECURITY ==============================

$script:LaAnalysis = [System.Collections.Generic.List[object]]::new()
$LaSeriesById = @{}
foreach ($g in ($LaDaily | Group-Object { ([string]$_.ResourceId).ToLower() })) {
    $LaSeriesById[$g.Name] = @($g.Group | Sort-Object Date | ForEach-Object { [double](N $_.BillableGB) })
}

$Analyzers.LogAnalyticsWorkspaces = { param($row)
    $id = ([string]$row.ResourceId).ToLower()
    $daily = [double[]]@($LaSeriesById[$id])
    $avg = N $row.AvgDailyBillableGB; $mx = N $row.MaxDailyBillableGB; $p95 = N $row.P95DailyBillableGB; $mn = N $row.MinDailyBillableGB
    $isCommit = ([string]$row.PricingTier) -match '(?i)capacityreservation'
    $laCommit = if ($isCommit) { N $row.CommitmentGBPerDay } else { $null }
    $sent = B $row.SentinelEnabled; $simplified = ([string]$row.SentinelPricing) -like 'Simplified*'
    $sentCommit = N $row.SentinelCommitmentGBPerDay
    $prov = "LA: $(if ($laCommit) { "$laCommit GB/day commitment" } else { $row.PricingTier }); retention $($row.RetentionDays) days$(if ($sent) { "; Sentinel: $($row.SentinelPricing)$(if ($sentCommit) { " $sentCommit GB/day" })" })"
    $util = if ($null -ne $avg) { "Billable ingestion avg $(Fmt $avg 1) GB/day | P95 $(Fmt $p95 1) | max $(Fmt $mx 1) | min $(Fmt $mn 1); total $(Fmt (N $row.TotalBillableGB) 0) GB in $($row.ActiveDays) day(s)" } else { "Ingestion query: $($row.IngestionQueryStatus)" }
    if ($null -eq $avg -and $row.IngestionQueryStatus -notlike 'OK*') { Set-Analysis $row 'No data' $prov $util $null; Add-LaAnalysisRow $row 'No data' @{}; return }
    if (($avg ?? 0) -eq 0) {
        Set-Analysis $row 'Unused' $prov 'No billable ingestion in the window' 0
        Add-LaAnalysisRow $row 'Unused' @{ AvgDailyGB = 0 }
        if ($laCommit -or $sentCommit) { Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' 'Commitment tier with zero ingestion - switch to pay-as-you-go immediately' 'Pay-as-you-go' $row.MonthlyCost 'High' $prov -Rule 'COMMIT-TIER' }
        else { Add-Rec $row 'LogAnalyticsWorkspaces' 'Delete' "No billable data ingested in $LookbackDays days - validate and delete the workspace" 'Delete' $row.MonthlyCost 'Medium' $prov -Rule 'NOACT' }
        return
    }
    $book = Get-LaPriceBook $row.Location
    if ($daily.Count -eq 0) { $daily = [double[]]@($avg) }
    $commitForUtil = if ($laCommit) { $laCommit } elseif ($sentCommit) { $sentCommit } else { $null }
    $utilPct = if ($commitForUtil) { Get-Pct $avg $commitForUtil } else { $null }
    $status = if ($commitForUtil) { if ($utilPct -lt 85) { 'Over-committed' } elseif ($utilPct -gt 125) { 'Under-committed' } else { 'Right-sized' } } elseif ($avg -ge 100) { 'Commitment opportunity' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $utilPct
    $evidence = "Max daily observed $(Fmt $mx 0) GB, P95 $(Fmt $p95 0) GB, average $(Fmt $avg 0) GB$(if ($commitForUtil) { " vs $commitForUtil GB/day commitment ($(Fmt $utilPct 0)% utilized; $($row.DaysBelowCommitment) of $($row.ActiveDays) days below commitment)" })"
    $curLa = $null; $bestLa = $null; $curSent = $null; $bestSent = $null; $laSave = 0; $sentSave = 0
    if ($sent -and $simplified) {
        # Simplified pricing: one unified meter (Log Analytics + Sentinel) billed on the workspace tier
        $uniPlans = Get-TierPlan -Daily $daily -Payg $book.UniPayg -Tiers $book.UniTiers
        $curLa = $uniPlans | Where-Object Tier -eq ([int]($laCommit ?? 0)) | Select-Object -First 1
        $bestLa = $uniPlans | Sort-Object Monthly | Select-Object -First 1
        if ($curLa -and $bestLa -and $bestLa.Tier -ne $curLa.Tier) { $laSave = [Math]::Round($curLa.Monthly - $bestLa.Monthly, 2) }
        if ($laSave -gt [Math]::Max(50.0, 0.05 * [double]$curLa.Monthly)) {
            Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "Change the Sentinel simplified pricing tier (Log Analytics + Sentinel) from $($curLa.Name) to $($bestLa.Name)" $bestLa.Name $laSave 'High' $evidence 'Low' 'Simplified pricing tiers: https://learn.microsoft.com/azure/sentinel/billing' -Rule 'COMMIT-TIER'
        }
        else { $laSave = 0 }
    }
    else {
        $laPlans = Get-TierPlan -Daily $daily -Payg $book.LaPayg -Tiers $book.LaTiers
        $curLa = if ($laCommit) { $laPlans | Where-Object Tier -eq ([int]$laCommit) | Select-Object -First 1 } else { $laPlans | Where-Object Tier -eq 0 | Select-Object -First 1 }
        $bestLa = $laPlans | Sort-Object Monthly | Select-Object -First 1
        if ($curLa -and $bestLa -and $bestLa.Tier -ne $curLa.Tier) {
            $laSave = [Math]::Round(($curLa.Monthly - $bestLa.Monthly) * $(if ($row.MonthlyCost -and $curLa.Monthly -gt 0 -and -not $sent) { [Math]::Min(1.5, $row.MonthlyCost / $curLa.Monthly) } else { 1 }), 2)
        }
        if ($laSave -gt [Math]::Max(50.0, 0.05 * [double]$curLa.Monthly)) {
            Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "Change Log Analytics pricing tier from $($curLa.Name) to $($bestLa.Name)" $bestLa.Name $laSave 'High' $evidence -Rule 'COMMIT-TIER'
        }
        else { $laSave = 0 }
        if ($sent) {
            # Classic pricing: Sentinel analysis billed separately (Sentinel-only price = unified - Log Analytics)
            $sentPlans = Get-TierPlan -Daily $daily -Payg $book.SentClassicPayg -Tiers $book.SentClassicTiers
            $curSent = $sentPlans | Where-Object Tier -eq ([int]($sentCommit ?? 0)) | Select-Object -First 1
            $bestSent = $sentPlans | Sort-Object Monthly | Select-Object -First 1
            if ($curSent -and $bestSent -and $bestSent.Tier -ne $curSent.Tier) { $sentSave = [Math]::Round($curSent.Monthly - $bestSent.Monthly, 2) }
            if ($sentSave -gt [Math]::Max(50.0, 0.05 * [double]$curSent.Monthly)) {
                Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "Change Microsoft Sentinel (classic) pricing tier from $($curSent.Name) to $($bestSent.Name) - or move to the simplified pricing model" $bestSent.Name $sentSave 'High' $evidence 'Low' 'Sentinel pricing: https://learn.microsoft.com/azure/sentinel/billing' -Rule 'COMMIT-TIER'
            }
            else { $sentSave = 0 }
        }
    }
    $freeRet = if ($sent) { 90 } else { 31 }
    $ret = [int]((N $row.RetentionDays) ?? 30)
    if ($ret -gt $freeRet) {
        $retGb = $avg * ($ret - $freeRet)
        $sav = if ($book.Retention) { [Math]::Round($retGb * ($book.Retention - 0.025), 2) } else { $null }
        Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "Interactive retention of $ret days (free: $freeRet) - keep $freeRet days interactive and move older data to long-term (archive) retention per table, subject to compliance" "$freeRet days interactive + long-term retention" $sav 'Low' "~$(Fmt $retGb 0) GB retained beyond the free period" 'Low' 'Data retention: https://learn.microsoft.com/azure/azure-monitor/logs/data-retention-configure' -Rule 'RETENTION'
    }
    $tops = @($LaTopTables | Where-Object { ([string]$_.ResourceId).ToLower() -eq $id } | Sort-Object { [double](N $_.AvgDailyGB) } -Descending | Select-Object -First 1)
    if ($tops.Count -and (N $tops[0].AvgDailyGB) -ge 10 -and (N $tops[0].AvgDailyGB) -ge 0.3 * $avg) {
        $t = $tops[0]
        Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "Table $($t.DataType) ingests $(Fmt (N $t.AvgDailyGB) 1) GB/day ($(Fmt ((N $t.AvgDailyGB) / $avg * 100) 0)% of workspace) - filter with DCR transformations or move to Basic/Auxiliary logs if not used for analytics" "Reduce $($t.DataType) volume" $null 'Low' "Top table by billable volume" 'Medium' 'Optimize log ingestion: https://learn.microsoft.com/azure/azure-monitor/logs/cost-logs' -Rule 'INGESTION'
    }
    if ($null -ne (N $row.QueryCount_Total) -and (N $row.QueryCount_Total) -eq 0 -and $avg -ge 1 -and -not $sent) {
        $sav = if ($book.Basic -and $book.LaPayg) { [Math]::Round($avg * 30.4375 * [Math]::Max(0.0, $book.LaPayg - $book.Basic), 2) } else { $null }
        Add-Rec $row 'LogAnalyticsWorkspaces' 'Configure' "$(Fmt $avg 1) GB/day ingested but the workspace was never queried in $LookbackDays days - stop unused collection, reduce retention or move tables to Basic/Auxiliary logs" 'Reduce / re-tier ingestion' $sav 'Medium' "Query Count = 0 over $LookbackDays days; average ingestion $(Fmt $avg 1) GB/day" 'Medium' 'Table plans: https://learn.microsoft.com/azure/azure-monitor/logs/logs-table-plans' -Rule 'INGESTION'
    }
    Add-LaAnalysisRow $row $status @{
        AvgDailyGB = $avg; P95DailyGB = $p95; MaxDailyGB = $mx; MinDailyGB = $mn; CommitmentUtilizationPct = $utilPct
        CurrentLAMonthlyEst = $curLa.Monthly; OptimalLATier = $bestLa.Name; OptimalLAMonthlyEst = $bestLa.Monthly
        CurrentSentinelMonthlyEst = $curSent.Monthly; OptimalSentinelTier = $bestSent.Name; OptimalSentinelMonthlyEst = $bestSent.Monthly
        EstMonthlySavings = [Math]::Round([double]$laSave + [double]$sentSave, 0); PriceSource = $book.Source
    }
}

function Add-LaAnalysisRow($row, [string]$Status, [hashtable]$V) {
    $laCommit = if (([string]$row.PricingTier) -match '(?i)capacityreservation') { N $row.CommitmentGBPerDay } else { $null }
    $script:LaAnalysis.Add([pscustomobject][ordered]@{
            SubscriptionName = $row.SubscriptionName; ResourceGroup = $row.ResourceGroup; Workspace = $row.ResourceName; Location = $row.Location
            LAPricingTier = $(if ($laCommit) { "$laCommit GB/day commitment" } else { 'Pay-as-you-go' }); SentinelEnabled = (B $row.SentinelEnabled); SentinelPricing = $row.SentinelPricing
            SentinelCommitmentGBPerDay = N $row.SentinelCommitmentGBPerDay; RetentionDays = $row.RetentionDays; ActiveDays = $row.ActiveDays
            AvgDailyGB = $V.AvgDailyGB; P95DailyGB = $V.P95DailyGB; MaxDailyGB = $V.MaxDailyGB; MinDailyGB = $V.MinDailyGB
            WeekdayAvgGB = N $row.WeekdayAvgGB; WeekendAvgGB = N $row.WeekendAvgGB
            CommitmentUtilizationPct = $V.CommitmentUtilizationPct; DaysBelowCommitment = $row.DaysBelowCommitment; DaysAboveCommitment = $row.DaysAboveCommitment
            CurrentLAMonthlyEst = $V.CurrentLAMonthlyEst; OptimalLATier = $V.OptimalLATier; OptimalLAMonthlyEst = $V.OptimalLAMonthlyEst
            CurrentSentinelMonthlyEst = $V.CurrentSentinelMonthlyEst; OptimalSentinelTier = $V.OptimalSentinelTier; OptimalSentinelMonthlyEst = $V.OptimalSentinelMonthlyEst
            EstMonthlySavings = $V.EstMonthlySavings; PriceSource = $V.PriceSource; TopTables = $row.TopTables; Status = $Status
        })
}

$Analyzers.LogAnalyticsClusters = { param($row)
    $commit = N $row.CommitmentGBPerDay
    $linked = @($TypeData['LogAnalyticsWorkspaces'] | Where-Object { $_.DedicatedCluster -eq $row.ResourceName })
    $sum = [double](($linked | ForEach-Object { [double]((N $_.AvgDailyBillableGB) ?? 0) } | Measure-Object -Sum).Sum)
    $pct = Get-Pct $sum $commit
    $status = if ($null -eq $pct) { 'No data' } elseif ($pct -lt 85) { 'Over-committed' } elseif ($pct -gt 125) { 'Under-committed' } else { 'Right-sized' }
    Set-Analysis $row $status "$commit GB/day cluster commitment, $($linked.Count) linked workspace(s)" "Linked workspaces ingest $(Fmt $sum 1) GB/day ($(Fmt $pct 0)% of commitment)" $pct
    if ($status -eq 'Over-committed' -and $commit -gt 100) { Add-Rec $row 'LogAnalyticsClusters' 'Configure' "Dedicated cluster commitment $commit GB/day vs $(Fmt $sum 0) GB/day ingested - lower the commitment tier (min 100 GB/day) or unlink and use workspace tiers" 'Lower cluster commitment' $null 'Medium' "Cluster utilization $(Fmt $pct 0)%" -Rule 'COMMIT-TIER' }
}

$Analyzers.ApplicationInsights = { param($row)
    $req = N $row.Requests_Total
    $prov = "$(if (B $row.IsClassic) { 'Classic' } else { "Workspace-based ($($row.WorkspaceName))" }), sampling $(if ($row.SamplingPct) { "$($row.SamplingPct)%" } else { 'not set' })"
    $util = "Requests $(Fmt $req 0)"
    if (B $row.IsClassic) { Add-Rec $row 'ApplicationInsights' 'Configure' 'Classic Application Insights (retired) - migrate to workspace-based' 'Workspace-based' $null 'High' 'Classic resource' 'Low' 'Migrate: https://learn.microsoft.com/azure/azure-monitor/app/convert-classic-resource' -Rule 'MODERNIZE' }
    if (($req ?? -1) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial', 'NoData')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'ApplicationInsights' 'Delete' "No requests tracked in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Low' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
}

$Analyzers.ManagedHSMs = { param($row)
    $hits = N $row.ApiHits_Total
    if (($hits ?? -1) -eq 0) { Set-Analysis $row 'Unused' "$($row.HsmSku)" "No API calls in $LookbackDays days" 0; Add-Rec $row 'ManagedHSMs' 'Delete' 'Managed HSM with no API calls - validate key dependencies and delete' 'Delete' ($row.MonthlyCost ?? 2336) 'Medium' "0 API hits in $LookbackDays days" -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' "$($row.HsmSku)" "API hits $(Fmt $hits 0)" $null
}

#endregion

#region ============================== ANALYZERS: MESSAGING & NETWORKING ==============================

function Test-NoActivity($row, [string[]]$Cols) {
    if ($row.MetricsStatus -notin @('OK', 'Partial', 'NoData')) { return $false }
    foreach ($c in $Cols) { if (([double]((N $row.$c) ?? 0)) -gt 0) { return $false } }
    return $true
}

$Analyzers.EventHubNamespaces = { param($row)
    $cap = N $row.Capacity; $tu = N $row.TUUtilPct_PeakHour; $cpu = N $row.CPU_P95
    $prov = "$($row.Tier), $(Fmt $cap 0) $(if ($row.Tier -eq 'Premium') { 'PU' } else { 'TU' })$(if (B $row.AutoInflate) { " (auto-inflate to $($row.MaxThroughputUnits))" })"
    $util = "Messages in $(Fmt (N $row.InMsgs_Total) 0) / out $(Fmt (N $row.OutMsgs_Total) 0); peak-hour ingress $(Fmt (N $row.PeakHourIngressMBps) 3) MB/s$(if ($null -ne $tu) { " ($(Fmt $tu)% of TU capacity)" }); throttled $(Fmt (N $row.Throttled_Total) 0)"
    if (Test-NoActivity $row @('InMsgs_Total', 'OutMsgs_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'EventHubNamespaces' 'Delete' "No messages in $LookbackDays days - delete the namespace" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    $pct = if ($row.Tier -eq 'Premium') { $cpu } else { $tu }
    $status = if ($null -eq $pct) { 'In use' } elseif ($pct -lt 25) { 'Over-provisioned' } elseif ((N $row.Throttled_Total) -gt 0 -or $pct -gt 85) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $pct
    if ($status -eq 'Over-provisioned' -and $cap -gt 1) {
        $t = [int][Math]::Max(1, [Math]::Ceiling($cap * $pct / $TargetUtilization))
        if ($t -lt $cap) { Add-Rec $row 'EventHubNamespaces' 'Rightsize' "Reduce $(if ($row.Tier -eq 'Premium') { 'processing units' } else { 'throughput units' }) from $(Fmt $cap 0) to $t$(if ($row.Tier -ne 'Premium' -and -not (B $row.AutoInflate)) { ' and enable auto-inflate' })" "$t" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $cap)) 'Medium' $util -Rule 'RS-SCALE' }
    }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'EventHubNamespaces' 'Performance Risk' 'Throttling / high utilization observed - increase capacity or enable auto-inflate' 'Scale up' $null 'Medium' $util -Rule 'PERF-RISK' }
}

$Analyzers.ServiceBusNamespaces = { param($row)
    $cap = N $row.Capacity; $cpu = N $row.CPU_P95
    $prov = "$($row.Tier)$(if ($row.Tier -eq 'Premium') { ", $(Fmt $cap 0) messaging unit(s)" })"
    $util = "Messages in $(Fmt (N $row.InMsgs_Total) 0) / out $(Fmt (N $row.OutMsgs_Total) 0); max connections $(Fmt (N $row.Connections_Max) 0)$(if ($null -ne $cpu) { "; CPU P95 $(Fmt $cpu)%" })"
    if (Test-NoActivity $row @('InMsgs_Total', 'OutMsgs_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'ServiceBusNamespaces' 'Delete' "No messages in $LookbackDays days - delete the namespace" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    $status = if ($row.Tier -ne 'Premium' -or $null -eq $cpu) { 'In use' } elseif ($cpu -lt 20) { 'Over-provisioned' } elseif ($cpu -gt 80) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $cpu
    if ($status -eq 'Over-provisioned' -and $cap -gt 1) { $t = [Math]::Max(1, [Math]::Floor($cap / 2)); Add-Rec $row 'ServiceBusNamespaces' 'Rightsize' "Reduce messaging units from $(Fmt $cap 0) to $t" "$t MU" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $cap)) 'Medium' $util -Rule 'RS-SCALE' }
    if ($row.Tier -eq 'Premium' -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'ServiceBusNamespaces' 'Configure' 'Premium namespace in non-production - use Standard tier' 'Standard' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.1) 'Medium' $util -Rule 'TIER-NONPROD' }
}

$Analyzers.LogicApps = { param($row)
    $runs = N $row.Runs_Total
    $util = "Runs $(Fmt $runs 0) (failed $(Fmt (N $row.RunsFailed_Total) 0)); billable executions $(Fmt (N $row.BillableExecutions_Total) 0)"
    if ($row.State -eq 'Disabled') { Set-Analysis $row 'Stopped' "Consumption workflow ($($row.State))" 'Disabled' $null; return }
    if (Test-NoActivity $row @('Runs_Total')) { Set-Analysis $row 'Unused' "Consumption workflow" $util 0; Add-Rec $row 'LogicApps' 'Delete' "Enabled workflow with no runs in $LookbackDays days - disable or delete" 'Delete' $row.MonthlyCost 'Low' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' 'Consumption workflow' $util $null
}

$Analyzers.APIManagement = { param($row)
    $cap = N $row.Capacity_P95; $units = N $row.Units; $req = N $row.Requests_Total
    $prov = "$($row.ApimSku) x $(Fmt $units 0) unit(s)$(if ($row.PlatformVersion) { ", $($row.PlatformVersion)" })"
    $util = "Capacity P95 $(Fmt $cap)%; requests $(Fmt $req 0)"
    if ($row.PlatformVersion -eq 'stv1') { Add-Rec $row 'APIManagement' 'Configure' 'stv1 compute platform is retired - migrate to stv2' 'stv2' $null 'High' 'Platform stv1' 'Medium' -Rule 'MODERNIZE' }
    if (Test-NoActivity $row @('Requests_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'APIManagement' 'Delete' "No API requests in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    $status = if ($null -eq $cap) { 'In use' } elseif ($cap -lt 20) { 'Over-provisioned' } elseif ($cap -gt 80) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $cap
    if ($status -eq 'Over-provisioned' -and $units -gt 1) { $t = [int][Math]::Max(1, [Math]::Ceiling($units * $cap / $TargetUtilization)); if ($t -lt $units) { Add-Rec $row 'APIManagement' 'Rightsize' "Reduce units from $(Fmt $units 0) to $t" "$t unit(s)" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $units)) 'Medium' $util -Rule 'RS-SCALE' } }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'APIManagement' 'Performance Risk' 'Gateway capacity above 80% - add units / autoscale' 'Scale out' $null 'Medium' $util -Rule 'PERF-RISK' }
}

$Analyzers.IoTHubs = { param($row)
    $q = N $row.QuotaUtilPct_Max; $units = N $row.Units
    $prov = "$($row.IotSku) x $(Fmt $units 0) unit(s) ($(Fmt (N $row.DailyMessageQuota) 0) msgs/day)"
    $util = "Peak daily quota used $(Fmt $q)%; telemetry messages $(Fmt (N $row.TelemetryMsgs_Total) 0)"
    if (Test-NoActivity $row @('TelemetryMsgs_Total', 'DailyMsgsUsed_Max')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'IoTHubs' 'Delete' "No device messages in $LookbackDays days" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    $status = if ($null -eq $q) { 'In use' } elseif ($q -lt 30) { 'Over-provisioned' } elseif ($q -gt 85) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $q
    if ($status -eq 'Over-provisioned' -and $units -gt 1) { $t = [int][Math]::Max(1, [Math]::Ceiling($units * $q / $TargetUtilization)); if ($t -lt $units) { Add-Rec $row 'IoTHubs' 'Rightsize' "Reduce IoT Hub units from $(Fmt $units 0) to $t" "$t unit(s)" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $units)) 'Medium' $util -Rule 'RS-SCALE' } }
}

$Analyzers.StreamAnalyticsJobs = { param($row)
    $su = N $row.SUUtilPct_P95; $n = N $row.StreamingUnits
    $prov = "$(Fmt $n 0) streaming unit(s), state $($row.JobState)"; $util = "SU utilization P95 $(Fmt $su)%; input events $(Fmt (N $row.InputEvents_Total) 0)"
    if ($row.JobState -ne 'Running') { Set-Analysis $row 'Stopped' $prov 'Job not running (not billed)' $null; return }
    $status = if ($null -eq $su) { 'No data' } elseif ($su -lt 30) { 'Over-provisioned' } elseif ($su -gt 80) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $su
    if ($status -eq 'Over-provisioned' -and $n -gt 1) { $t = [Math]::Max(1, [Math]::Ceiling($n * $su / $TargetUtilization)); if ($t -lt $n) { Add-Rec $row 'StreamAnalyticsJobs' 'Rightsize' "Reduce streaming units from $(Fmt $n 0) to $t" "$t SU" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $n)) 'Medium' $util -Rule 'RS-SCALE' } }
}

$Analyzers.ApplicationGateways = { param($row)
    $cu = N $row.CapacityUnits_P95; $req = N $row.Requests_Total; $targets = [int]((N $row.BackendTargets) ?? 0)
    $prov = "$($row.GatewaySku)$(if ($row.FixedCapacity) { " x $($row.FixedCapacity) (fixed)" } else { " autoscale $($row.AutoscaleMin)-$($row.AutoscaleMax)" }), $($row.BackendPools) pool(s), $targets backend target(s)"
    $util = "Capacity units P95 $(Fmt $cu); throughput P95 $(Fmt (N $row.ThroughputMbps_P95) 2) Mbps; requests $(Fmt $req 0)"
    if ($row.OperationalState -eq 'Stopped') { Set-Analysis $row 'Stopped' $prov 'Gateway stopped' $null; return }
    if ($targets -eq 0) { Set-Analysis $row 'Orphaned' $prov 'No backend targets configured' 0; Add-Rec $row 'ApplicationGateways' 'Delete' 'Application Gateway with no backend targets - delete' 'Delete' $row.MonthlyCost 'High' $prov -Rule 'ORPHAN'; return }
    if (Test-NoActivity $row @('Requests_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'ApplicationGateways' 'Delete' "No requests in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row $(if ($row.GatewayTier -match 'v2' -and $null -ne $cu -and $cu -lt 2) { 'Over-provisioned' } else { 'In use' }) $prov $util $null
    if ($row.GatewayTier -match 'v2' -and $row.FixedCapacity -and -not $row.AutoscaleMax) { Add-Rec $row 'ApplicationGateways' 'Configure' "Fixed capacity ($($row.FixedCapacity) instances) - enable autoscaling with a low minimum (0-2)" 'Autoscale min 0-2' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.7) 'Medium' $util -Rule 'AUTOSCALE' }
    if ($row.GatewayTier -notmatch 'v2') { Add-Rec $row 'ApplicationGateways' 'Configure' 'Application Gateway v1 is retiring - migrate to v2' 'Standard_v2 / WAF_v2' $null 'High' "SKU $($row.GatewaySku)" 'Medium' -Rule 'MODERNIZE' }
}

$Analyzers.LoadBalancers = { param($row)
    $targets = [int]((N $row.BackendTargets) ?? 0); $data = N $row.DataGB_Total
    $prov = "$($row.LbSku) ($($row.Exposure)), $($row.BackendPools) pool(s), $targets target(s), $($row.LbRules) rule(s)"
    $util = "Data processed $(Fmt $data 2) GB"
    if ($row.LbSku -eq 'Basic') { Set-Analysis $row 'Not assessed' $prov 'Basic SKU (no metrics)' $null; Add-Rec $row 'LoadBalancers' 'Configure' 'Basic Load Balancer is retired - upgrade to Standard' 'Standard' $null 'High' 'Basic SKU' 'Medium' -Rule 'MODERNIZE'; return }
    if ($row.ManagedByService -eq 'AKS') { Set-Analysis $row 'Managed by service' $prov $util $null; return }
    if ($targets -eq 0) { Set-Analysis $row 'Orphaned' $prov 'No backend targets' 0; Add-Rec $row 'LoadBalancers' 'Delete' 'Standard Load Balancer without backend targets - delete' 'Delete' $row.MonthlyCost 'High' $prov -Rule 'ORPHAN'; return }
    if (Test-NoActivity $row @('DataGB_Total', 'Packets_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'LoadBalancers' 'Delete' "No traffic in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
}

$Analyzers.PublicIPAddresses = { param($row)
    $prov = "$($row.PipSku) / $($row.AllocationMethod) $($row.IpAddress)"
    if (-not (B $row.IsAttached)) {
        $billed = $row.PipSku -eq 'Standard' -or $row.AllocationMethod -eq 'Static'
        Set-Analysis $row 'Orphaned' $prov 'Not associated with any resource' 0
        if ($billed) { Add-Rec $row 'PublicIPAddresses' 'Delete' 'Unassociated public IP address - delete (still billed)' 'Delete' ($row.MonthlyCost ?? 3.65) 'High' $prov -Rule 'ORPHAN' }
        return
    }
    Set-Analysis $row 'In use' $prov "Attached to $($row.AttachedToType) '$($row.AttachedTo)'" $null
    if ($row.PipSku -eq 'Basic') { Add-Rec $row 'PublicIPAddresses' 'Configure' 'Basic SKU public IP is retired - upgrade to Standard' 'Standard' $null 'High' 'Basic SKU' 'Medium' -Rule 'MODERNIZE' }
}

$Analyzers.NatGateways = { param($row)
    $prov = "$($row.PublicIPs) public IP(s), $($row.Subnets) subnet(s)"; $util = "Data $(Fmt (N $row.DataGB_Total) 2) GB; connections $(Fmt (N $row.Connections_Total) 0)"
    if ([int]((N $row.Subnets) ?? 0) -eq 0) { Set-Analysis $row 'Orphaned' $prov 'Not attached to a subnet' 0; Add-Rec $row 'NatGateways' 'Delete' 'NAT gateway not attached to any subnet - delete' 'Delete' ($row.MonthlyCost ?? 32.85) 'High' $prov -Rule 'ORPHAN'; return }
    if (Test-NoActivity $row @('DataGB_Total', 'Connections_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'NatGateways' 'Delete' "No outbound connections or traffic in $LookbackDays days" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
}

$Analyzers.AzureFirewalls = { param($row)
    $tp = N $row.ThroughputMbps_P95; $data = N $row.DataProcessedGB_Total
    $prov = "$($row.FirewallTier) ($($row.FirewallSkuName))"; $util = "Throughput P95 $(Fmt $tp 2) Mbps | max $(Fmt (N $row.ThroughputMbps_Max) 2); data processed $(Fmt $data 1) GB"
    if ([int]((N $row.IpConfigurations) ?? 0) -eq 0 -and -not $row.VirtualHub) { Set-Analysis $row 'Deallocated' $prov 'Firewall deallocated (no IP configuration)' $null; return }
    if (Test-NoActivity $row @('DataProcessedGB_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'AzureFirewalls' 'Delete' "No traffic processed in $LookbackDays days - deallocate or delete" 'Deallocate / delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
    if ($row.FirewallTier -eq 'Premium' -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'AzureFirewalls' 'Configure' 'Premium firewall in non-production - use Standard (or Basic for low throughput)' 'Standard / Basic' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.7) 'Medium' $util -Rule 'TIER-NONPROD' }
    elseif ($row.FirewallTier -eq 'Standard' -and $row.Environment -eq 'Non-Production' -and ($tp ?? 999) -lt 50) { Add-Rec $row 'AzureFirewalls' 'Configure' 'Low-throughput non-production firewall - evaluate Azure Firewall Basic' 'Basic' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.3) 'Low' $util -Rule 'TIER-NONPROD' }
}

$VpnCapacityMbps = @{ VpnGw1 = 650; VpnGw2 = 1000; VpnGw3 = 1250; VpnGw4 = 5000; VpnGw5 = 10000; VpnGw1AZ = 650; VpnGw2AZ = 1000; VpnGw3AZ = 1250; VpnGw4AZ = 5000; VpnGw5AZ = 10000; Basic = 100 }
$VpnDown = @{ VpnGw5 = 'VpnGw4'; VpnGw4 = 'VpnGw3'; VpnGw3 = 'VpnGw2'; VpnGw2 = 'VpnGw1'; VpnGw5AZ = 'VpnGw4AZ'; VpnGw4AZ = 'VpnGw3AZ'; VpnGw3AZ = 'VpnGw2AZ'; VpnGw2AZ = 'VpnGw1AZ' }
$VpnHourly = @{ VpnGw1 = 0.19; VpnGw2 = 0.49; VpnGw3 = 1.25; VpnGw4 = 2.10; VpnGw5 = 3.65; VpnGw1AZ = 0.361; VpnGw2AZ = 0.564; VpnGw3AZ = 1.438; VpnGw4AZ = 2.42; VpnGw5AZ = 4.20 }

$Analyzers.VirtualNetworkGateways = { param($row)
    $conns = [int]((N $row.Connections) ?? 0); $p2s = B $row.P2SConfigured; $sku = [string]$row.GatewaySku
    $bw = if ($row.GatewayType -eq 'ExpressRoute') { N $row.BitsMbps_P95 } else { N $row.AvgBandwidthMbps_P95 }
    $prov = "$($row.GatewayType) $sku, $conns connection(s)$(if ($p2s) { ', P2S configured' })"
    $util = "Bandwidth P95 $(Fmt $bw 2) Mbps$(if ($p2s) { "; max P2S connections $(Fmt (N $row.P2SConnections_Max) 0)" })"
    if ($conns -eq 0 -and (-not $p2s -or ((N $row.P2SConnections_Max) ?? 0) -eq 0)) { Set-Analysis $row 'Orphaned' $prov 'No connections' 0; Add-Rec $row 'VirtualNetworkGateways' 'Delete' 'Gateway without site-to-site/ExpressRoute connections or P2S users - delete' 'Delete' $row.MonthlyCost 'High' $prov -Rule 'ORPHAN'; return }
    $cap = $VpnCapacityMbps[$sku]
    $pct = if ($cap -and $null -ne $bw) { Get-Pct $bw $cap } else { $null }
    $status = if ($null -eq $pct) { 'In use' } elseif ($pct -lt 10) { 'Over-provisioned' } elseif ($pct -gt 80) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $pct
    if ($status -eq 'Over-provisioned' -and $VpnDown[$sku]) {
        $t = $VpnDown[$sku]
        if ($bw -lt 0.5 * $VpnCapacityMbps[$t]) { Add-Rec $row 'VirtualNetworkGateways' 'Rightsize' "Resize gateway from $sku to $t (P95 bandwidth $(Fmt $bw 1) Mbps)" $t (Get-Savings -MonthlyCost $row.MonthlyCost -CurrentUnit $VpnHourly[$sku] -TargetUnit $VpnHourly[$t]) 'Medium' $util -Rule 'RS-SKU' }
    }
}

$Analyzers.ExpressRouteCircuits = { param($row)
    $bwc = N $row.BandwidthMbps; $p95 = N $row.P95UtilPct; $peak = N $row.PeakUtilPct
    $prov = "$(Fmt $bwc 0) Mbps $($row.CircuitTier) / $($row.CircuitFamily) via $($row.Provider) ($($row.PeeringLocation))"
    $util = "Utilization P95 $(Fmt $p95)% | peak $(Fmt $peak)% of bandwidth"
    if ($row.ProviderState -eq 'NotProvisioned') { Set-Analysis $row 'Unused' $prov 'Circuit not provisioned by the provider (billing still applies)' 0; Add-Rec $row 'ExpressRouteCircuits' 'Delete' 'Circuit not provisioned by the connectivity provider but billed - complete provisioning or delete' 'Delete' $row.MonthlyCost 'High' $prov -Rule 'ORPHAN'; return }
    $status = if ($null -eq $p95) { 'No data' } elseif ($p95 -lt 10) { 'Over-provisioned' } elseif ($p95 -gt 80) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $p95
    if ($status -eq 'Over-provisioned' -and $bwc -gt 50) { Add-Rec $row 'ExpressRouteCircuits' 'Rightsize' "Bandwidth P95 is $(Fmt $p95)% of $(Fmt $bwc 0) Mbps - evaluate a lower bandwidth (provider dependent) or metered data plan" 'Lower bandwidth' $null 'Low' $util -Rule 'RS-SKU' }
    if ($status -eq 'Under-provisioned') { Add-Rec $row 'ExpressRouteCircuits' 'Performance Risk' 'Circuit utilization above 80% - plan a bandwidth upgrade' 'Higher bandwidth' $null 'Medium' $util -Rule 'PERF-RISK' }
}

$Analyzers.BastionHosts = { param($row)
    $s = N $row.Sessions_Max
    $prov = "$($row.BastionSku), $($row.ScaleUnits) scale unit(s)"; $util = "Max concurrent sessions $(Fmt $s 0); idle $(Fmt (N $row.Pattern_IdleHoursPct) 0)% of hours"
    if ($row.BastionSku -eq 'Developer') { Set-Analysis $row 'Not assessed' $prov 'Developer SKU (free)' $null; return }
    if (Test-NoActivity $row @('Sessions_Max')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'BastionHosts' 'Delete' "No sessions in $LookbackDays days - delete Bastion (or use the Developer SKU / deploy on demand)" 'Delete / Developer SKU' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
    if ([int]((N $row.ScaleUnits) ?? 2) -gt 2 -and ($s ?? 0) -lt 20) { Add-Rec $row 'BastionHosts' 'Rightsize' "Reduce scale units from $($row.ScaleUnits) to 2 (max $(Fmt $s 0) concurrent sessions)" '2 scale units' $null 'Medium' $util -Rule 'RS-SCALE' }
}

$Analyzers.PrivateEndpoints = { param($row)
    $prov = "-> $($row.TargetType) '$($row.TargetResource)' ($($row.GroupIds))"; $util = "Traffic in $(Fmt (N $row.BytesInMB_Total) 1) MB / out $(Fmt (N $row.BytesOutMB_Total) 1) MB; connection $($row.ConnectionState)"
    if ($row.ConnectionState -in @('Disconnected', 'Rejected')) { Set-Analysis $row 'Orphaned' $prov $util 0; Add-Rec $row 'PrivateEndpoints' 'Delete' "Private endpoint connection is $($row.ConnectionState) - delete" 'Delete' ($row.MonthlyCost ?? 7.3) 'High' $util -Rule 'ORPHAN'; return }
    if ($row.ManagedByService) { Set-Analysis $row 'Managed by service' $prov $util $null; return }
    if (Test-NoActivity $row @('BytesInMB_Total', 'BytesOutMB_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'PrivateEndpoints' 'Delete' "No traffic through the private endpoint in $LookbackDays days - validate and delete" 'Delete' ($row.MonthlyCost ?? 7.3) 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
}

$Analyzers.FrontDoorCdn = { param($row)
    $req = N $row.Requests_Total; $prov = "$($row.ProfileSku)"; $util = "Requests $(Fmt $req 0); response data $(Fmt (N $row.ResponseGB_Total) 1) GB"
    if ($row.ProfileSku -eq 'Classic Front Door') { Add-Rec $row 'FrontDoorCdn' 'Configure' 'Front Door (classic) is retiring - migrate to Front Door Standard/Premium' 'AFD Standard/Premium' $null 'High' 'Classic Front Door' 'Medium' -Rule 'MODERNIZE' }
    if (Test-NoActivity $row @('Requests_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'FrontDoorCdn' 'Delete' "No requests in $LookbackDays days - delete the profile" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
    if ($row.ProfileSku -eq 'Premium_AzureFrontDoor' -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'FrontDoorCdn' 'Configure' 'Front Door Premium in non-production - use Standard unless Private Link / managed WAF rules are required' 'Standard_AzureFrontDoor' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.2) 'Low' $util -Rule 'TIER-NONPROD' }
}

$Analyzers.DdosProtectionPlans = { param($row)
    $v = [int]((N $row.ProtectedVNets) ?? 0)
    Set-Analysis $row $(if ($v -eq 0) { 'Orphaned' } else { 'In use' }) 'DDoS Network Protection plan' "$v protected VNet(s): $($row.VNetNames)" $null
    if ($v -eq 0) { Add-Rec $row 'DdosProtectionPlans' 'Delete' 'DDoS protection plan not linked to any VNet - delete (billed monthly)' 'Delete' ($row.MonthlyCost ?? 2944) 'High' 'No protected VNets' -Rule 'ORPHAN' }
    elseif ($v -le 2) { Add-Rec $row 'DdosProtectionPlans' 'Configure' "Only $v VNet(s) protected - compare with DDoS IP Protection (per public IP)" 'DDoS IP Protection' $null 'Low' "$v VNet(s)" -Rule 'RS-SKU' }
}

$Analyzers.VirtualHubs = { param($row) Set-Analysis $row 'Not assessed' "Hub $($row.HubSku); routing units $($row.RoutingInfraUnits)" "VPN: $($row.VpnGateway); ER: $($row.ErGateway); Firewall: $($row.Firewall)" $null }

#endregion

#region ============================== ANALYZERS: ANALYTICS & AI ==============================

$Analyzers.DataFactories = { param($row)
    $runs = [double]((N $row.PipelineRunsOK_Total) ?? 0) + [double]((N $row.PipelineRunsFailed_Total) ?? 0)
    $prov = "Integration runtimes: $(if ($row.IntegrationRuntimes) { $row.IntegrationRuntimes } else { 'AutoResolve only' })"; $util = "Pipeline runs $(Fmt $runs 0) (failed $(Fmt (N $row.PipelineRunsFailed_Total) 0)); activity runs $(Fmt (N $row.ActivityRuns_Total) 0)"
    if ([int]((N $row.RunningSsisIRs) ?? 0) -gt 0) { Add-Rec $row 'DataFactories' 'Configure' 'Azure-SSIS integration runtime running continuously - schedule start/stop around package executions' 'Scheduled SSIS IR' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Medium' $prov 'Low' 'Schedule SSIS IR: https://learn.microsoft.com/azure/data-factory/how-to-schedule-azure-ssis-integration-runtime' -Rule 'SCHEDULE' }
    if (Test-NoActivity $row @('PipelineRunsOK_Total', 'PipelineRunsFailed_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'DataFactories' 'Delete' "No pipeline runs in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Low' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
}

$Analyzers.SynapseSqlPools = { param($row)
    $p95 = N $row.DWUUsedPct_P95; $active = N $row.Pattern_ActiveHoursPct; $idle = N $row.Pattern_IdleHoursPct
    $prov = "$($row.PoolSku), status $($row.Status)"; $util = "DWU used P95 $(Fmt $p95)%; online $(Fmt $active 0)% of hours, idle $(Fmt $idle 0)% of online hours"
    if ($null -eq $p95) { Set-Analysis $row $(if ($row.Status -eq 'Paused') { 'Stopped' } else { 'No data' }) $prov $util $null; return }
    $status = Get-StatusFromCpu (N $row.DWUUsedPct_Avg) $p95 $null $null
    Set-Analysis $row $status $prov $util $p95
    if (($idle ?? 0) -gt 40) { Add-Rec $row 'SynapseSqlPools' 'Configure' "Pool idle $(Fmt $idle 0)% of online hours - automate pause/resume" 'Pause schedule' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio (1 - $idle / 100) -Share 0.9) 'Medium' $util -Rule 'SCHEDULE' }
    if ($status -in @('Idle', 'Over-provisioned')) { Add-Rec $row 'SynapseSqlPools' 'Rightsize' "DWU P95 $(Fmt $p95)% - scale down the service level" 'Lower DWU' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5 -Share 0.9) 'Medium' $util -Rule 'RS-SKU' }
}

$Analyzers.SynapseSparkPools = { param($row)
    Set-Analysis $row 'Not assessed' "$($row.NodeSize) x $($row.NodeCount) node(s), autoscale $($row.AutoScale), auto-pause $($row.AutoPause) ($($row.AutoPauseDelayMin) min)" 'Billed per session (auto-pause/scale settings reviewed)' $null
    if (-not (B $row.AutoPause)) { Add-Rec $row 'SynapseSparkPools' 'Configure' 'Auto-pause disabled - enable auto-pause (e.g. 15 minutes)' 'Auto-pause on' $null 'Medium' 'Auto-pause disabled' -Rule 'SCHEDULE' }
}

$Analyzers.DataExplorerClusters = { param($row)
    $cpu = N $row.CPU_P95; $cache = N $row.CacheUtilPct_P95; $n = N $row.Instances
    $prov = "$($row.ClusterSku) x $(Fmt $n 0) instance(s)$(if (B $row.OptimizedAutoscale) { " (optimized autoscale $($row.AutoscaleMin)-$($row.AutoscaleMax))" })"
    $util = "CPU P95 $(Fmt $cpu)%; cache utilization P95 $(Fmt $cache)%; ingestion utilization max $(Fmt (N $row.IngestionUtilPct_Max))%"
    if ($row.State -eq 'Stopped') { Set-Analysis $row 'Stopped' $prov 'Cluster stopped' $null; return }
    $status = if ($null -eq $cpu) { 'No data' } elseif ($cpu -lt 20 -and ($cache ?? 0) -lt 50) { 'Over-provisioned' } elseif ($cpu -gt 85 -or ($cache ?? 0) -gt 95) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $cpu
    if ($status -eq 'Over-provisioned' -and $n -gt 2 -and -not (B $row.OptimizedAutoscale)) { Add-Rec $row 'DataExplorerClusters' 'Rightsize' 'Low CPU and cache utilization - enable optimized autoscale / reduce instance count' 'Optimized autoscale' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.6) 'Medium' $util -Rule 'RS-SCALE' }
    if ($row.Environment -eq 'Non-Production' -and -not (B $row.AutoStop)) { Add-Rec $row 'DataExplorerClusters' 'Configure' 'Enable auto-stop for the non-production cluster' 'Auto-stop' $null 'Low' 'Auto-stop disabled' -Rule 'SCHEDULE' }
}

$Analyzers.Databricks = { param($row) Set-Analysis $row 'Not assessed' "$($row.WorkspaceSku) workspace" 'Compute runs in the managed resource group (see VMs/Disks/Cost)' $null }

$Analyzers.HDInsightClusters = { param($row)
    $prov = "$($row.ClusterKind): $($row.Roles)"; $util = "Gateway requests $(Fmt (N $row.GatewayRequests_Total) 0); max active workers $(Fmt (N $row.ActiveWorkers_Max) 0)"
    if (Test-NoActivity $row @('GatewayRequests_Total')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'HDInsightClusters' 'Delete' "No gateway requests in $LookbackDays days - delete the cluster (HDInsight bills while provisioned)" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
    Add-Rec $row 'HDInsightClusters' 'Configure' 'HDInsight bills 24x7 while provisioned - enable autoscale and delete/recreate clusters for batch workloads' 'Autoscale / ephemeral clusters' $null 'Low' $util -Rule 'SCHEDULE'
}

$Analyzers.AnalysisServices = { param($row)
    $q = N $row.QPU_P95
    $prov = "$($row.ServerSku) x $($row.Replicas) replica(s), state $($row.State)"; $util = "QPU P95 $(Fmt $q 1); memory P95 $(Fmt (N $row.MemoryGB_P95) 1) GB"
    if ($row.State -eq 'Paused') { Set-Analysis $row 'Stopped' $prov 'Paused' $null; return }
    Set-Analysis $row $(if ($null -eq $q) { 'No data' } else { 'In use' }) $prov $util $null
    if ($row.Environment -eq 'Non-Production') { Add-Rec $row 'AnalysisServices' 'Configure' 'Non-production server - automate pause outside business hours' 'Pause schedule' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.4) 'Medium' $util -Rule 'SCHEDULE' }
}

$Analyzers.PowerBIFabricCapacities = { param($row)
    $cpu = N $row.CPU_P95
    $prov = "$($row.CapacitySku) ($($row.ArmType -replace '^microsoft\.', '')), state $($row.State)"; $util = if ($null -ne $cpu) { "CPU P95 $(Fmt $cpu)%" } else { 'Use the Fabric Capacity Metrics app for CU utilization' }
    if ($row.State -eq 'Paused') { Set-Analysis $row 'Stopped' $prov 'Paused (not billed)' $null; return }
    Set-Analysis $row $(if ($null -ne $cpu -and $cpu -lt 20) { 'Over-provisioned' } else { 'In use' }) $prov $util $cpu
    if ($row.Environment -ne 'Production') { Add-Rec $row 'PowerBIFabricCapacities' 'Configure' 'Capacity active 24x7 - pause outside business hours (Fabric/Embedded capacities bill while active) or purchase a reservation for steady use' 'Pause schedule / reservation' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio 0.5) 'Low' $util -Rule 'SCHEDULE' }
}

$Analyzers.CognitiveServices = { param($row)
    $calls = N $row.Calls_Total; $ptu = [int]((N $row.TotalPTUs) ?? 0)
    $prov = "$($row.ServiceKind) $($row.AccountSku)$(if ($row.Deployments) { ", $($row.Deployments) deployment(s)" })$(if ($ptu) { ", $ptu PTU provisioned" })"
    $util = "Calls $(Fmt $calls 0)$(if ($null -ne (N $row.PromptTokens_Total)) { "; prompt tokens $(Fmt (N $row.PromptTokens_Total) 0); generated tokens $(Fmt (N $row.GeneratedTokens_Total) 0)" })$(if ($ptu) { "; PTU utilization P95 (max deployment) $(Fmt (N $row.PTUUtilPct_P95_Max))%" })"
    if (Test-NoActivity $row @('Calls_Total')) {
        Set-Analysis $row 'Unused' $prov $util 0
        Add-Rec $row 'CognitiveServices' 'Delete' "No API calls in $LookbackDays days$(if ($ptu) { " while $ptu PTU are provisioned" }) - validate and delete" 'Delete' $row.MonthlyCost $(if ($ptu) { 'High' } else { 'Low' }) $util -Rule 'NOACT'
        return
    }
    Set-Analysis $row 'In use' $prov $util (N $row.PTUUtilPct_P95_Max)
}

$Analyzers.AIModelDeployments = { param($row)
    $p95 = N $row.PTUUtilPct_P95; $ptu = N $row.PTUs
    $prov = "$($row.Model) $($row.ModelVersion) - $($row.DeploymentSku) (capacity $($row.Capacity))"
    if (-not (B $row.IsProvisioned)) { Set-Analysis $row 'Not assessed' $prov 'Pay-per-token deployment' $null; return }
    $util = "PTU utilization avg $(Fmt (N $row.PTUUtilPct_Avg))% | P95 $(Fmt $p95)% | max $(Fmt (N $row.PTUUtilPct_Max))%"
    if ($null -eq $p95) { Set-Analysis $row 'No data' $prov $util $null; return }
    $status = if ($p95 -lt 5) { 'Idle' } elseif ($p95 -lt 50) { 'Over-provisioned' } elseif ($p95 -gt 95) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $p95
    if ($status -in @('Idle', 'Over-provisioned') -and $ptu) {
        $t = [Math]::Max(15, [Math]::Ceiling($ptu * $p95 / 80 / 5) * 5)
        if ($t -lt $ptu) { Add-Rec $row 'AIModelDeployments' 'Rightsize' "Reduce provisioned throughput from $ptu to $t PTU (or move spiky traffic to Standard/Global Standard with spillover)" "$t PTU" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $ptu)) 'Medium' $util -Rule 'RS-SCALE' }
    }
}

$Analyzers.SearchServices = { param($row)
    $qps = N $row.QPS_P95; $rep = [int]((N $row.Replicas) ?? 1)
    $prov = "$($row.SearchSku), $rep replica(s) x $($row.Partitions) partition(s) = $($row.SearchUnits) SU"; $util = "Queries/sec P95 $(Fmt $qps 2) | max $(Fmt (N $row.QPS_Max) 2); throttled $(Fmt (N $row.ThrottledPct_Max))%"
    if ($row.SearchSku -eq 'free') { Set-Analysis $row 'Not assessed' $prov 'Free tier' $null; return }
    if (Test-NoActivity $row @('QPS_Max')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'SearchServices' 'Delete' "No search queries in $LookbackDays days - validate and delete" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    Set-Analysis $row 'In use' $prov $util $null
    if ($rep -gt 1 -and $row.Environment -eq 'Non-Production') { Add-Rec $row 'SearchServices' 'Rightsize' "Non-production service with $rep replicas - use 1 replica" '1 replica' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio (1.0 / $rep)) 'Medium' $util -Rule 'RS-SCALE' }
}

$Analyzers.MLWorkspaces = { param($row) Set-Analysis $row 'Not assessed' "$($row.WorkspaceKind) workspace" 'Compute assessed on the ML Computes sheet' $null }

$Analyzers.MLComputes = { param($row)
    $prov = "$($row.ComputeType) $($row.VmSize)$(if ($row.ComputeType -eq 'AmlCompute') { " nodes $($row.MinNodes)-$($row.MaxNodes)" })"
    Set-Analysis $row $(if ($row.State -match '(?i)stopped|deallocated') { 'Stopped' } else { 'In use' }) $prov "State $($row.State); current nodes $($row.CurrentNodes); idle shutdown $($row.IdleShutdown)" $null
    $hourly = Get-VmHourly $row.Location ("Standard_" + ([string]$row.VmSize -replace '^(?i)standard_', '')) 'Linux'
    if ($row.ComputeType -eq 'AmlCompute' -and [int]((N $row.MinNodes) ?? 0) -gt 0) { Add-Rec $row 'MLComputes' 'Configure' "Compute cluster keeps $($row.MinNodes) node(s) always on - set minimum nodes to 0" 'Min nodes = 0' $(if ($hourly) { [Math]::Round($hourly * (N $row.MinNodes) * 730, 2) } else { $null }) 'High' $prov -Rule 'AUTOSCALE' }
    if ($row.ComputeType -eq 'ComputeInstance' -and [string]::IsNullOrEmpty([string]$row.IdleShutdown) -and $row.State -match '(?i)running') { Add-Rec $row 'MLComputes' 'Configure' 'Compute instance running without idle shutdown - enable idle shutdown / schedule' 'Idle shutdown 60 min' $(if ($hourly) { [Math]::Round($hourly * 730 * 0.6, 2) } else { $null }) 'Medium' $prov -Rule 'SCHEDULE' }
}

$Analyzers.BatchAccounts = { param($row)
    $ded = [int]((N $row.DedicatedNodes) ?? 0); $idlePct = N $row.IdleNodePct_Avg; $tasks = N $row.TasksCompleted_Total; $avgNodes = N $row.TotalNodes_Avg
    $prov = "$($row.Pools) pool(s), $ded dedicated + $($row.LowPriorityNodes) spot node(s) now; $($row.PoolSummary)"
    $util = "Avg nodes $(Fmt $avgNodes 1) | idle $(Fmt $idlePct 0)% of node time | tasks completed $(Fmt $tasks 0)"
    if (($avgNodes ?? 0) -eq 0) { Set-Analysis $row 'Right-sized' $prov 'No nodes allocated (scaled to zero)' 0; return }
    if (($tasks ?? 0) -eq 0) {
        Set-Analysis $row 'Idle' $prov $util 0
        Add-Rec $row 'BatchAccounts' 'Delete' "Pools kept $(Fmt $avgNodes 1) node(s) allocated on average with no tasks completed in $LookbackDays days - scale pools to zero / enable autoscale" 'Scale to zero' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'
        return
    }
    $status = if ($null -eq $idlePct) { 'In use' } elseif ($idlePct -ge 50) { 'Over-provisioned' } elseif ($idlePct -le 5) { 'Right-sized' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $(if ($null -ne $idlePct) { 100 - $idlePct } else { $null })
    if ($status -eq 'Over-provisioned') {
        Add-Rec $row 'BatchAccounts' 'Rightsize' "Nodes idle $(Fmt $idlePct 0)% of the time$(if ([int]((N $row.FixedScalePoolsWithNodes) ?? 0) -gt 0) { " ($($row.FixedScalePoolsWithNodes) fixed-scale pool(s))" }) - use an autoscale formula based on pending tasks and spot nodes where possible" 'Autoscale formula / spot nodes' (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio (1 - $idlePct / 100) -Share 0.8) 'Medium' $util -Rule 'AUTOSCALE'
    }
}

$Analyzers.DedicatedHosts = { param($row)
    $vms = [int]((N $row.VMsOnHost) ?? 0)
    $prov = "$($row.HostSku) host in group $($row.HostGroup)"; $util = "$vms VM(s) on host$(if ($row.RemainingCapacity) { "; remaining capacity: $($row.RemainingCapacity)" })"
    if ($vms -eq 0) { Set-Analysis $row 'Orphaned' $prov 'No VMs placed on the host (host billed in full)' 0; Add-Rec $row 'DedicatedHosts' 'Delete' 'Dedicated host without VMs - the full host is billed; delete it or consolidate VMs' 'Delete host' $row.MonthlyCost 'High' $util -Rule 'ORPHAN'; return }
    Set-Analysis $row 'In use' $prov $util $null
    if ($row.RemainingCapacity) { Add-Rec $row 'DedicatedHosts' 'Configure' "Host has spare capacity ($($row.RemainingCapacity)) - consolidate VMs from other hosts in the group to release a host" 'Consolidate hosts' $null 'Low' $util -Rule 'CONSOLIDATE' }
    if ([string]::IsNullOrEmpty([string]$row.LicenseType) -or $row.LicenseType -eq 'None') { Add-Rec $row 'DedicatedHosts' 'Configure' 'Windows Server / SQL Server licenses not applied at host level - evaluate Azure Hybrid Benefit for dedicated hosts' 'Host license type' $null 'Low' "LicenseType = $($row.LicenseType)" -Rule 'LICENSE' }
}

$Analyzers.CapacityReservations = { param($row)
    $res = [int]((N $row.ReservedQuantity) ?? 0); $alloc = [int]((N $row.VMsAllocated) ?? 0); $unused = [int]((N $row.UnusedReservedVMs) ?? 0)
    $prov = "$res x $($row.VmSize) reserved"; $util = "$alloc VM(s) allocated ($(Fmt (N $row.ReservationUtilizationPct) 0)% of reserved capacity)"
    if ($res -eq 0) { Set-Analysis $row 'Not assessed' $prov 'Reservation quantity 0' $null; return }
    $pct = N $row.ReservationUtilizationPct
    $status = if ($alloc -eq 0) { 'Unused' } elseif ($pct -lt 100) { 'Over-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $pct
    if ($unused -gt 0) {
        $h = Get-VmHourly $row.Location $row.VmSize 'Linux'
        $sav = if ($row.MonthlyCost -gt 0) { [Math]::Round($row.MonthlyCost * $unused / $res, 2) } elseif ($h) { [Math]::Round($h * 730 * $unused, 2) } else { $null }
        Add-Rec $row 'CapacityReservations' $(if ($alloc -eq 0) { 'Delete' } else { 'Rightsize' }) "$unused of $res reserved $($row.VmSize) instance(s) are not used by any VM - unused reserved capacity is billed at the pay-as-you-go rate; reduce the reserved quantity to $alloc" "Quantity $alloc" $sav 'High' $util -Rule $(if ($alloc -eq 0) { 'ORPHAN' } else { 'RS-SCALE' })
    }
}

$Analyzers.ElasticSans = { param($row)
    $provT = N $row.ProvisionedTiB; $allocPct = N $row.AllocatedPct; $usedPct = N $row.UsedPct
    $prov = "$($row.SanSku): $(Fmt $provT 1) TiB provisioned (base $(Fmt (N $row.BaseSizeTiB) 1) + extended $(Fmt (N $row.ExtendedCapacityTiB) 1)), $(Fmt (N $row.ProvisionedIops) 0) IOPS"
    $util = "Volumes allocated $(Fmt (N $row.AllocatedVolumeGiB) 0) GiB ($(Fmt $allocPct 0)%); used $(Fmt $usedPct 0)%; transactions $(Fmt (N $row.Transactions_Total) 0)"
    if ($null -eq $allocPct) { Set-Analysis $row 'No data' $prov $util $null; return }
    $status = if ($allocPct -lt 50) { 'Over-provisioned' } elseif ($allocPct -gt 90) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $allocPct
    if ($status -eq 'Over-provisioned' -and $provT -gt 1) {
        $t = [Math]::Max(1, [Math]::Ceiling([double](N $row.AllocatedVolumeGiB) / 1024 * 1.25))
        if ($t -lt $provT) { Add-Rec $row 'ElasticSans' 'Rightsize' "Only $(Fmt $allocPct 0)% of the provisioned SAN is allocated to volumes - reduce extended / base capacity from $(Fmt $provT 1) to $t TiB" "$t TiB" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $provT)) 'Medium' $util -Rule 'RS-STORAGE' }
    }
}

$Analyzers.SignalRWebPubSub = { param($row)
    $units = [int]((N $row.Units) ?? 1); $load = N $row.ServerLoad_P95; $conn = N $row.ConnectionUtilPct_Max; $msgs = N $row.Messages_Total
    $prov = "$($row.ServiceKind) $($row.ServiceSku) x $units unit(s) ($(Fmt (N $row.ConnectionCapacity) 0) connections)"
    $util = "Server load P95 $(Fmt $load)%; peak connections $(Fmt (N $row.Connections_Max) 0) ($(Fmt $conn)% of capacity); messages $(Fmt $msgs 0)"
    if ($row.Tier -eq 'Free') { Set-Analysis $row 'Not assessed' $prov 'Free tier' $null; return }
    if ((N $row.Connections_Max) -eq 0 -and ($msgs ?? 0) -eq 0 -and $row.MetricsStatus -in @('OK', 'Partial', 'NoData')) { Set-Analysis $row 'Unused' $prov $util 0; Add-Rec $row 'SignalRWebPubSub' 'Delete' "No connections or messages in $LookbackDays days - delete the service" 'Delete' $row.MonthlyCost 'Medium' $util -Rule 'NOACT'; return }
    $pct = [Math]::Max([double]($load ?? 0), [double]($conn ?? 0))
    $status = if ($pct -lt 30) { 'Over-provisioned' } elseif ($pct -gt 85) { 'Under-provisioned' } else { 'Right-sized' }
    Set-Analysis $row $status $prov $util $pct
    if ($status -eq 'Over-provisioned' -and $units -gt 1) { $t = [int][Math]::Max(1, [Math]::Ceiling($units * $pct / $TargetUtilization)); if ($t -lt $units) { Add-Rec $row 'SignalRWebPubSub' 'Rightsize' "Reduce units from $units to $t" "$t unit(s)" (Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $units)) 'Medium' $util -Rule 'RS-SCALE' } }
    elseif ($status -eq 'Under-provisioned') { Add-Rec $row 'SignalRWebPubSub' 'Performance Risk' 'Service running close to unit capacity - add units / enable autoscale' 'Scale out' $null 'Medium' $util -Rule 'PERF-RISK' }
}

# Generic analyzer for every auto-discovered resource type
$GenericRefs = @(Import-RawCsv 'Reference_GenericMetrics.csv')
# Percent-of-service-limit metrics measure headroom (never a right-sizing signal); configuration gauges are not activity
$GenericLimitRx = '(?i)capacityutilization|percentutilization|saturation|shoebox|limit|quota'
function Test-GenericGauge([string]$Name) { return ($Name -match '(?i)(limit|quota)$') -or ($Name -match '(?i)(count|size|capacity)$' -and $Name -notmatch '(?i)request|message|call|transaction|event|quer|connection|session|hit|job|run|operation|invocation|deliver|publish|api|incoming|outgoing') }
function Test-GenericActivity($Ref) { return $Ref.Role -eq 'Activity' -and -not (Test-GenericGauge $Ref.Metric) -and [string]$Ref.Unit -notmatch '(?i)^(milliseconds|seconds)$' -and "$($Ref.Metric) $($Ref.DisplayName)" -notmatch '(?i)latency|duration|response ?time' }
$GenericRefByType = @{}
foreach ($g in $GenericRefs) { if ($g.Alias) { $k = ([string]$g.ArmType).ToLower(); if (-not $GenericRefByType.ContainsKey($k)) { $GenericRefByType[$k] = [System.Collections.Generic.List[object]]::new() }; $GenericRefByType[$k].Add($g) } }
$Analyzers.Generic = { param($row)
    $key = [string]$row.TypeKey
    $util = N $row.GenericUtilP95; $utilName = [string]$row.GenericUtilMetric; $utilAvg = N $row.GenericUtilAvg; $utilMax = N $row.GenericUtilMax
    $limName = [string]$row.GenericLimitMetric; $lim = N $row.GenericLimitP95
    if ($utilName -and $utilName -match $GenericLimitRx) { $limName = $utilName; $lim = $util; $util = $null; $utilName = '' }
    $act = N $row.GenericActivityTotal; $actNames = [string]$row.GenericActivityMetrics
    $refs = $GenericRefByType[([string]$row.ArmType).ToLower()]
    if ($refs -and @($refs | Where-Object Role -eq 'Activity').Count) {
        $sum = 0.0; $n = 0; $names = [System.Collections.Generic.List[string]]::new()
        foreach ($g in ($refs | Where-Object { Test-GenericActivity $_ })) {
            $v = if ($row.PSObject.Properties["$($g.Alias)_Total"]) { N $row."$($g.Alias)_Total" } else { N $row."$($g.Alias)_Max" }
            if ($null -ne $v) { $sum += $v; $n++; $names.Add($g.Alias) }
        }
        $act = if ($n) { $sum } else { $null }; $actNames = $names -join ', '
    }
    $disp = $script:SheetNames[$key] ?? $key
    $metricTxt = @()
    if ($utilName) { $metricTxt += "$utilName P95 $(Fmt $util)% (avg $(Fmt $utilAvg)%, max $(Fmt $utilMax)%)" }
    if ($limName) { $metricTxt += "service-limit headroom: $limName P95 $(Fmt $lim)% of the limit" }
    if ($actNames) { $metricTxt += "activity ($actNames) total $(Fmt $act 0)" }
    $prov = "$disp$(if ($row.SkuName) { " - SKU $($row.SkuName)$(if ($row.SkuTier) { " ($($row.SkuTier))" })$(if ($row.SkuCapacity) { " x $($row.SkuCapacity)" })" })"
    $utilTxt = if ($metricTxt.Count) { $metricTxt -join '; ' } else { [string]$row.MetricsStatus }
    if ([string]$row.MetricsStatus -like 'N/A*') { Set-Analysis $row 'Not assessed' $prov 'No platform utilization metrics (cost tracked)' $null; return }
    if ($row.MetricsStatus -notin @('OK', 'Partial', 'NoData')) { Set-Analysis $row 'No data' $prov $utilTxt $null; return }
    if ($null -ne $act -and $act -eq 0 -and $null -eq $util) {
        Set-Analysis $row 'Unused' $prov $utilTxt 0
        if ((N $row.MonthlyCost) -gt 0 -and (N $row.MonthlyCost) -ge $GenericMinRecCost) { Add-Rec $row $key 'Delete' "No activity on any auto-discovered usage metric ($actNames) in $LookbackDays days while billed - validate and remove" 'Delete' $row.MonthlyCost 'Medium' $utilTxt -Rule 'GEN-NOACT' }
        return
    }
    $nearLimit = $null -ne $lim -and $lim -gt $GenericHigh
    if ($null -ne $util) {
        $status = if ($util -lt $CpuIdleThreshold -and ($act ?? 0) -eq 0) { 'Idle' } elseif ($util -gt $GenericHigh -or $nearLimit) { 'Under-provisioned' } elseif ($util -lt $GenericLow) { 'Over-provisioned' } else { 'Right-sized' }
        Set-Analysis $row $status $prov $utilTxt $util
        $cap = N $row.SkuCapacity; $cost = N $row.MonthlyCost
        if ($status -in @('Idle', 'Over-provisioned') -and $cost -gt 0 -and ($cap -gt 1 -or $cost -ge 5)) {
            $t = if ($cap -gt 1) { [int][Math]::Max(1, [Math]::Ceiling($cap * [Math]::Max($util, 1) / $TargetUtilization)) } else { $null }
            $sav = if ($t -and $t -lt $cap) { Get-Savings -MonthlyCost $row.MonthlyCost -Ratio ($t / $cap) } else { $null }
            $action = if ($t -and $t -lt $cap) { "$utilName P95 is $(Fmt $util)% of current capacity - reduce capacity units from $(Fmt $cap 0) to $t" } else { "$utilName P95 is only $(Fmt $util)% of the current SKU capacity - evaluate a smaller SKU / tier" }
            Add-Rec $row $key 'Rightsize' $action $(if ($t) { "$t unit(s)" } else { 'Smaller SKU' }) $sav $(if ($t) { 'Medium' } else { 'Low' }) $utilTxt -Rule 'GEN-UTIL'
        }
        elseif ($status -eq 'Under-provisioned') {
            $what = if ($util -gt $GenericHigh) { "$utilName P95 at $(Fmt $util)% - capacity close to its limit" } else { "$limName P95 at $(Fmt $lim)% of the service limit" }
            Add-Rec $row $key 'Performance Risk' $what $(if ($util -gt $GenericHigh) { 'Scale up / out' } else { 'Request a limit increase or split the workload' }) $null 'Medium' $utilTxt -Rule 'GEN-HOT'
        }
        return
    }
    if ($nearLimit) {
        Set-Analysis $row 'Under-provisioned' $prov $utilTxt $null
        Add-Rec $row $key 'Performance Risk' "$limName P95 at $(Fmt $lim)% of the service limit - throttling risk" 'Request a limit increase or split the workload' $null 'Medium' $utilTxt -Rule 'GEN-HOT'
        return
    }
    Set-Analysis $row $(if (($act ?? 0) -gt 0) { 'In use' } elseif ($null -ne $lim) { 'In use' } else { 'No data' }) $prov $utilTxt $null
}

$Analyzers.AzureVirtualDesktop = { param($row)
    Set-Analysis $row 'Not assessed' "$($row.HostPoolType) host pool ($($row.LoadBalancer)), max sessions $($row.MaxSessionLimit)" 'Session hosts analyzed on the Virtual Machines sheet' $null
    if (-not (B $row.StartVMOnConnect)) { Add-Rec $row 'AzureVirtualDesktop' 'Configure' 'Enable Start VM on Connect and an autoscale scaling plan so session hosts are deallocated when not in use' 'Scaling plan + Start VM on Connect' $null 'Medium' 'Start VM on Connect disabled' 'Low' 'AVD autoscale: https://learn.microsoft.com/azure/virtual-desktop/autoscale-scaling-plan' -Rule 'SCHEDULE' }
}

$Analyzers.Default = { param($row) Set-Analysis $row 'Not assessed' '' 'Inventory only' $null }

#endregion

#region ============================== RUN ANALYSIS ==============================

function Rows([string]$k) { if ($TypeData.Contains($k)) { return @($TypeData[$k] | Where-Object { $_ }) }; return @() }

Write-Log 'Analyzing provisioned vs. utilized and generating recommendations...' Section
$script:SqlVmIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($s in (Rows 'SqlVirtualMachines')) { if ($s.VirtualMachineId) { [void]$script:SqlVmIds.Add(([string]$s.VirtualMachineId).ToLower()) } }

# Common enrichment: type key, environment, monthly run-rate cost, portal link
foreach ($key in $TypeData.Keys) {
    foreach ($row in $TypeData[$key]) {
        $row | Add-Member -NotePropertyName TypeKey -NotePropertyValue $key -Force
        $row | Add-Member -NotePropertyName Environment -NotePropertyValue (Get-Environment $row) -Force
        $row | Add-Member -NotePropertyName MonthlyCost -NotePropertyValue (Get-MonthlyCost $row.ResourceId) -Force
        $row | Add-Member -NotePropertyName PortalLink -NotePropertyValue $(if ($row.ResourceId) { "https://portal.azure.com/#@$($Meta.TenantId)/resource$($row.ResourceId)" } else { $null }) -Force
    }
}
# Cost allocation for rows that do not map 1:1 to a billed resource. Child rows (AKS node pools, provisioned file shares,
# PTU deployments) carry their share of a parent's cost; the duplicate side is flagged ExcludeFromTotals so that
# cross-type totals count each billed resource once.
$aksVmss = @{}
foreach ($v in (Rows 'VMScaleSets')) {
    if ($v.ManagedByService -eq 'AKS' -and $v.Tags -match 'aks-managed-poolName=([^;]+)') { $aksVmss["$($v.ResourceGroup)|$($Matches[1].Trim())".ToLower()] = $v }
}
$seenCluster = @{}
foreach ($p in (Rows 'AKSNodePools')) {
    $clusterCost = $p.MonthlyCost
    $v = $aksVmss["$($p.NodeResourceGroup)|$($p.NodePool)".ToLower()]
    $p.MonthlyCost = if ($v) { $v.MonthlyCost } else { $null }
    if ($v) { $v | Add-Member -NotePropertyName ExcludeFromTotals -NotePropertyValue $true -Force }
    $cid = ([string]$p.ResourceId).ToLower()
    $p | Add-Member -NotePropertyName ClusterMonthlyCost -NotePropertyValue $(if (-not $seenCluster.ContainsKey($cid)) { $seenCluster[$cid] = $true; $clusterCost } else { $null }) -Force
}
foreach ($grp in ((Rows 'PremiumFileShares') | Group-Object ResourceId)) {
    $acct = Get-MonthlyCost $grp.Name; $tot = [double](($grp.Group | ForEach-Object { [double]((N $_.ProvisionedGiB) ?? 0) } | Measure-Object -Sum).Sum)
    foreach ($s in $grp.Group) {
        $s.MonthlyCost = if ($null -ne $acct -and $tot -gt 0) { [Math]::Round($acct * (N $s.ProvisionedGiB) / $tot, 2) } else { $null }
        $s | Add-Member -NotePropertyName ExcludeFromTotals -NotePropertyValue $true -Force
    }
}
foreach ($grp in ((Rows 'AIModelDeployments') | Group-Object ResourceId)) {
    $acct = Get-MonthlyCost $grp.Name; $tot = [double](($grp.Group | ForEach-Object { [double]((N $_.PTUs) ?? 0) } | Measure-Object -Sum).Sum)
    foreach ($d in $grp.Group) {
        $d.MonthlyCost = if ($null -ne $acct -and $tot -gt 0 -and (N $d.PTUs)) { [Math]::Round($acct * (N $d.PTUs) / $tot, 2) } else { $null }
        $d | Add-Member -NotePropertyName ExcludeFromTotals -NotePropertyValue $true -Force
    }
}
function Get-TotalCost($row) {
    if ($row.PSObject.Properties['ExcludeFromTotals'] -and $row.ExcludeFromTotals) { return 0.0 }
    $cluster = if ($row.PSObject.Properties['ClusterMonthlyCost']) { [double]((N $row.ClusterMonthlyCost) ?? 0) } else { 0.0 }
    return [double]((N $row.MonthlyCost) ?? 0) + $cluster
}
foreach ($m in (Rows 'MLComputes')) { $m.MonthlyCost = $null }
foreach ($d in (Rows 'ManagedDisks')) {
    if ($null -eq $d.MonthlyCost -and $d.DiskTier) { $d.MonthlyCost = Get-DiskMonthly $d.Location $d.DiskTier ((([string]$d.DiskSku) -split '_')[-1]) }
    if ($d.AttachedToId) { $k = ([string]$d.AttachedToId).ToLower(); $script:DiskCostByVm[$k] = [Math]::Round([double]($script:DiskCostByVm[$k] ?? 0) + [double]($d.MonthlyCost ?? 0), 2) }
}
foreach ($a in (Rows 'AppServices')) {
    $k = ([string]$a.AppServicePlanId).ToLower(); if (-not $k) { continue }
    if (-not $script:AppsByPlan.ContainsKey($k)) { $script:AppsByPlan[$k] = @{ Total = 0; Stopped = 0; Names = '' } }
    $e = $script:AppsByPlan[$k]; $e.Total++; if ($a.State -eq 'Stopped') { $e.Stopped++ }
    if ($e.Total -le 5) { $e.Names = (@($e.Names, $a.ResourceName) | Where-Object { $_ }) -join ', ' }
}

$order = @('ManagedDisks', 'VirtualMachines', 'AppServices') + @($TypeData.Keys | Where-Object { $_ -notin @('ManagedDisks', 'VirtualMachines', 'AppServices', 'SqlVirtualMachines', 'LogAnalyticsClusters') }) + @('SqlVirtualMachines', 'LogAnalyticsClusters')
foreach ($key in $order) {
    if (-not $TypeData.Contains($key)) { continue }
    $fn = $Analyzers[$key] ?? $(if ($script:GenericKeys.Contains($key)) { $Analyzers.Generic } else { $Analyzers.Default })
    $i = 0
    foreach ($row in $TypeData[$key]) {
        $i++
        try { & $fn $row } catch { Write-Log "  Analyzer error [$key] $($row.ResourceName): $($_.Exception.Message)" Warning; Set-Analysis $row 'Not assessed' '' "Analysis error: $($_.Exception.Message)" $null }
        if ($key -eq 'VirtualMachines') { $script:VmById[([string]$row.ResourceId).ToLower()] = $row }
    }
    Write-Log "  $($script:SheetNames[$key] ?? $key): $i row(s) analyzed" Info
}

# Priorities + de-duplicated savings per resource row (Delete supersedes; otherwise capped at current cost)
foreach ($r in $script:Recs) {
    $a = $r.EstAnnualSavings
    $r.Priority = if ($r.Category -eq 'Performance Risk') { 'Medium' } elseif ($null -ne $a -and $a -ge 10000) { 'High' } elseif ($null -ne $a -and $a -ge 1000) { 'Medium' } elseif ($r.Category -eq 'Delete' -and $r.Confidence -eq 'High') { 'Medium' } else { 'Low' }
}
$recGroups = @{}
foreach ($r in $script:Recs) { $k = "$($r.TypeKey)|$($r.ResourceId)|$($r.ResourceName)".ToLower(); if (-not $recGroups.ContainsKey($k)) { $recGroups[$k] = [System.Collections.Generic.List[object]]::new() }; $recGroups[$k].Add($r) }
# De-duplicated savings for any set of recommendations. Per resource a Delete supersedes the other actions and the total is
# capped at the current cost - except DEALLOC, whose savings are the attached disks that are billed separately. Child rows
# (file shares, PTU deployments) are excluded when their parent account is deleted in the same set (the account cost includes them).
function Get-DedupSavings([object[]]$Recs) {
    $res = [ordered]@{ Delete = 0.0; Rightsize = 0.0; Configure = 0.0; Total = 0.0; ByKey = @{} }
    $parentDel = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in $Recs) { if ($r.Category -eq 'Delete' -and $r.TypeKey -in @('StorageAccounts', 'CognitiveServices') -and $r.ResourceId) { [void]$parentDel.Add(([string]$r.ResourceId).ToLower()) } }
    $groups = @{}
    foreach ($r in $Recs) { $k = "$($r.TypeKey)|$($r.ResourceId)|$($r.ResourceName)".ToLower(); if (-not $groups.ContainsKey($k)) { $groups[$k] = [System.Collections.Generic.List[object]]::new() }; $groups[$k].Add($r) }
    foreach ($k in $groups.Keys) {
        $g = $groups[$k]; $first = $g[0]
        if ($first.TypeKey -in @('PremiumFileShares', 'AIModelDeployments') -and $parentDel.Contains(([string]$first.ResourceId).ToLower())) { $res.ByKey[$k] = 0.0; continue }
        $cap = if (@($g | Where-Object RuleCode -eq 'DEALLOC').Count) { $null } else { N $first.CurrentMonthlyCost }
        $del = @($g | Where-Object { $_.Category -eq 'Delete' -and $null -ne $_.EstMonthlySavings })
        if ($del.Count) {
            $v = [double](($del | Measure-Object -Property EstMonthlySavings -Maximum).Maximum)
            if ($null -ne $cap -and $cap -gt 0 -and $v -gt $cap) { $v = $cap }
            $res.Delete += $v; $res.Total += $v; $res.ByKey[$k] = $v; continue
        }
        $sa = [double](($g | Where-Object { $_.Category -eq 'Rightsize' -and $null -ne $_.EstMonthlySavings } | Measure-Object -Property EstMonthlySavings -Sum).Sum)
        $sb = [double](($g | Where-Object { $_.Category -eq 'Configure' -and $null -ne $_.EstMonthlySavings } | Measure-Object -Property EstMonthlySavings -Sum).Sum)
        if ($null -ne $cap -and $cap -gt 0 -and ($sa + $sb) -gt $cap) { $scale = $cap / ($sa + $sb); $sa *= $scale; $sb *= $scale }
        $res.Rightsize += $sa; $res.Configure += $sb; $res.Total += $sa + $sb; $res.ByKey[$k] = $sa + $sb
    }
    return $res
}
$script:DedupAll = Get-DedupSavings $script:Recs.ToArray()
$script:DedupMonthly = [double]$script:DedupAll.Total
$RowSavings = @{}
foreach ($k in $script:DedupAll.ByKey.Keys) { $RowSavings[$k] = [Math]::Round([double]$script:DedupAll.ByKey[$k], 2) }
foreach ($key in $TypeData.Keys) {
    foreach ($row in $TypeData[$key]) {
        $nm = [string]$row.ResourceName
        foreach ($sub in @('NodePool', 'ShareName', 'Deployment', 'ComputeName')) { if ($row.PSObject.Properties[$sub] -and $row.$sub) { $nm = "$nm / $($row.$sub)"; break } }
        $k = "$key|$($row.ResourceId)|$nm".ToLower()
        $g = $recGroups[$k]
        $gs = if ($g) { @($g | Sort-Object @{ e = { @{ 'Delete' = 0; 'Rightsize' = 1; 'Configure' = 2; 'Performance Risk' = 3 }[$_.Category] } }, @{ e = { [double]($_.EstAnnualSavings ?? 0) }; Descending = $true }) } else { @() }
        $row | Add-Member -NotePropertyName SuggestedRecommendation -NotePropertyValue $(if ($g) { ($gs | ForEach-Object { "[$($_.Category)] $($_.RecommendedAction)" }) -join ' || ' } else { 'No action - utilization within thresholds' }) -Force
        $row | Add-Member -NotePropertyName RecommendationCategory -NotePropertyValue $(if ($g) { ($gs | Select-Object -ExpandProperty Category -Unique) -join ', ' } else { '' }) -Force
        $row | Add-Member -NotePropertyName RuleIds -NotePropertyValue $(if ($g) { ($gs | Select-Object -ExpandProperty RuleId -Unique) -join ', ' } else { '' }) -Force
        $row | Add-Member -NotePropertyName RecommendationCriteria -NotePropertyValue $(if ($g) { ($gs | ForEach-Object { "$($_.RuleId): $($_.Criteria)" } | Select-Object -Unique) -join ' || ' } else { '' }) -Force
        $row | Add-Member -NotePropertyName RecommendationPriority -NotePropertyValue $(if ($g) { @('High', 'Medium', 'Low') | Where-Object { $_ -in @($gs | ForEach-Object Priority) } | Select-Object -First 1 } else { '' }) -Force
        $row | Add-Member -NotePropertyName RecommendationConfidence -NotePropertyValue $(if ($g) { $gs[0].Confidence } else { '' }) -Force
        $row | Add-Member -NotePropertyName EstMonthlySavings -NotePropertyValue $(if ($g) { $RowSavings[$k] } else { $null }) -Force
        $row | Add-Member -NotePropertyName EstAnnualSavings -NotePropertyValue $(if ($g -and $null -ne $RowSavings[$k]) { [Math]::Round($RowSavings[$k] * 12, 2) } else { $null }) -Force
        if (-not $row.PSObject.Properties['UtilizationStatus']) { Set-Analysis $row 'Not assessed' '' '' $null }
        # Utilization efficiency score (0-100): P95 utilization vs the target, idle/unused = 0, stopped/not assessed = n/a
        $st = [string]$row.UtilizationStatus; $u = N $row.UtilizationPct
        $score = if ($st -in @('Idle', 'Unused', 'Orphaned', 'Stopped (billed)')) { 0 }
        elseif ($st -in @('Stopped', 'Deallocated', 'Not assessed', 'No data', 'Managed by service', 'Aged', 'Intermittent')) { $null }
        elseif ($null -ne $u) { [Math]::Round([Math]::Min(100.0, 100 * $u / $TargetUtilization), 0) }
        elseif ($st -in @('Right-sized', 'In use', 'Under-provisioned', 'Under-committed')) { 100 }
        elseif ($st -in @('Over-provisioned', 'Over-committed')) { 30 } else { $null }
        $row | Add-Member -NotePropertyName EfficiencyScore -NotePropertyValue $score -Force
    }
}
$RecsSorted = @($script:Recs | Sort-Object @{ e = { [double]($_.EstAnnualSavings ?? -1) }; Descending = $true }, Category)
$AnnualSavings = [Math]::Round($script:DedupMonthly * 12, 0)
Write-Log "Recommendations: $($script:Recs.Count); estimated savings $(Money $AnnualSavings)/year (de-duplicated per resource)" Success

# Paid resources not covered by a dedicated analysis (ensures every billed resource is visible)
$analyzedIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($key in $TypeData.Keys) { foreach ($row in $TypeData[$key]) { if ($row.ResourceId) { [void]$analyzedIds.Add(([string]$row.ResourceId).ToLower()) } } }
$OtherPaid = @(foreach ($c in ($CostByResource | Group-Object { ([string]$_.ResourceId).ToLower() })) {
        if ($analyzedIds.Contains($c.Name) -or [string]::IsNullOrEmpty($c.Name)) { continue }
        $cost = [double](($c.Group | ForEach-Object { [double]((N $(if ($_.CostUSD) { $_.CostUSD } else { $_.Cost })) ?? 0) } | Measure-Object -Sum).Sum) * $MonthFactor
        if ($cost -lt 1) { continue }
        $f = $c.Group[0]
        [pscustomobject][ordered]@{
            SubscriptionName = $f.SubscriptionName; ResourceGroup = $f.ResourceGroup; ResourceName = $f.ResourceName
            ServiceName = (($c.Group | Select-Object -ExpandProperty ServiceName -Unique) -join ', '); MonthlyCost = [Math]::Round($cost, 2)
            AnnualCost = [Math]::Round($cost * 12, 0); ResourceType = if ($c.Name -match '/providers/([^/]+/[^/]+)') { $Matches[1] } else { '' }; ResourceId = $c.Name
        }
    }) | Sort-Object MonthlyCost -Descending

#endregion

#region ============================== EXCEL HELPERS ==============================

function Col([string]$Hex) { return [System.Drawing.ColorTranslator]::FromHtml($Hex) }
function Get-ColLetter([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][Math]::Floor(($n - $m) / 26) }; return $s }

$StatusColors = [ordered]@{
    'Idle' = @('#F8D7DA', '#842029'); 'Unused' = @('#F8D7DA', '#842029'); 'Orphaned' = @('#F8D7DA', '#842029'); 'Stopped (billed)' = @('#F8D7DA', '#842029')
    'Over-committed' = @('#FFE5D0', '#8A4B08'); 'Over-provisioned' = @('#FFE5D0', '#8A4B08'); 'Deallocated' = @('#FFF3CD', '#664D03'); 'Aged' = @('#FFF3CD', '#664D03')
    'Commitment opportunity' = @('#CFE2FF', '#084298'); 'Under-committed' = @('#E2D9F3', '#432874'); 'Under-provisioned' = @('#E2D9F3', '#432874')
    'Right-sized' = @('#D1E7DD', '#0F5132'); 'In use' = @('#D1E7DD', '#0F5132'); 'No data' = @('#E9ECEF', '#495057'); 'Not assessed' = @('#E9ECEF', '#495057')
}
$CategoryColors = @{ 'Delete' = @('#F8D7DA', '#842029'); 'Rightsize' = @('#FFE5D0', '#8A4B08'); 'Configure' = @('#CFE2FF', '#084298'); 'Performance Risk' = @('#E2D9F3', '#432874') }

function Format-Worksheet {
    param($Pkg, [string]$Sheet, [int]$HeaderRow = 1, [int]$FirstCol = 1, [int]$LastCol = 0, [int]$LastRow = 0)
    $ws = $Pkg.Workbook.Worksheets[$Sheet]
    if (-not $ws -or -not $ws.Dimension) { return }
    # Scope every format to this table only - several tables can share a sheet (e.g. the Executive Dashboard)
    $last = if ($LastRow -gt 0) { $LastRow } else { $ws.Dimension.End.Row }
    $cols = if ($LastCol -gt 0) { $LastCol } else { $ws.Dimension.End.Column }
    if ($last -le $HeaderRow) { return }
    for ($c = $FirstCol; $c -le $cols; $c++) {
        $h = [string]$ws.Cells[$HeaderRow, $c].Value
        if (-not $h) { continue }
        if ($h.StartsWith('Prop.') -or $h.StartsWith('Plan.')) { $ws.Column($c).Width = 22; continue }
        if ($h.StartsWith('Tag.')) { $ws.Column($c).Width = 18; continue }
        $L = Get-ColLetter $c; $rng = "$L$($HeaderRow + 1):$L$last"
        if ($h -match '(?i)(Savings|Cost|MonthlyEst|PaygMonthly)$') { Set-ExcelRange -Worksheet $ws -Range $rng -NumberFormat '$#,##0' }
        elseif ($h -match '(?i)(Pct|Percent|Pct_\w+)$|^UtilizationPct$') { Set-ExcelRange -Worksheet $ws -Range $rng -NumberFormat '0.0' }
        if ($h -in @('UtilizationStatus', 'Status')) {
            foreach ($s in $StatusColors.Keys) { Add-ConditionalFormatting -Worksheet $ws -Address $rng -RuleType Equal -ConditionValue "`"$s`"" -BackgroundColor (Col $StatusColors[$s][0]) -ForegroundColor (Col $StatusColors[$s][1]) }
        }
        elseif ($h -in @('Category', 'RecommendationCategory')) {
            foreach ($s in $CategoryColors.Keys) { Add-ConditionalFormatting -Worksheet $ws -Address $rng -RuleType ContainsText -ConditionValue $s -BackgroundColor (Col $CategoryColors[$s][0]) -ForegroundColor (Col $CategoryColors[$s][1]) }
        }
        elseif ($h -in @('Priority', 'RecommendationPriority')) {
            Add-ConditionalFormatting -Worksheet $ws -Address $rng -RuleType Equal -ConditionValue '"High"' -BackgroundColor (Col '#F8D7DA') -ForegroundColor (Col '#842029')
            Add-ConditionalFormatting -Worksheet $ws -Address $rng -RuleType Equal -ConditionValue '"Medium"' -BackgroundColor (Col '#FFF3CD') -ForegroundColor (Col '#664D03')
        }
        elseif ($h -in @('UtilizationPct', 'CommitmentUtilizationPct', 'AvgUtilizationPct', 'TagCoveragePct', 'CostCoveragePct')) { Add-ConditionalFormatting -Worksheet $ws -Address $rng -DataBarColor (Col '#5B9BD5') }
        elseif ($h -in @('EfficiencyScore')) {
            $r = Add-ConditionalFormatting -Worksheet $ws -Address $rng -RuleType ThreeColorScale -PassThru
            if ($r) { $r.LowValue.Color = Col '#F8696B'; $r.MiddleValue.Color = Col '#FFEB84'; $r.HighValue.Color = Col '#63BE7B' }
        }
        elseif ($h -match '(?i)^(EstAnnualSavings)$') { Add-ConditionalFormatting -Worksheet $ws -Address $rng -DataBarColor (Col '#70AD47') }
        if ($h -match '(?i)^(ProvisionedSummary|UtilizedSummary|SuggestedRecommendation|RecommendationCriteria|RecommendedAction|Evidence|Criteria|Rationale|ObservedUtilization|CurrentConfiguration|Comment|Tags|ResourceGroupTags|ResourceId|TopTables|Detail|Problem|Solution|PoolSummary|Description)$') {
            $ws.Column($c).Width = if ($h -in @('Tags', 'ResourceGroupTags', 'ResourceId')) { 40 } else { 60 }; $ws.Column($c).Style.WrapText = $false
        }
    }
    $hdr = $ws.Cells["$(Get-ColLetter $FirstCol)$($HeaderRow):$(Get-ColLetter $cols)$($HeaderRow)"]
    $hdr.Style.Fill.PatternType = 'Solid'; $hdr.Style.Fill.BackgroundColor.SetColor((Col '#0F4C81'))
    $hdr.Style.Font.Color.SetColor([System.Drawing.Color]::White); $hdr.Style.Font.Bold = $true
}

function Add-Heatmap($ws, [string]$Address) {
    $r = Add-ConditionalFormatting -Worksheet $ws -Address $Address -RuleType ThreeColorScale -PassThru
    if ($r) { $r.LowValue.Color = Col '#F7FBFF'; $r.MiddleValue.Color = Col '#9ECAE1'; $r.HighValue.Color = Col '#08519C' }
}

$script:UsedSheetNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
function Get-SheetName([string]$Name) {
    $s = ($Name -replace '[\[\]:*?/\\]', '-').Trim(); if ($s.Length -gt 31) { $s = $s.Substring(0, 31) }
    $base = $s; $i = 2
    while (-not $script:UsedSheetNames.Add($s)) { $suffix = " ($i)"; $s = $base.Substring(0, [Math]::Min($base.Length, 31 - $suffix.Length)) + $suffix; $i++ }
    return $s
}

function Export-Sheet {
    param($Pkg, [object[]]$Data, [string]$Sheet, [string]$Title, [int]$StartRow = 1, [int]$StartColumn = 1, [string]$TableStyle = 'Medium2', [switch]$NoFormat, [int[]]$FreezePane)
    if (-not $Data -or $Data.Count -eq 0) { return $Pkg }
    $tbl = ('T_' + ($Sheet -replace '[^A-Za-z0-9]', '') + "_$StartRow`_$StartColumn")
    $p = @{ ExcelPackage = $Pkg; WorksheetName = $Sheet; TableName = $tbl; TableStyle = $TableStyle; StartRow = $StartRow; StartColumn = $StartColumn; PassThru = $true; BoldTopRow = $true }
    if ($Title) { $p.Title = $Title; $p.TitleBold = $true; $p.TitleSize = 13 }
    if ($Data.Count -le 3000) { $p.AutoSize = $true; $p.MaxAutoSizeRows = 500 }
    if ($FreezePane) { $p.FreezePane = $FreezePane } elseif ($StartRow -eq 1 -and -not $Title) { $p.FreezeTopRow = $true }
    $Pkg = $Data | Export-Excel @p
    if (-not $NoFormat) {
        $hdrRow = $StartRow + $(if ($Title) { 1 } else { 0 })
        $nCols = @($Data[0].PSObject.Properties).Count
        Format-Worksheet -Pkg $Pkg -Sheet $Sheet -HeaderRow $hdrRow -FirstCol $StartColumn -LastCol ($StartColumn + $nCols - 1) -LastRow ($hdrRow + $Data.Count)
    }
    return $Pkg
}

$LeadCols = @('SubscriptionName', 'ResourceGroup', 'ResourceName', 'NodePool', 'ShareName', 'Deployment', 'ComputeName', 'Location', 'Environment',
    'UtilizationStatus', 'UtilizationPct', 'EfficiencyScore', 'ProvisionedSummary', 'UtilizedSummary',
    'SuggestedRecommendation', 'RecommendationCategory', 'RecommendationPriority', 'RecommendationConfidence', 'RuleIds', 'RecommendationCriteria',
    'EstMonthlySavings', 'EstAnnualSavings', 'MonthlyCost', 'ClusterMonthlyCost', 'RetailPaygMonthly', 'EffectiveDiscountPct', 'AnalysisNote',
    'SkuName', 'SkuTier', 'SkuCapacity', 'Kind', 'CreatedTime', 'AgeDays', 'LastChangedTime')
$TailCols = @('MetricsStatus', 'MetricsErrors', 'MetricCalls', 'Tags', 'ResourceGroupTags', 'PortalLink', 'SubscriptionId', 'ResourceId')
$HiddenCols = @('ResourceType', 'ArmType', 'TypeKey', 'ExcludeFromTotals')
function Get-OrderedRows([object[]]$Rows) {
    if (-not $Rows -or $Rows.Count -eq 0) { return @() }
    $all = [System.Collections.Generic.List[string]]::new(); $seen = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in $Rows) { foreach ($n in $r.PSObject.Properties.Name) { if ($seen.Add($n)) { $all.Add($n) } } }
    $lead = @($LeadCols | Where-Object { $seen.Contains($_) })
    $core = @($all | Where-Object { $_ -notin $LeadCols -and $_ -notin $TailCols -and $_ -notin $HiddenCols -and -not $_.StartsWith('Tag.') -and -not $_.StartsWith('Prop.') -and -not $_.StartsWith('Plan.') })
    $tags = @($all | Where-Object { $_.StartsWith('Tag.') })
    $props = @($all | Where-Object { $_.StartsWith('Prop.') -or $_.StartsWith('Plan.') })
    $tail = @($TailCols | Where-Object { $seen.Contains($_) })
    return @($Rows | Select-Object -Property ($lead + $core + $tags + $props + $tail))
}

$WindowDates = @(0..($LookbackDays - 1) | ForEach-Object { ([DateTime]::ParseExact($Meta.WindowStartLocal, 'yyyy-MM-dd', $null)).AddDays($_).ToString('yyyy-MM-dd') })
$DailyCache = @{}
function Get-DailyRows([string]$Key) {
    if ($DailyCache.ContainsKey($Key)) { return $DailyCache[$Key] }
    $p = Join-Path $RawPath 'DailyUtilization' "$($Key)_Daily.csv"
    $d = if (Test-Path $p) { @(Import-Csv -Path $p -Encoding utf8) } else { @() }
    $DailyCache[$Key] = $d
    return $d
}
function New-DailyPivot([object[]]$Daily, [string]$Metric, [string]$Field) {
    $sel = @($Daily | Where-Object { $_.Metric -eq $Metric })
    $hasDim = @($sel | Where-Object { $_.Dimension }).Count -gt 0
    foreach ($g in ($sel | Group-Object { "$($_.SubscriptionName)|$($_.ResourceGroup)|$($_.ResourceName)|$($_.Dimension)" })) {
        $f = $g.Group[0]; $byDate = @{}
        foreach ($x in $g.Group) { $byDate[$x.Date] = N $x.$Field }
        $vals = @($byDate.Values | Where-Object { $null -ne $_ })
        $o = [ordered]@{ SubscriptionName = $f.SubscriptionName; ResourceGroup = $f.ResourceGroup; ResourceName = $f.ResourceName }
        if ($hasDim) { $o.Dimension = $f.Dimension }
        $o.WindowAverage = if ($vals.Count) { Rnd (($vals | Measure-Object -Average).Average) 2 } else { $null }
        $o.WindowPeak = if ($vals.Count) { Rnd (($vals | Measure-Object -Maximum).Maximum) 2 } else { $null }
        foreach ($d in $WindowDates) { $o[$d] = Rnd $byDate[$d] 2 }
        [pscustomobject]$o
    }
}
function Get-MetricField([object[]]$Rows, [string]$Metric) {
    $names = if ($Rows.Count) { $Rows[0].PSObject.Properties.Name } else { @() }
    if ("$($Metric)_Avg" -in $names -or "$($Metric)_P95" -in $names) { return 'Average' }
    if ("$($Metric)_Total" -in $names) { return 'Total' }
    return 'Average'
}
$script:DayNamesShort = @('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat')
function Get-HowMatrix([string]$Key) {
    $p = Join-Path $RawPath 'UsagePatterns' "$($Key)_HourOfWeek.csv"
    if (-not (Test-Path $p)) { return $null }
    $rows = @(Import-Csv -Path $p -Encoding utf8)
    $best = $rows | Group-Object Metric | Sort-Object { [int](N ($_.Group[0].Resources)) } -Descending | Select-Object -First 1
    if (-not $best) { return $null }
    $lookup = @{}; foreach ($x in $best.Group) { $lookup["$($x.DayIndex)|$($x.Hour)"] = Rnd (N $x.AvgValue) 3 }
    $m = foreach ($di in 0..6) {
        $o = [ordered]@{ Day = $script:DayNamesShort[$di] }
        foreach ($h in 0..23) { $o[('{0:00}:00' -f $h)] = $lookup["$di|$h"] }
        [pscustomobject]$o
    }
    return [pscustomobject]@{ Metric = $best.Name; Resources = [int](N $best.Group[0].Resources); Matrix = @($m) }
}

#endregion

#region ============================== SUMMARY DATASETS ==============================

$StatusGroups = [ordered]@{
    'Idle / Unused / Orphaned' = @('Idle', 'Unused', 'Orphaned', 'Stopped (billed)')
    'Over-provisioned'         = @('Over-provisioned', 'Over-committed')
    'Right-sized / In use'     = @('Right-sized', 'In use', 'Intermittent')
    'Under-provisioned'        = @('Under-provisioned', 'Under-committed', 'Commitment opportunity')
    'Stopped / Deallocated'    = @('Stopped', 'Deallocated', 'Aged')
    'No data / Not assessed'   = @('No data', 'Not assessed', 'Managed by service')
}
function Get-StatusGroup([string]$s) { foreach ($k in $StatusGroups.Keys) { if ($s -in $StatusGroups[$k]) { return $k } }; return 'No data / Not assessed' }
function Get-WeightedScore([object[]]$Rows) {
    $w = 0.0; $sw = 0.0; $plain = [System.Collections.Generic.List[double]]::new()
    foreach ($r in $Rows) {
        $s = N $r.EfficiencyScore; if ($null -eq $s) { continue }
        $plain.Add($s); $c = Get-TotalCost $r
        if ($c -gt 0) { $w += $s * $c; $sw += $c }
    }
    if ($sw -gt 0) { return [Math]::Round($w / $sw, 0) }
    if ($plain.Count) { return [Math]::Round(($plain | Measure-Object -Average).Average, 0) }
    return $null
}

$AllRows = [System.Collections.Generic.List[object]]::new()
$TypeSummary = foreach ($key in $TypeData.Keys) {
    $rows = @(Rows $key); if ($rows.Count -eq 0) { continue }
    foreach ($r in $rows) { $AllRows.Add($r) }
    $o = [ordered]@{ Category = $script:TypeCategory[$key] ?? 'Other Services'; ResourceType = $script:SheetNames[$key] ?? $key; Resources = $rows.Count }
    foreach ($g in $StatusGroups.Keys) { $o[$g] = @($rows | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq $g }).Count }
    $pcts = @($rows | ForEach-Object { N $_.UtilizationPct } | Where-Object { $null -ne $_ })
    $o.AvgUtilizationPct = if ($pcts.Count) { Rnd (($pcts | Measure-Object -Average).Average) 1 } else { $null }
    $o.EfficiencyScore = Get-WeightedScore $rows
    $o.MonthlyCost = Rnd (($rows | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
    $o.IdleMonthlyCost = Rnd (($rows | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Idle / Unused / Orphaned' } | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
    $o.Recommendations = @($script:Recs | Where-Object { $_.TypeKey -eq $key }).Count
    $o.EstAnnualSavings = Rnd (($rows | ForEach-Object { [double]((N $_.EstAnnualSavings) ?? 0) } | Measure-Object -Sum).Sum) 0
    $o.Coverage = if ($script:GenericKeys.Contains($key)) { 'Auto-discovered metrics' } else { 'Dedicated analysis' }
    $o.TypeKey = $key
    [pscustomobject]$o
}
$TypeSummary = @($TypeSummary | Sort-Object Category, ResourceType)

$SubSummary = @($AllRows | Group-Object SubscriptionName | ForEach-Object {
        $g = $_.Group; $sn = $_.Name
        [pscustomobject][ordered]@{
            SubscriptionName = $sn; ResourcesAnalyzed = $g.Count
            IdleUnusedOrphaned = @($g | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Idle / Unused / Orphaned' }).Count
            OverProvisioned = @($g | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Over-provisioned' }).Count
            UnderProvisioned = @($g | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Under-provisioned' }).Count
            EfficiencyScore = Get-WeightedScore $g
            MonthlyCost = Rnd (($g | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
            Recommendations = @($script:Recs | Where-Object { $_.SubscriptionName -eq $sn }).Count
            EstAnnualSavings = Rnd (($g | ForEach-Object { [double]((N $_.EstAnnualSavings) ?? 0) } | Measure-Object -Sum).Sum) 0
        }
    } | Sort-Object EstAnnualSavings -Descending)

$RgSummary = @($AllRows | Group-Object { "$($_.SubscriptionName)|$($_.ResourceGroup)" } | ForEach-Object {
        $g = $_.Group
        [pscustomobject][ordered]@{
            SubscriptionName = $g[0].SubscriptionName; ResourceGroup = $g[0].ResourceGroup; ResourcesAnalyzed = $g.Count
            IdleUnusedOrphaned = @($g | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Idle / Unused / Orphaned' }).Count
            OverProvisioned = @($g | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Over-provisioned' }).Count
            EfficiencyScore = Get-WeightedScore $g
            MonthlyCost = Rnd (($g | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
            EstAnnualSavings = Rnd (($g | ForEach-Object { [double]((N $_.EstAnnualSavings) ?? 0) } | Measure-Object -Sum).Sum) 0
        }
    } | Sort-Object EstAnnualSavings, MonthlyCost -Descending)

$CatSavings = @('Delete', 'Rightsize', 'Configure') | ForEach-Object {
    $c = $_; $rs = @($script:Recs | Where-Object Category -eq $c)
    [pscustomobject][ordered]@{ Category = $c; Recommendations = $rs.Count; EstAnnualSavings = Rnd ([double]$script:DedupAll[$c] * 12) 0 }
}
$PerfCount = @($script:Recs | Where-Object Category -eq 'Performance Risk').Count
$StatusTotals = [ordered]@{}; foreach ($g in $StatusGroups.Keys) { $StatusTotals[$g] = @($AllRows | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq $g }).Count }
$AnalyzedMonthly = Rnd (($AllRows | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
$EfficiencyOverall = Get-WeightedScore $AllRows
$IdleMonthly = Rnd (($AllRows | Where-Object { (Get-StatusGroup $_.UtilizationStatus) -eq 'Idle / Unused / Orphaned' } | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0
$CategorySummary = @($TypeSummary | Group-Object Category | ForEach-Object {
        $cn = $_.Name
        $rs = @($AllRows | Where-Object { ($script:TypeCategory[$_.TypeKey] ?? 'Other Services') -eq $cn })
        [pscustomobject][ordered]@{
            Category = $_.Name; ResourceTypes = $_.Count; Resources = ($_.Group | Measure-Object Resources -Sum).Sum
            MonthlyCost = Rnd (($_.Group | Measure-Object MonthlyCost -Sum).Sum) 0; EfficiencyScore = Get-WeightedScore $rs
            EstAnnualSavings = Rnd (($_.Group | Measure-Object EstAnnualSavings -Sum).Sum) 0
        }
    } | Sort-Object MonthlyCost -Descending)
$RuleSummary = @($script:Recs | Group-Object RuleCode | ForEach-Object {
        $rc = $script:RuleCatalog[$_.Name]
        [pscustomobject][ordered]@{ RuleFamily = $_.Name; RuleName = $rc.Name; Recommendations = $_.Count; ResourceTypes = (@($_.Group | Select-Object -ExpandProperty ResourceType -Unique) -join ', ')
            EstAnnualSavings = Rnd ((Get-DedupSavings @($_.Group)).Total * 12) 0 }
    } | Sort-Object EstAnnualSavings -Descending)

# ---- Tags: coverage and cost allocation
$TagStats = @{}
foreach ($r in $AllRows) {
    $cost = Get-TotalCost $r
    foreach ($p in $r.PSObject.Properties) {
        if (-not $p.Name.StartsWith('Tag.') -or [string]::IsNullOrEmpty([string]$p.Value)) { continue }
        $k = $p.Name.Substring(4)
        if ($k -match '^(?i)(aks-managed-|hidden-|k8s-azure-|ms-resource-usage)') { continue }
        if (-not $TagStats.ContainsKey($k)) { $TagStats[$k] = @{ Resources = 0; Cost = 0.0; Values = @{} } }
        $t = $TagStats[$k]; $t.Resources++; $t.Cost += $cost
        $v = [string]$p.Value
        if (-not $t.Values.ContainsKey($v)) { $t.Values[$v] = @{ Resources = 0; Cost = 0.0; Savings = 0.0 } }
        $t.Values[$v].Resources++; $t.Values[$v].Cost += $cost; $t.Values[$v].Savings += [double]((N $r.EstAnnualSavings) ?? 0)
    }
}
$TagCoverage = @($TagStats.Keys | ForEach-Object {
        $t = $TagStats[$_]
        [pscustomobject][ordered]@{ TagKey = $_; ResourcesTagged = $t.Resources; TagCoveragePct = Rnd (100.0 * $t.Resources / [Math]::Max(1, $AllRows.Count)) 1
            DistinctValues = $t.Values.Count; MonthlyCostTagged = Rnd $t.Cost 0; CostCoveragePct = if ($AnalyzedMonthly -gt 0) { Rnd (100.0 * $t.Cost / $AnalyzedMonthly) 1 } else { $null } }
    } | Sort-Object ResourcesTagged -Descending)
$TopTagKeys = @($TagCoverage | Select-Object -First 8 | ForEach-Object TagKey)
$CostByTag = @(foreach ($k in $TopTagKeys) {
        $t = $TagStats[$k]
        foreach ($v in ($t.Values.Keys | Sort-Object { $t.Values[$_].Cost } -Descending)) {
            [pscustomobject][ordered]@{ TagKey = $k; TagValue = $v; Resources = $t.Values[$v].Resources; MonthlyCost = Rnd $t.Values[$v].Cost 0; EstAnnualSavings = Rnd $t.Values[$v].Savings 0 }
        }
        $untagged = @($AllRows | Where-Object { [string]::IsNullOrEmpty([string]($_.PSObject.Properties["Tag.$k"].Value)) })
        [pscustomobject][ordered]@{ TagKey = $k; TagValue = '(not tagged)'; Resources = $untagged.Count; MonthlyCost = Rnd (($untagged | ForEach-Object { Get-TotalCost $_ } | Measure-Object -Sum).Sum) 0; EstAnnualSavings = Rnd (($untagged | ForEach-Object { [double]((N $_.EstAnnualSavings) ?? 0) } | Measure-Object -Sum).Sum) 0 }
    })

# ---- Cost history (monthly) - validates the impact of previous optimization
$CostMonthly = Import-RawCsv 'Cost_MonthlyTrend.csv'
$MonthlyPivot = @(); $MonthlyTotals = @(); $Months = @()
if ($CostMonthly.Count) {
    $useUsdM = @($CostMonthly | Where-Object { $_.CostUSD }).Count -gt 0
    $mVal = { param($x) [double]((N $(if ($useUsdM) { $x.CostUSD } else { $x.Cost })) ?? 0) }
    $Months = @($CostMonthly | Select-Object -ExpandProperty Month -Unique | Sort-Object)
    $svcTot = @{}; foreach ($x in $CostMonthly) { $svcTot[$x.ServiceName] = [double]($svcTot[$x.ServiceName] ?? 0) + (& $mVal $x) }
    $topSvcM = @($svcTot.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object Key)
    $cell = @{}; foreach ($x in $CostMonthly) { $s = if ($x.ServiceName -in $topSvcM) { $x.ServiceName } else { 'Other services' }; $k = "$s|$($x.Month)"; $cell[$k] = [double]($cell[$k] ?? 0) + (& $mVal $x) }
    $MonthlyPivot = @(foreach ($s in ($topSvcM + @('Other services'))) {
            $o = [ordered]@{ ServiceName = $s }
            foreach ($m in $Months) { $o[$m] = Rnd ([double]($cell["$s|$m"] ?? 0)) 0 }
            [pscustomobject]$o
        })
    $partial = @($CostMonthly | Where-Object { B $_.PartialMonth } | Select-Object -ExpandProperty Month -Unique)
    $prev = $null
    $MonthlyTotals = @(foreach ($m in $Months) {
            $t = [double](($CostMonthly | Where-Object Month -eq $m | ForEach-Object { & $mVal $_ } | Measure-Object -Sum).Sum)
            [pscustomobject][ordered]@{ Month = $m; MonthlyCost = Rnd $t 0; ChangeVsPreviousPct = if ($prev -and $prev -gt 0) { Rnd (100 * ($t - $prev) / $prev) 1 } else { $null }; PartialMonth = ($m -in $partial) }
            $prev = $t
        })
}
$VmCommitCandidates = @($RecsSorted | Where-Object { $_.RuleCode -eq 'COMMIT-RI' })
$quickWins = @($RecsSorted | Where-Object { $_.Category -in @('Delete', 'Configure') -and $_.Effort -eq 'Low' -and $_.Confidence -ne 'Low' -and $_.EstAnnualSavings -gt 0 })

$fullMonths = @($MonthlyTotals | Where-Object { -not $_.PartialMonth })
$trendTxt = if ($fullMonths.Count -ge 2) { "$($fullMonths[0].Month): $(Money $fullMonths[0].MonthlyCost) -> $($fullMonths[-1].Month): $(Money $fullMonths[-1].MonthlyCost) ($(Fmt (100 * ($fullMonths[-1].MonthlyCost - $fullMonths[0].MonthlyCost) / [Math]::Max(1.0, [double]$fullMonths[0].MonthlyCost)) 1)%)" } else { 'n/a' }

$ExecKpis = @(
    [pscustomobject]@{ Metric = 'Customer'; Value = $CustomerName }
    [pscustomobject]@{ Metric = 'Assessment window'; Value = "$($Meta.WindowStartLocal) to $($Meta.WindowEndLocal) ($LookbackDays days, $($Meta.TimeZoneId))" }
    [pscustomobject]@{ Metric = 'Business hours / weekend (usage patterns)'; Value = "$($Meta.BusinessHoursStart):00-$($Meta.BusinessHoursEnd):00; weekend $(@($Meta.WeekendDays) -join ', ')" }
    [pscustomobject]@{ Metric = 'Subscriptions in scope'; Value = $Meta.SubscriptionCount }
    [pscustomobject]@{ Metric = 'Resources inventoried'; Value = $Meta.InventoryCount }
    [pscustomobject]@{ Metric = 'Resources analyzed (provisioned vs. utilized)'; Value = $AllRows.Count }
    [pscustomobject]@{ Metric = 'Resource types analyzed (dedicated + auto-discovered)'; Value = "$($TypeSummary.Count) ($(@($TypeSummary | Where-Object Coverage -eq 'Auto-discovered metrics').Count) auto-discovered)" }
    [pscustomobject]@{ Metric = 'Azure Monitor metric requests'; Value = "$($Meta.MetricCallsTotal) ($(if ($Meta.MetricsBatchApi) { 'Metrics Batch API' } else { 'per-resource' }))" }
    [pscustomobject]@{ Metric = "Monthly run-rate cost - all resources ($($Meta.CostType), $script:Currency)"; Value = $(if ($null -ne $TotalMonthlyCost) { Money $TotalMonthlyCost } else { "n/a ($($Meta.DataSources.CostData))" }) }
    [pscustomobject]@{ Metric = 'Monthly run-rate cost - analyzed resources'; Value = $(if ($script:HasCost) { Money $AnalyzedMonthly } else { 'n/a' }) }
    [pscustomobject]@{ Metric = 'Utilization efficiency score (cost-weighted, 0-100)'; Value = $(if ($null -ne $EfficiencyOverall) { "$EfficiencyOverall / 100" } else { 'n/a' }) }
    [pscustomobject]@{ Metric = 'Paid but idle / unused / orphaned (monthly cost)'; Value = "$($StatusTotals['Idle / Unused / Orphaned']) resources - $(Money $IdleMonthly)/month" }
    [pscustomobject]@{ Metric = 'Estimated annual savings (de-duplicated)'; Value = Money $AnnualSavings }
    [pscustomobject]@{ Metric = 'Quick wins (low effort)'; Value = "$($quickWins.Count) actions - $(Money ((Get-DedupSavings $quickWins).Total * 12))/year" }
    [pscustomobject]@{ Metric = 'Recommendations (Delete / Rightsize / Configure / Performance)'; Value = "$($CatSavings[0].Recommendations) / $($CatSavings[1].Recommendations) / $($CatSavings[2].Recommendations) / $PerfCount" }
    [pscustomobject]@{ Metric = 'Over-provisioned resources'; Value = $StatusTotals['Over-provisioned'] }
    [pscustomobject]@{ Metric = 'Under-provisioned (performance risk)'; Value = $StatusTotals['Under-provisioned'] }
    [pscustomobject]@{ Metric = 'Monthly cost trend (full months)'; Value = $trendTxt }
    [pscustomobject]@{ Metric = 'Log Analytics workspaces / with commitment tier'; Value = "$($script:LaAnalysis.Count) / $(@($script:LaAnalysis | Where-Object { $_.LAPricingTier -ne 'Pay-as-you-go' -or $_.SentinelCommitmentGBPerDay }).Count)" }
    [pscustomobject]@{ Metric = 'Data sources'; Value = (($Meta.DataSources.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }) -join ' | ') }
)

#endregion

#region ============================== EXCEL: MASTER WORKBOOK ==============================

$safeCustomer = ($CustomerName -replace '[\\/:*?"<>|]', '-').Trim()
$XlsxPath = Join-Path $OutputPath "$safeCustomer - Utilization & Cost Optimization Report.xlsx"
Write-Log "Building master workbook: $XlsxPath" Section
if (Test-Path $XlsxPath) { Remove-Item $XlsxPath -Force }
$pkg = Open-ExcelPackage -Path $XlsxPath -Create
$recCols = @('Priority', 'RuleId', 'RuleName', 'Category', 'SubscriptionName', 'ResourceGroup', 'ResourceName', 'ResourceType', 'Environment', 'EstAnnualSavings', 'EstMonthlySavings', 'CurrentMonthlyCost',
    'RecommendedAction', 'TargetConfiguration', 'Evidence', 'Criteria', 'CurrentConfiguration', 'ObservedUtilization', 'Confidence', 'Effort', 'Rationale', 'CustomerComment', 'Comment', 'Location', 'ResourceId')

# 0. Executive overview - what the utility is, what is collected, coverage of this run and how to read the workbook
$ov = Get-SheetName 'Executive Overview'
$autoTypes = @($TypeSummary | Where-Object Coverage -eq 'Auto-discovered metrics').Count
$qwAnnual = (Get-DedupSavings $quickWins).Total * 12
$About = @(
    [pscustomobject][ordered]@{ Topic = 'What this is'; Details = 'Azure Utilization & Cost Optimizer - a read-only assessment that measures what is provisioned versus what is actually used for every billed Azure resource (IaaS, PaaS, databases, consumption services, Log Analytics and Microsoft Sentinel) and turns the measured evidence into rule-based, prioritized optimization recommendations.' }
    [pscustomobject][ordered]@{ Topic = 'Questions answered'; Details = '1) Is every paid resource consumed at all (idle, unused, orphaned, stopped but still billed)? 2) How much of the current SKU / capacity is used (utilization % of the SKU, efficiency score)? 3) What should change - right-size, delete, reconfigure, schedule or commit - and what does it save? 4) Do commitments (Log Analytics / Sentinel tiers, reservations, savings plans) match real daily consumption?' }
    [pscustomobject][ordered]@{ Topic = 'Data collected'; Details = "Azure Resource Graph: every resource with all ARM properties and tags. Azure Monitor platform metrics at 1-hour grain (average, P95, P99, maximum, totals) via the Metrics Batch API, with daily series and hour-of-week usage patterns. Log Analytics: daily billable ingestion per workspace and per table. Cost Management ($($Meta.CostType)): monthly run-rate per resource, daily cost and monthly history. Azure Advisor cost and performance recommendations, orphaned-resource detection, and public retail / reservation / savings plan prices." }
    [pscustomobject][ordered]@{ Topic = 'Resource coverage'; Details = "$($TypeSummary.Count) resource types analyzed in this run ($autoTypes with auto-discovered metrics) across compute, containers, App Service, databases, storage, networking, monitoring & security, integration, analytics and AI. Billed types without a dedicated analyzer are covered automatically through their platform metrics and cost; any other billed item is listed on 'Other Paid Resources' so nothing that is paid for is missed." }
    [pscustomobject][ordered]@{ Topic = 'Assessment scope'; Details = "$CustomerName - $($Meta.SubscriptionCount) subscription(s), $($Meta.InventoryCount) resources inventoried, $($AllRows.Count) analyzed. Window $($Meta.WindowStartLocal) to $($Meta.WindowEndLocal) ($LookbackDays days, $($Meta.TimeZoneId)); business hours $($Meta.BusinessHoursStart):00-$($Meta.BusinessHoursEnd):00, weekend $(@($Meta.WeekendDays) -join '/')." }
    [pscustomobject][ordered]@{ Topic = 'Headline results'; Details = "Monthly run-rate $(if ($script:HasCost) { Money $AnalyzedMonthly } else { 'n/a' }) (analyzed resources); estimated savings $(Money $AnnualSavings) per year (de-duplicated per resource); utilization efficiency score $(if ($null -ne $EfficiencyOverall) { "$EfficiencyOverall/100" } else { 'n/a' }) (cost-weighted); $($script:Recs.Count) recommendations, of which $($quickWins.Count) low-effort quick wins worth $(Money $qwAnnual) per year." }
    [pscustomobject][ordered]@{ Topic = 'How recommendations are built'; Details = "Each suggestion references a rule (Rule ID = <service prefix>-<rule family>, e.g. VM-RS-SKU) with explicit criteria rendered with the active thresholds (see 'Recommendation Rules'), the measured evidence, confidence, effort and estimated savings. Savings are indicative - validate owners, licensing, reservations and compliance before changing resources." }
    [pscustomobject][ordered]@{ Topic = 'Safety'; Details = 'Read-only: no changes are made to the environment. The analysis runs offline from the exported data and can be re-run with different thresholds without new Azure calls.' }
    [pscustomobject][ordered]@{ Topic = 'Solution developed by'; Details = 'Zahir Hussain Shah - Microsoft Qatar' }
)
$Coverage = @($TypeSummary | Group-Object Category | Sort-Object Name | ForEach-Object {
        [pscustomobject][ordered]@{ Category = $_.Name; ResourceTypes = (@($_.Group | ForEach-Object ResourceType) -join ', '); Types = $_.Count; Resources = ($_.Group | Measure-Object Resources -Sum).Sum
            MonthlyCost = Rnd (($_.Group | Measure-Object MonthlyCost -Sum).Sum) 0; EstAnnualSavings = Rnd (($_.Group | Measure-Object EstAnnualSavings -Sum).Sum) 0 }
    })
$Guide = @(
    [pscustomobject][ordered]@{ Sheet = 'Executive Dashboard'; WhatItShows = 'KPIs (run-rate, efficiency, waste, savings, quick wins, cost trend) and charts' }
    [pscustomobject][ordered]@{ Sheet = 'Suggested Recommendations'; WhatItShows = 'Every suggestion with Rule ID, criteria, evidence, current vs. target configuration, savings, confidence and effort' }
    [pscustomobject][ordered]@{ Sheet = 'Action Tracker'; WhatItShows = 'One row per action with owner, target date, status and customer decision drop-downs for the review' }
    [pscustomobject][ordered]@{ Sheet = 'Quick Wins / Rightsize / Delete / Configure / Performance Risk'; WhatItShows = 'Recommendations by action type' }
    [pscustomobject][ordered]@{ Sheet = 'Efficiency Scorecard / By Category / By Subscription / By Resource Group / Savings by Rule'; WhatItShows = 'Utilization efficiency, status distribution, cost and savings by dimension' }
    [pscustomobject][ordered]@{ Sheet = 'Commitments / LA Daily Ingestion / LA Daily by Table'; WhatItShows = 'Log Analytics & Sentinel tiers vs. actual daily ingestion, reservations, savings plans and steady-VM commitment candidates' }
    [pscustomobject][ordered]@{ Sheet = 'Cost Trend / Daily Cost / Cost by Tag / Tag Coverage'; WhatItShows = 'Monthly cost history (impact of previous optimization), daily cost, cost allocation and tagging quality' }
    [pscustomobject][ordered]@{ Sheet = 'Usage Patterns / Daily Utilization Trend'; WhatItShows = 'Hour-of-week heat maps and daily utilization trends per service' }
    [pscustomobject][ordered]@{ Sheet = 'One sheet per resource type'; WhatItShows = 'Every resource with utilization statistics vs. its SKU, efficiency score, suggested recommendation, all tags (Tag.*) and all properties (Prop.*)' }
    [pscustomobject][ordered]@{ Sheet = 'Orphaned Resources / Other Paid Resources / Advisor / Defender Plans / Inventory'; WhatItShows = 'Completeness views - everything that is billed or deployed is visible' }
    [pscustomobject][ordered]@{ Sheet = 'Recommendation Rules / Data Dictionary / Methodology / Scope'; WhatItShows = 'Rule catalog with criteria, column definitions incl. the Azure Monitor metric behind each column, methodology and subscriptions in scope' }
    [pscustomobject][ordered]@{ Sheet = 'Companion files'; WhatItShows = 'ResourceTypes\<Type>.xlsx (one workbook per resource type with daily utilization), the interactive HTML dashboard and Recommendations.csv' }
)
$pkg = Export-Sheet -Pkg $pkg -Data $About -Sheet $ov -Title "Azure Utilization & Cost Optimizer - $CustomerName" -StartRow 1
$covRow = $About.Count + 5
$pkg = Export-Sheet -Pkg $pkg -Data $Coverage -Sheet $ov -Title 'Resource coverage in this assessment' -StartRow $covRow
$guideRow = $covRow + $Coverage.Count + 4
$pkg = Export-Sheet -Pkg $pkg -Data $Guide -Sheet $ov -Title 'How to use this workbook' -StartRow $guideRow
$wsOv = $pkg.Workbook.Worksheets[$ov]
$wsOv.Column(1).Width = 34; $wsOv.Column(2).Width = 120; foreach ($ci in 3..6) { $wsOv.Column($ci).Width = 16 }
$wsOv.Column(2).Style.WrapText = $true; $wsOv.Cells["A1:F$($guideRow + $Guide.Count + 2)"].Style.VerticalAlignment = 'Top'
$wsOv.Cells['A1'].Style.Font.Size = 16; $wsOv.Cells['A1'].Style.Font.Color.SetColor((Col '#0F4C81'))

# 1. Executive dashboard (KPIs + native Excel charts)
$dash = Get-SheetName 'Executive Dashboard'
$pkg = Export-Sheet -Pkg $pkg -Data $ExecKpis -Sheet $dash -Title "$CustomerName - Azure Utilization & Cost Optimization" -StartRow 1
$pkg = Export-Sheet -Pkg $pkg -Data @($CatSavings) -Sheet $dash -Title 'Savings by action' -StartRow 1 -StartColumn 4
$statusTable = @($StatusTotals.Keys | ForEach-Object { [pscustomobject]@{ Status = $_; Resources = $StatusTotals[$_] } })
$pkg = Export-Sheet -Pkg $pkg -Data $statusTable -Sheet $dash -Title 'Utilization status' -StartRow 7 -StartColumn 4
$catTable = @($CategorySummary | Select-Object Category, Resources, MonthlyCost, EfficiencyScore, EstAnnualSavings)
$pkg = Export-Sheet -Pkg $pkg -Data $catTable -Sheet $dash -Title 'By service category' -StartRow 16 -StartColumn 4
$topTypes = @($TypeSummary | Sort-Object EstAnnualSavings -Descending | Select-Object -First 10 ResourceType, EstAnnualSavings)
$tRow = 18 + $catTable.Count + 2
$pkg = Export-Sheet -Pkg $pkg -Data $topTypes -Sheet $dash -Title 'Top resource types by savings' -StartRow $tRow -StartColumn 4
$ws = $pkg.Workbook.Worksheets[$dash]; $ws.Column(1).Width = 52; $ws.Column(2).Width = 75
$q = "'$dash'"
try {
    Add-ExcelChart -Worksheet $ws -ChartType Pie -XRange "$q!D3:D5" -YRange "$q!F3:F5" -Title 'Estimated annual savings by action' -Row 0 -Column 9 -Width 480 -Height 300 -ShowPercent | Out-Null
    Add-ExcelChart -Worksheet $ws -ChartType BarClustered -XRange "$q!D9:D14" -YRange "$q!E9:E14" -Title 'Resources by utilization status' -Row 16 -Column 9 -Width 480 -Height 300 -NoLegend | Out-Null
    if ($topTypes.Count) { Add-ExcelChart -Worksheet $ws -ChartType BarClustered -XRange "$q!D$($tRow + 2):D$($tRow + 1 + $topTypes.Count)" -YRange "$q!E$($tRow + 2):E$($tRow + 1 + $topTypes.Count)" -Title 'Top resource types by annual savings' -Row 32 -Column 9 -Width 480 -Height 320 -NoLegend | Out-Null }
    if ($catTable.Count) { Add-ExcelChart -Worksheet $ws -ChartType ColumnClustered -XRange "$q!D18:D$(17 + $catTable.Count)" -YRange "$q!F18:F$(17 + $catTable.Count)" -Title 'Monthly cost by service category' -Row 0 -Column 17 -Width 520 -Height 300 -NoLegend | Out-Null }
}
catch { Write-Log "Chart creation skipped: $($_.Exception.Message)" Warning }

# 2. Suggested recommendations + action tracker + category views
$pkg = Export-Sheet -Pkg $pkg -Data @($RecsSorted | Select-Object $recCols) -Sheet (Get-SheetName 'Suggested Recommendations') -FreezePane 2, 4
$i = 0
$tracker = @($RecsSorted | Where-Object { $_.Category -ne 'Performance Risk' } | ForEach-Object {
        $i++
        [pscustomobject][ordered]@{ ActionId = ('A{0:000}' -f $i); Priority = $_.Priority; RuleId = $_.RuleId; ResourceName = $_.ResourceName; ResourceType = $_.ResourceType; SubscriptionName = $_.SubscriptionName
            RecommendedAction = $_.RecommendedAction; EstAnnualSavings = $_.EstAnnualSavings; Effort = $_.Effort; Confidence = $_.Confidence
            Owner = ''; TargetDate = ''; Status = 'Not started'; CustomerDecision = ''; Notes = '' }
    })
$trk = Get-SheetName 'Action Tracker'
$pkg = Export-Sheet -Pkg $pkg -Data $tracker -Sheet $trk -FreezePane 2, 5
if ($tracker.Count) {
    $wsT = $pkg.Workbook.Worksheets[$trk]; $n = $tracker.Count + 1
    try {
        Add-ExcelDataValidationRule -Worksheet $wsT -Range "M2:M$n" -ValidationType List -ValueSet @('Not started', 'In progress', 'Done', 'Blocked') -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Status' -ErrorBody 'Choose a status from the list' | Out-Null
        Add-ExcelDataValidationRule -Worksheet $wsT -Range "N2:N$n" -ValidationType List -ValueSet @('Accepted', 'Rejected', 'Deferred', 'Needs review') -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Decision' -ErrorBody 'Choose a decision from the list' | Out-Null
    }
    catch { Write-Log "Validation lists skipped: $($_.Exception.Message)" Warning }
}
$pkg = Export-Sheet -Pkg $pkg -Data @($quickWins | Select-Object $recCols) -Sheet (Get-SheetName 'Quick Wins') -FreezePane 2, 4
foreach ($c in @('Rightsize', 'Delete', 'Configure', 'Performance Risk')) {
    $pkg = Export-Sheet -Pkg $pkg -Data @($RecsSorted | Where-Object Category -eq $c | Select-Object $recCols) -Sheet (Get-SheetName $c) -FreezePane 2, 4
}

# 3. Efficiency scorecard & utilization summaries
$pkg = Export-Sheet -Pkg $pkg -Data @($TypeSummary | Sort-Object MonthlyCost -Descending | Select-Object -ExcludeProperty TypeKey) -Sheet (Get-SheetName 'Efficiency Scorecard')
$pkg = Export-Sheet -Pkg $pkg -Data $CategorySummary -Sheet (Get-SheetName 'By Category')
$pkg = Export-Sheet -Pkg $pkg -Data $SubSummary -Sheet (Get-SheetName 'By Subscription')
$pkg = Export-Sheet -Pkg $pkg -Data $RgSummary -Sheet (Get-SheetName 'By Resource Group')
$pkg = Export-Sheet -Pkg $pkg -Data $RuleSummary -Sheet (Get-SheetName 'Savings by Rule')

# 4. Commitments
$cs = Get-SheetName 'Commitments'; $row0 = 1
if ($script:LaAnalysis.Count) { $pkg = Export-Sheet -Pkg $pkg -Data @($script:LaAnalysis) -Sheet $cs -Title 'Log Analytics & Microsoft Sentinel - pricing tier vs. actual daily ingestion' -StartRow $row0; $row0 += $script:LaAnalysis.Count + 4 }
if ($VmCommitCandidates.Count) { $pkg = Export-Sheet -Pkg $pkg -Data @($VmCommitCandidates | Select-Object ResourceName, SubscriptionName, ResourceGroup, ObservedUtilization, RecommendedAction, TargetConfiguration, CurrentMonthlyCost, EstAnnualSavings, Evidence) -Sheet $cs -Title 'Steady 24x7 VMs - reservation / savings plan candidates' -StartRow $row0; $row0 += $VmCommitCandidates.Count + 4 }
if ($Reservations.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $Reservations -Sheet $cs -Title 'Reservations - utilization' -StartRow $row0; $row0 += $Reservations.Count + 4 }
if ($SavingsPlans.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $SavingsPlans -Sheet $cs -Title 'Savings plans - utilization' -StartRow $row0 }

# 5. Log Analytics daily ingestion (per workspace, per table)
$laPivot = @()
if ($LaDaily.Count) {
    $laPivot = foreach ($g in ($LaDaily | Group-Object { ([string]$_.ResourceId).ToLower() })) {
        $f = $g.Group[0]; $vals = @($g.Group | ForEach-Object { [double](N $_.BillableGB) })
        $o = [ordered]@{ SubscriptionName = $f.SubscriptionName; ResourceGroup = $f.ResourceGroup; Workspace = $f.Workspace; CommitmentGBPerDay = N $f.CommitmentGBPerDay; SentinelCommitmentGBPerDay = N $f.SentinelCommitmentGBPerDay
            AvgGB = Rnd (($vals | Measure-Object -Average).Average) 2; MaxGB = Rnd (($vals | Measure-Object -Maximum).Maximum) 2 }
        $byDate = @{}; foreach ($x in $g.Group) { $byDate[$x.Date] = Rnd (N $x.BillableGB) 2 }
        foreach ($d in $WindowDates) { $o[$d] = $byDate[$d] }
        [pscustomobject]$o
    }
    $laPivot = @($laPivot | Sort-Object AvgGB -Descending)
    $las = Get-SheetName 'LA Daily Ingestion'
    $pkg = Export-Sheet -Pkg $pkg -Data $laPivot -Sheet $las -Title 'Daily billable ingestion (GB) per workspace vs. commitment' -NoFormat
    Add-Heatmap $pkg.Workbook.Worksheets[$las] "H3:$(Get-ColLetter (7 + $WindowDates.Count))$(2 + $laPivot.Count)"
}
$LaByTable = Import-RawCsv 'LogAnalytics_DailyByTable.csv'
if ($LaByTable.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $LaByTable -Sheet (Get-SheetName 'LA Daily by Table') -NoFormat }

# 6. Cost history, daily cost, cost by tag
if ($MonthlyPivot.Count) {
    $ct = Get-SheetName 'Cost Trend'
    $pkg = Export-Sheet -Pkg $pkg -Data $MonthlyTotals -Sheet $ct -Title "Monthly cost ($($Meta.CostType)) - validates impact of previous optimization" -StartRow 1
    $pkg = Export-Sheet -Pkg $pkg -Data $MonthlyPivot -Sheet $ct -Title 'Monthly cost by service' -StartRow ($MonthlyTotals.Count + 5) -NoFormat
    try { Add-ExcelChart -Worksheet $pkg.Workbook.Worksheets[$ct] -ChartType ColumnClustered -XRange "'$ct'!A3:A$(2 + $MonthlyTotals.Count)" -YRange "'$ct'!B3:B$(2 + $MonthlyTotals.Count)" -Title 'Monthly cost trend' -Row 0 -Column 6 -Width 560 -Height 300 -NoLegend | Out-Null } catch { Write-Log "Trend chart skipped: $($_.Exception.Message)" Warning }
}
if ($CostDaily.Count) {
    $useUsdD = @($CostDaily | Where-Object { $_.CostUSD }).Count -gt 0
    $dailyPivot = @($CostDaily | Group-Object ServiceName | ForEach-Object {
            $o = [ordered]@{ ServiceName = $_.Name }; $byD = @{}
            foreach ($x in $_.Group) { $byD[$x.Date] = [double]($byD[$x.Date] ?? 0) + [double]((N $(if ($useUsdD) { $x.CostUSD } else { $x.Cost })) ?? 0) }
            $o.WindowTotal = Rnd (($byD.Values | Measure-Object -Sum).Sum) 0
            foreach ($d in $WindowDates) { $o[$d] = Rnd $byD[$d] 1 }
            [pscustomobject]$o
        } | Sort-Object WindowTotal -Descending)
    $dc = Get-SheetName 'Daily Cost'
    $pkg = Export-Sheet -Pkg $pkg -Data $dailyPivot -Sheet $dc -Title 'Daily cost by service' -NoFormat
    Add-Heatmap $pkg.Workbook.Worksheets[$dc] "C3:$(Get-ColLetter (2 + $WindowDates.Count))$(2 + $dailyPivot.Count)"
}
$pkg = Export-Sheet -Pkg $pkg -Data $CostByTag -Sheet (Get-SheetName 'Cost by Tag')
$pkg = Export-Sheet -Pkg $pkg -Data $TagCoverage -Sheet (Get-SheetName 'Tag Coverage')

# 7. Usage patterns & daily utilization trend
$script:HowMatrices = [ordered]@{}
$up = Get-SheetName 'Usage Patterns'; $pr = 1
foreach ($key in $TypeData.Keys) {
    $hm = Get-HowMatrix $key
    if (-not $hm) { continue }
    $script:HowMatrices[$key] = $hm
    $pkg = Export-Sheet -Pkg $pkg -Data $hm.Matrix -Sheet $up -Title "$($script:SheetNames[$key] ?? $key) - average $($hm.Metric) by local day/hour ($($hm.Resources) resources)" -StartRow $pr -NoFormat -TableStyle 'Light1'
    Add-Heatmap $pkg.Workbook.Worksheets[$up] "B$($pr + 2):Y$($pr + 8)"
    $pr += 11
}
$script:DailyTrend = [ordered]@{}
foreach ($key in $TypeData.Keys) {
    $rows = @(Rows $key); $prim = ($rows | Where-Object PrimaryMetric | Group-Object PrimaryMetric | Sort-Object Count -Descending | Select-Object -First 1).Name
    if (-not $prim) { continue }
    $field = Get-MetricField $rows $prim
    $daily = @(Get-DailyRows $key | Where-Object { $_.Metric -eq $prim })
    if ($daily.Count -eq 0) { continue }
    $byDate = @{}; foreach ($g in ($daily | Group-Object Date)) { $v = @($g.Group | ForEach-Object { N $_.$field } | Where-Object { $null -ne $_ }); if ($v.Count) { $byDate[$g.Name] = Rnd (($v | Measure-Object -Average).Average) 2 } }
    $script:DailyTrend["$($script:SheetNames[$key] ?? $key) ($prim)"] = @($WindowDates | ForEach-Object { $byDate[$_] })
}
if ($script:DailyTrend.Count) {
    $trendRows = foreach ($k in $script:DailyTrend.Keys) { $o = [ordered]@{ Series = $k }; $i = 0; foreach ($d in $WindowDates) { $o[$d] = $script:DailyTrend[$k][$i]; $i++ }; [pscustomobject]$o }
    $pkg = Export-Sheet -Pkg $pkg -Data @($trendRows) -Sheet (Get-SheetName 'Daily Utilization Trend') -Title 'Average of the primary utilization metric across resources, per day' -NoFormat
}

# 8. One sheet per resource type (every property + every tag + utilization + suggested recommendation)
$script:TypeSheetMap = @{}
foreach ($ts in ($TypeSummary | Sort-Object Category, ResourceType)) {
    $sheet = Get-SheetName $ts.ResourceType
    $script:TypeSheetMap[$ts.TypeKey] = $sheet
    $pkg = Export-Sheet -Pkg $pkg -Data (Get-OrderedRows (Rows $ts.TypeKey)) -Sheet $sheet -FreezePane 2, 4
}

# 9. Supporting sheets
if ($Orphans.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $Orphans -Sheet (Get-SheetName 'Orphaned Resources') }
if ($OtherPaid.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $OtherPaid -Sheet (Get-SheetName 'Other Paid Resources') }
if ($Advisor.Count) { $pkg = Export-Sheet -Pkg $pkg -Data @($Advisor | Select-Object -ExcludeProperty Details) -Sheet (Get-SheetName 'Advisor (Cost & Perf)') }
if ($Defender.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $Defender -Sheet (Get-SheetName 'Defender Plans') }
if ($Inventory.Count) { $pkg = Export-Sheet -Pkg $pkg -Data $Inventory -Sheet (Get-SheetName 'Inventory (All Resources)') -FreezePane 2, 4 }
# Rule applicability: resource types where each rule family is implemented (read from the analyzer code itself)
$RuleApplies = @{}
try {
    $astRoot = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null)
    foreach ($c in $astRoot.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Add-Rec' }, $true)) {
        $els = $c.CommandElements; $pos = [System.Collections.Generic.List[object]]::new(); $rules = @()
        for ($i = 1; $i -lt $els.Count; $i++) {
            $e = $els[$i]
            if ($e -is [System.Management.Automation.Language.CommandParameterAst]) {
                $arg = $e.Argument
                if (-not $arg -and ($i + 1) -lt $els.Count) { $i++; $arg = $els[$i] }
                if ($e.ParameterName -eq 'Rule' -and $arg) {
                    # Constant rule or an expression choosing between rules: collect every rule code it can produce
                    $rules = @($arg.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true) | ForEach-Object Value | Where-Object { $script:RuleCatalog.Contains([string]$_) } | Select-Object -Unique)
                }
                continue
            }
            $pos.Add($e)
        }
        if ($rules.Count -eq 0) { $rules = @('GENERAL') }
        $keys = if ($pos.Count -ge 2 -and $pos[1] -is [System.Management.Automation.Language.StringConstantExpressionAst]) { @($pos[1].Value) }
        else {
            $p = $c.Parent; $owner = ''
            while ($p) { if ($p -is [System.Management.Automation.Language.AssignmentStatementAst]) { $owner = $p.Left.Extent.Text; break }; $p = $p.Parent }
            if ($owner -match '^\$Analyzers\.Generic$') { @('(auto-discovered types)') }
            elseif ($owner -match '^\$Analyzers\.(\w+)$') { @($Matches[1]) }
            elseif ($owner -eq '$FlexAnalyzer') { @('MySQLFlexibleServers', 'PostgreSQLFlexibleServers') }
            else { @('(various)') }
        }
        foreach ($rule in $rules) {
            if (-not $RuleApplies.ContainsKey($rule)) { $RuleApplies[$rule] = [System.Collections.Generic.SortedSet[string]]::new() }
            foreach ($k in $keys) { [void]$RuleApplies[$rule].Add($(if ($script:SheetNames.Contains($k)) { [string]$script:SheetNames[$k] } else { $k })) }
        }
    }
}
catch { Write-Log "Rule applicability scan skipped: $($_.Exception.Message)" Warning }
$RulesSheet = @(foreach ($code in $script:RuleCatalog.Keys) {
        $rc = $script:RuleCatalog[$code]
        $applied = @($script:Recs | Where-Object RuleCode -eq $code)
        [pscustomobject][ordered]@{ RuleFamily = $code; RuleName = $rc.Name; Criteria = Format-RuleText $rc.Criteria; Rationale = $rc.Rationale
            AppliesTo = $(if ($RuleApplies.ContainsKey($code)) { @($RuleApplies[$code]) -join ', ' } else { '' })
            TriggeredRecommendations = $applied.Count; EstAnnualSavings = Rnd ((Get-DedupSavings $applied).Total * 12) 0
            ResourceTypesTriggered = (@($applied | Select-Object -ExpandProperty ResourceType -Unique) -join ', ')
            RuleIdsTriggered = (@($applied | Select-Object -ExpandProperty RuleId -Unique) -join ', ') }
    })
$pkg = Export-Sheet -Pkg $pkg -Data $RulesSheet -Sheet (Get-SheetName 'Recommendation Rules')
$suffixDesc = [ordered]@{ '_Avg' = 'Average of hourly values over the window'; '_P95' = '95th percentile of hourly values'; '_P99' = '99th percentile of hourly values'; '_Med' = 'Median of hourly values'
    '_Max' = 'Highest hourly maximum'; '_Min' = 'Lowest hourly minimum'; '_MinAvg' = 'Lowest hourly average'; '_P05' = '5th percentile of hourly values'; '_Total' = 'Sum over the window'; '_MaxT' = 'Largest hourly total (peak hour)' }
$MetricCatalog = Import-RawCsv 'Reference_MetricCatalog.csv'
$Dictionary = [System.Collections.Generic.List[object]]::new()
foreach ($d in @(
        @('UtilizationStatus', 'Analysis', 'Idle / Unused / Orphaned / Over-provisioned / Right-sized / In use / Under-provisioned / Stopped / Deallocated / No data'),
        @('UtilizationPct', 'Analysis', 'Primary utilization of the current SKU capacity (P95 of CPU, DTU, RU, quota, IOPS... depending on the service)'),
        @('EfficiencyScore', 'Analysis', 'Utilization efficiency 0-100: P95 utilization relative to the target utilization (idle/unused = 0)'),
        @('ProvisionedSummary', 'Analysis', 'What is provisioned (SKU, size, units, capacity)'), @('UtilizedSummary', 'Analysis', 'What is actually used (avg / P95 / peak)'),
        @('SuggestedRecommendation', 'Recommendation', 'All suggested actions for the resource, ordered Delete > Rightsize > Configure > Performance'),
        @('RuleIds', 'Recommendation', 'Rule identifiers that triggered (see Recommendation Rules sheet)'), @('RecommendationCriteria', 'Recommendation', 'The criteria of every triggered rule'),
        @('EstMonthlySavings', 'Recommendation', 'De-duplicated monthly savings for the resource'), @('MonthlyCost', 'Cost', 'Monthly run-rate cost from Cost Management ($($Meta.CostType))'),
        @('EffectiveDiscountPct', 'Cost', 'Discount vs. retail pay-as-you-go for the hours run (reservation / savings plan / EA benefits)'),
        @('AgeDays', 'Inventory', 'Days since the resource was created'), @('Tag.<key>', 'Tags', 'One column per tag key on the resource'),
        @('Prop.<path>', 'Properties', 'Every ARM property flattened to a column (arrays of objects as compact JSON)'), @('PortalLink', 'Inventory', 'Direct Azure portal link'),
        @('Pattern_*', 'Usage pattern', 'Business-hours vs off-hours / weekday vs weekend averages, idle hours %, peak hour and day for the primary metric'))) {
    $Dictionary.Add([pscustomobject][ordered]@{ ResourceType = '(all)'; Column = $d[0]; Area = $d[1]; Description = $d[2]; AzureMetric = '' })
}
foreach ($m in $MetricCatalog) {
    foreach ($o in (([string]$m.Outputs) -split ',')) {
        $sfx = "_$o"
        $Dictionary.Add([pscustomobject][ordered]@{ ResourceType = $m.ResourceType; Column = "$($m.Alias)$sfx"; Area = 'Metric'; Description = "$($suffixDesc[$sfx] ?? $o) of '$($m.AzureMetric)'$(if ([double]((N $m.Scale) ?? 1) -ne 1) { " (scaled by 1/$($m.Scale))" })$(if ($m.SubResource) { " on $($m.SubResource)" })$(if ($m.DimensionFilter) { " split by $($m.DimensionFilter)" })"; AzureMetric = $m.AzureMetric })
    }
}
$pkg = Export-Sheet -Pkg $pkg -Data $Dictionary.ToArray() -Sheet (Get-SheetName 'Data Dictionary')
$Methodology = @(
    [pscustomobject]@{ Topic = 'Approach'; Detail = 'Read-only collection. Azure Resource Graph (configuration, every property and tag), Azure Monitor platform metrics at 1-hour grain via the Metrics Batch API, Log Analytics Usage table, Cost Management (actual/amortized cost, monthly history), Azure Advisor (cost & performance) and public retail prices (pay-as-you-go, reservation, savings plan).' }
    [pscustomobject]@{ Topic = 'Coverage'; Detail = 'Dedicated analysis for 60+ resource types; every other resource type that has platform metrics or cost is covered automatically (metric definitions discovered at run time - utilization % and activity metrics).' }
    [pscustomobject]@{ Topic = 'Utilization statistics'; Detail = 'Avg = mean of hourly averages; P95/P99 = percentiles of hourly values (hourly peaks for peak-type metrics); Max = highest hourly maximum; VM memory peak = lowest hourly-average available memory. Counters with no data points are reported as 0.' }
    [pscustomobject]@{ Topic = 'Idle'; Detail = "Average utilization below $CpuIdleThreshold% and P95 below $(2 * $CpuIdleThreshold)% (memory below 50% where measured)." }
    [pscustomobject]@{ Topic = 'Over-provisioned'; Detail = "P95 utilization below $CpuLowThreshold% and peak memory below $MemLowThreshold% (where measured); capacity services use their own utilization metric (DTU, RU, capacity units, quota, IOPS, connections). Auto-discovered types: P95 below $GenericLow%." }
    [pscustomobject]@{ Topic = 'Under-provisioned'; Detail = "P95 above $CpuHighThreshold% (or memory P95 above $MemHighThreshold%), throttling (HTTP 429), session limits or deadlocks - flagged as performance risk." }
    [pscustomobject]@{ Topic = 'Right-size target'; Detail = "Smallest SKU in the same family that keeps P95 at or below $TargetUtilization% and memory at or below 80% of the observed peak; one step down at most when memory is not measured." }
    [pscustomobject]@{ Topic = 'Efficiency score'; Detail = "Per resource: min(100, P95 utilization / $TargetUtilization% x 100); idle/unused/orphaned = 0; stopped/not assessed excluded. Type, category and overall scores are cost-weighted." }
    [pscustomobject]@{ Topic = 'Savings'; Detail = 'Monthly run-rate cost x (1 - target price / current price) where cost data exists (keeps negotiated discounts); otherwise retail price deltas. Totals are de-duplicated per resource (Delete supersedes; capped at current cost).' }
    [pscustomobject]@{ Topic = 'Commitments'; Detail = "Log Analytics / Sentinel tiers are re-priced against the actual daily series. VMs running >= $SteadyHoursPct% of hours, right-sized and billed near pay-as-you-go are flagged as reservation / savings plan candidates; reservation and savings plan rates cover compute only, so Windows VMs keep the license at pay-as-you-go (VMs with Azure Hybrid Benefit are compared with the base compute rate)." }
    [pscustomobject]@{ Topic = 'Usage patterns'; Detail = "Business hours $($Meta.BusinessHoursStart):00-$($Meta.BusinessHoursEnd):00, weekend $(@($Meta.WeekendDays) -join '/'), time zone $($Meta.TimeZoneId)." }
    [pscustomobject]@{ Topic = 'Environment'; Detail = 'Environment/Env tags first, then subscription / resource group / resource naming (patterns configurable in CostOptimization.Config.json).' }
    [pscustomobject]@{ Topic = 'Caveats'; Detail = 'Estimates are indicative; validate owners, licensing, reservations and compliance before change. Guest-level metrics (disk space, process data) require Azure Monitor Agent and are not used.' }
)
$pkg = Export-Sheet -Pkg $pkg -Data $Methodology -Sheet (Get-SheetName 'Methodology')
$pkg.Workbook.Worksheets['Methodology'].Column(2).Width = 140
$pkg = Export-Sheet -Pkg $pkg -Data $Subscriptions -Sheet (Get-SheetName 'Scope')
$sheetCount = $pkg.Workbook.Worksheets.Count
Close-ExcelPackage $pkg
Write-Log "Master workbook saved ($sheetCount sheets)" Success

$RecsSorted | Select-Object $recCols | Export-Csv -Path (Join-Path $OutputPath 'Recommendations.csv') -NoTypeInformation -Encoding utf8BOM

#endregion

#region ============================== EXCEL: PER-TYPE WORKBOOKS ==============================

Write-Log 'Exporting one workbook per resource type (resources + recommendations + metric statistics + daily utilization + usage pattern)...' Section
foreach ($ts in $TypeSummary) {
    $key = $ts.TypeKey; $rows = @(Rows $key)
    $fileName = ((($ts.ResourceType -replace '\s*&\s*', ' and ') -replace '[\\/:*?"<>|]', '-') -replace '\s+', '_') + '.xlsx'
    $path = Join-Path $TypeOutPath $fileName
    if (Test-Path $path) { Remove-Item $path -Force }
    $script:UsedSheetNames.Clear()
    $tp = Open-ExcelPackage -Path $path -Create
    $tp = Export-Sheet -Pkg $tp -Data (Get-OrderedRows $rows) -Sheet (Get-SheetName 'Resources') -FreezePane 2, 4
    $typeRecs = @($RecsSorted | Where-Object TypeKey -eq $key | Select-Object $recCols)
    if ($typeRecs.Count) { $tp = Export-Sheet -Pkg $tp -Data $typeRecs -Sheet (Get-SheetName 'Recommendations') -FreezePane 2, 4 }
    # Fleet statistics for every numeric metric column
    $numCols = @($rows[0].PSObject.Properties.Name | Where-Object { $_ -match '_(Avg|P95|P99|Med|Max|Min|MinAvg|P05|Total|MaxT)$|Pct|Score$' -and -not $_.StartsWith('Prop.') })
    $stats = foreach ($c in $numCols) {
        $v = @($rows | ForEach-Object { N $_.$c } | Where-Object { $null -ne $_ } | Sort-Object)
        if ($v.Count -eq 0) { continue }
        [pscustomobject][ordered]@{ Column = $c; Resources = $v.Count; Min = Rnd $v[0] 2; Average = Rnd (($v | Measure-Object -Average).Average) 2; Median = Rnd $v[[int][Math]::Floor(($v.Count - 1) / 2)] 2
            P95 = Rnd $v[[Math]::Max(0, [int][Math]::Ceiling(0.95 * $v.Count) - 1)] 2; Max = Rnd $v[-1] 2 }
    }
    if ($stats) { $tp = Export-Sheet -Pkg $tp -Data @($stats) -Sheet (Get-SheetName 'Metric Statistics') }
    $daily = @(Get-DailyRows $key)
    if ($daily.Count) {
        $allM = @($daily | Group-Object Metric | Sort-Object Count -Descending | ForEach-Object Name)
        $prim = [string](($rows | Where-Object PrimaryMetric | Group-Object PrimaryMetric | Sort-Object Count -Descending | Select-Object -First 1).Name)
        $rowCols = $rows[0].PSObject.Properties.Name
        $ranked = @($allM | Sort-Object @{ e = { if ($_ -eq $prim) { 0 } elseif ("$($_)_P95" -in $rowCols) { 1 } else { 2 } } }, @{ e = { [array]::IndexOf($allM, $_) } })
        foreach ($m in ($ranked | Select-Object -First 3)) {
            $field = Get-MetricField $rows $m
            $pv = @(New-DailyPivot $daily $m $field)
            if ($pv.Count) {
                $sn = Get-SheetName "$m daily $(if ($field -eq 'Total') { 'total' } else { 'avg' })"
                $tp = Export-Sheet -Pkg $tp -Data $pv -Sheet $sn -NoFormat
                $firstDateCol = 4 + [int]($null -ne $pv[0].PSObject.Properties['Dimension']) + 2
                Add-Heatmap $tp.Workbook.Worksheets[$sn] "$(Get-ColLetter $firstDateCol)2:$(Get-ColLetter ($firstDateCol + $WindowDates.Count - 1))$($pv.Count + 1)"
            }
            if ($m -eq $prim -and $field -eq 'Average') {
                $pvx = @(New-DailyPivot $daily $m 'Maximum')
                if ($pvx.Count) { $tp = Export-Sheet -Pkg $tp -Data $pvx -Sheet (Get-SheetName "$m daily max") -NoFormat }
            }
        }
        if ($daily.Count -le 150000) { $tp = Export-Sheet -Pkg $tp -Data $daily -Sheet (Get-SheetName 'Daily (all metrics)') -NoFormat }
    }
    $hm = $script:HowMatrices[$key]
    if ($hm) {
        $hs = Get-SheetName 'Hour-of-Week Pattern'
        $tp = Export-Sheet -Pkg $tp -Data $hm.Matrix -Sheet $hs -Title "Average $($hm.Metric) by local day and hour ($($hm.Resources) resources)" -NoFormat -TableStyle 'Light1'
        Add-Heatmap $tp.Workbook.Worksheets[$hs] 'B3:Y9'
    }
    if ($key -eq 'LogAnalyticsWorkspaces' -and $laPivot) {
        $tp = Export-Sheet -Pkg $tp -Data $laPivot -Sheet (Get-SheetName 'Daily Ingestion vs Commitment') -NoFormat
        if ($LaByTable.Count) { $tp = Export-Sheet -Pkg $tp -Data $LaByTable -Sheet (Get-SheetName 'Daily Ingestion by Table') -NoFormat }
        if ($LaTopTables.Count) { $tp = Export-Sheet -Pkg $tp -Data $LaTopTables -Sheet (Get-SheetName 'Top Tables') }
    }
    Close-ExcelPackage $tp
}
Write-Log "Per-type workbooks: $($TypeSummary.Count) file(s) in $TypeOutPath" Success

#endregion

#region ============================== HTML DASHBOARD: DATA MODEL ==============================

Write-Log 'Building the interactive HTML dashboard...' Section
$HtmlPath = Join-Path (Resolve-Path -LiteralPath $OutputPath).ProviderPath "$safeCustomer - Utilization & Optimization Dashboard.html"
$HtmlPropCap = [int](Get-Cfg 'Report.HtmlMaxPropertiesPerType' 100)
$HtmlDailyBudget = [int](Get-Cfg 'Report.HtmlMaxDailyValues' 1500000)
$EChartsUrl = [string](Get-Cfg 'Report.EChartsUrl' 'https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js')
$propCap = [Math]::Min($HtmlPropCap, [Math]::Max(15, [int](200000 / [Math]::Max(1, $AllRows.Count))))

function Jn($v) {
    $d = N $v
    if ($null -eq $d -or [double]::IsNaN($d) -or [double]::IsInfinity($d)) { return $null }
    if ($d -eq 0) { return 0 }
    $a = [Math]::Abs($d)
    return [Math]::Round($d, $(if ($a -ge 100) { 1 } elseif ($a -ge 1) { 2 } elseif ($a -ge 0.01) { 4 } else { 6 }))
}
function Jv($v, [int]$Max = 400) {
    if ($null -eq $v) { return $null }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) { return Jn $v }
    $s = [string]$v
    if ($s.Length -eq 0) { return $null }
    if ($s.Length -le 15 -and $s -match '^-?\d+(\.\d+)?$' -and $s -notmatch '^-?0\d') { return Jn $s }
    if ($s.Length -gt $Max) { return $s.Substring(0, $Max) + ' ...' }
    return $s
}
$jsSubs = [System.Collections.Generic.List[string]]::new(); $jsSubIdx = @{}
function Get-SubIdx([string]$Name) { if (-not $jsSubIdx.ContainsKey($Name)) { $jsSubIdx[$Name] = $jsSubs.Count; $jsSubs.Add($Name) }; return $jsSubIdx[$Name] }
foreach ($s in $Subscriptions) { [void](Get-SubIdx $s.SubscriptionName) }

$dateIdx = @{}; for ($i = 0; $i -lt $WindowDates.Count; $i++) { $dateIdx[$WindowDates[$i]] = $i }
$nD = $WindowDates.Count
$MetricColRx = '^(?!Pattern_)[A-Za-z0-9]+_(Avg|P95|P99|Med|Max|Min|MinAvg|P05|Total|MaxT|PeakHour)$'
$CoreCols = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($c in @('SubscriptionName', 'SubscriptionId', 'ResourceGroup', 'ResourceName', 'ResourceType', 'ArmType', 'Location', 'Environment', 'TypeKey', 'MonthlyCost',
        'ClusterMonthlyCost', 'ExcludeFromTotals', 'PortalLink', 'UtilizationStatus', 'ProvisionedSummary', 'UtilizedSummary', 'UtilizationPct', 'AnalysisNote',
        'SuggestedRecommendation', 'RecommendationCategory', 'RuleIds', 'RecommendationCriteria', 'RecommendationPriority', 'RecommendationConfidence', 'EstMonthlySavings',
        'EstAnnualSavings', 'EfficiencyScore', 'Kind', 'CreatedTime', 'AgeDays', 'PrimaryMetric', 'MetricsStatus', 'MetricsErrors', 'MetricCalls', 'Tags', 'ResourceId',
        'RetailPaygMonthly', 'EffectiveDiscountPct', 'ResourceGroupTags')) { [void]$CoreCols.Add($c) }
$SkuCols = @('VmSize', 'ServiceObjective', 'ServerSku', 'SkuName', 'Sku', 'NodeVmSize', 'PricingTier', 'SkuTier', 'Tier', 'DiskSku')
$PatternCols = @('Pattern_ActiveHoursPct', 'Pattern_IdleHoursPct', 'Pattern_BizHoursAvg', 'Pattern_OffHoursAvg', 'Pattern_WeekdayAvg', 'Pattern_WeekendAvg', 'Pattern_PeakHour')

$MetricLabel = @{}; $MetricUnit = @{}
foreach ($m in $MetricCatalog) { if ($m.Alias) { $MetricLabel["$($m.TypeKey)|$($m.Alias)"] = [string]$m.AzureMetric } }
foreach ($g in (Import-RawCsv 'Reference_GenericMetrics.csv')) {
    if (-not $g.Alias) { continue }
    $gk = 'Gen_' + (($g.ArmType -replace '^microsoft\.', '') -replace '[^A-Za-z0-9]+', '_')
    $MetricLabel["$gk|$($g.Alias)"] = [string]$(if ($g.DisplayName) { $g.DisplayName } else { $g.Metric })
    if ($g.Unit) { $MetricUnit["$gk|$($g.Alias)"] = [string]$g.Unit }
}

function Get-HowAll([string]$Key) {
    $p = Join-Path $RawPath 'UsagePatterns' "$($Key)_HourOfWeek.csv"
    if (-not (Test-Path $p)) { return $null }
    $o = [ordered]@{}
    foreach ($g in (Import-Csv -Path $p -Encoding utf8 | Group-Object Metric)) {
        $z = [object[]]::new(168)
        foreach ($x in $g.Group) {
            $di = N $x.DayIndex; $h = N $x.Hour
            if ($null -ne $di -and $null -ne $h -and $di -ge 0 -and $di -le 6 -and $h -ge 0 -and $h -le 23) { $z[[int]$di * 24 + [int]$h] = Jn $x.AvgValue }
        }
        $o[$g.Name] = [ordered]@{ r = [int]((N $g.Group[0].Resources) ?? 0); z = $z }
    }
    return $o
}

$jsTypes = [System.Collections.Generic.List[object]]::new(); $typeIdx = @{}
$jsRes = [System.Collections.Generic.List[object]]::new()
$resIdxByKey = @{}; $resIdxBySRN = @{}
$dailyUsed = 0
foreach ($ts in $TypeSummary) {
    $key = $ts.TypeKey; $rows = @(Rows $key)
    if ($rows.Count -eq 0) { continue }
    $typeIdx[$key] = $jsTypes.Count
    $names = [System.Collections.Generic.List[string]]::new(); $colSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $rows) { foreach ($n in $r.PSObject.Properties.Name) { if ($colSet.Add($n)) { $names.Add($n) } } }
    $mc = @($names | Where-Object { $_ -match $MetricColRx })
    $fc = @($names | Where-Object { -not $CoreCols.Contains($_) -and $_ -notmatch $MetricColRx -and -not $_.StartsWith('Pattern_') -and -not $_.StartsWith('Tag.') -and -not $_.StartsWith('Prop.') -and -not $_.StartsWith('Plan.') })
    $pop = @{}
    foreach ($c in @($names | Where-Object { $_.StartsWith('Prop.') -or $_.StartsWith('Plan.') })) {
        $cnt = 0; foreach ($r in $rows) { if (-not [string]::IsNullOrEmpty([string]$r.$c)) { $cnt++ } }
        if ($cnt) { $pop[$c] = $cnt }
    }
    $pc = @($names | Where-Object { $pop.ContainsKey($_) } | Select-Object -First $propCap)
    $pmCommon = [string](($rows | Where-Object PrimaryMetric | Group-Object PrimaryMetric | Sort-Object Count -Descending | Select-Object -First 1).Name)
    $ml = [ordered]@{}; $mu = [ordered]@{}
    foreach ($c in $mc) {
        $alias = $c -replace '_(Avg|P95|P99|Med|Max|Min|MinAvg|P05|Total|MaxT|PeakHour)$', ''
        if (-not $ml.Contains($alias)) { $ml[$alias] = $MetricLabel["$key|$alias"] ?? $alias; if ($MetricUnit["$key|$alias"]) { $mu[$alias] = $MetricUnit["$key|$alias"] } }
    }
    $jsTypes.Add([ordered]@{ k = $key; n = [string]$ts.ResourceType; c = [string]$ts.Category; a = $script:GenericKeys.Contains($key); pfx = [string]($script:TypePrefix[$key] ?? 'GEN')
            pm = $pmCommon; mc = $mc; fc = $fc; pc = $pc; ml = $ml; mu = $mu; hw = (Get-HowAll $key) })

    # Daily series for this type: "<resourceId>|<metric>|<dimension>" -> arrays aligned to the window dates
    $dmap = @{}; $byRid = @{}
    foreach ($x in (Get-DailyRows $key)) {
        $di = $dateIdx[[string]$x.Date]; if ($null -eq $di) { continue }
        $rid = ([string]$x.ResourceId).ToLower(); if (-not $rid) { continue }
        $mk = "$rid|$($x.Metric)|$($x.Dimension)"
        $e = $dmap[$mk]
        if (-not $e) {
            $e = @{ a = [object[]]::new($nD); x = [object[]]::new($nD); t = [object[]]::new($nD); m = [string]$x.Metric; dim = [string]$x.Dimension }
            $dmap[$mk] = $e
            if (-not $byRid.ContainsKey($rid)) { $byRid[$rid] = [System.Collections.Generic.List[object]]::new() }
            $byRid[$rid].Add($e)
        }
        $e.a[$di] = Jn $x.Average; $e.x[$di] = Jn $x.Maximum; $e.t[$di] = Jn $x.Total
    }

    foreach ($row in $rows) {
        $child = ''
        foreach ($sub in @('NodePool', 'ShareName', 'Deployment', 'ComputeName')) { if ($row.PSObject.Properties[$sub] -and $row.$sub) { $child = [string]$row.$sub; break } }
        $nm = if ($child) { "$($row.ResourceName) / $child" } else { [string]$row.ResourceName }
        $ri = $jsRes.Count
        $resIdxByKey["$key|$($row.ResourceId)|$nm".ToLower()] = $ri
        $srn = "$($row.SubscriptionName)|$($row.ResourceGroup)|$($row.ResourceName)".ToLower()
        if (-not $resIdxBySRN.ContainsKey($srn)) { $resIdxBySRN[$srn] = $ri }
        $sku = ''; foreach ($c in $SkuCols) { if ($colSet.Contains($c) -and $row.$c) { $sku = [string]$row.$c; break } }
        $o = [ordered]@{
            t = $typeIdx[$key]; s = Get-SubIdx ([string]$row.SubscriptionName); g = [string]$row.ResourceGroup; n = $nm; l = [string]$row.Location; e = [string]$row.Environment
            st = [string]$row.UtilizationStatus; u = Jn $row.UtilizationPct; es = Jn $row.EfficiencyScore; c = Jn $row.MonthlyCost; sv = Jn $row.EstAnnualSavings
            pv = [string]$row.ProvisionedSummary; us = [string]$row.UtilizedSummary; rc = [string]$row.RecommendationCategory; rp = [string]$row.RecommendationPriority
            sku = $sku; k = [string]$row.Kind; ag = Jn $row.AgeDays; cr = $(if ($row.CreatedTime) { ([string]$row.CreatedTime).Substring(0, [Math]::Min(10, ([string]$row.CreatedTime).Length)) } else { '' })
            pm = [string]$row.PrimaryMetric; ms = [string]$row.MetricsStatus; id = [string]$row.ResourceId
        }
        if ($row.PSObject.Properties['ClusterMonthlyCost'] -and $null -ne (N $row.ClusterMonthlyCost)) { $o.cc = Jn $row.ClusterMonthlyCost }
        if ($row.PSObject.Properties['ExcludeFromTotals'] -and $row.ExcludeFromTotals) { $o.x = 1 }
        if ($row.PSObject.Properties['AnalysisNote'] -and $row.AnalysisNote) { $o.an = [string]$row.AnalysisNote }
        if ($row.PSObject.Properties['EffectiveDiscountPct'] -and $null -ne (N $row.EffectiveDiscountPct)) { $o.ed = Jn $row.EffectiveDiscountPct; $o.rpm = Jn $row.RetailPaygMonthly }
        if ($colSet.Contains('Pattern_ActiveHoursPct') -and $null -ne (N $row.Pattern_ActiveHoursPct)) {
            $pat = @($PatternCols | Select-Object -First 6 | ForEach-Object { Jn $row.$_ })
            $hasSignal = ([double]($pat[2] ?? 0) -gt 0) -or ([double]($pat[3] ?? 0) -gt 0) -or ([double]($pat[4] ?? 0) -gt 0) -or ([double]($pat[5] ?? 0) -gt 0)
            $peakH = if ($hasSignal -and ([string]$row.Pattern_PeakHour) -match '^(\d{1,2})') { [int]$Matches[1] } else { $null }
            $o.pat = @($pat + @($peakH, $(if ($hasSignal) { [string]$row.Pattern_PeakDay } else { '' })))
        }
        $tg = [ordered]@{}
        foreach ($p in $row.PSObject.Properties) { if ($p.Name.StartsWith('Tag.') -and -not [string]::IsNullOrEmpty([string]$p.Value)) { $tg[$p.Name.Substring(4)] = [string]$p.Value } }
        if ($tg.Count) { $o.tg = $tg }
        if ($mc.Count) { $o.m = @(foreach ($c in $mc) { Jn $row.$c }) }
        if ($fc.Count) { $o.f = @(foreach ($c in $fc) { Jv $row.$c }) }
        if ($pc.Count) { $o.p = @(foreach ($c in $pc) { Jv $row.$c }) }
        $rid = ([string]$row.ResourceId).ToLower()
        if ($rid -and $byRid.ContainsKey($rid)) {
            $pm = [string]$row.PrimaryMetric; $dm = [ordered]@{}
            foreach ($e in $byRid[$rid]) {
                if ($e.dim -and $child -and $e.dim -ne $child) { continue }
                $isTotal = $colSet.Contains("$($e.m)_Total") -and -not $colSet.Contains("$($e.m)_Avg") -and -not $colSet.Contains("$($e.m)_P95")
                if ($e.m -eq $pm -and -not $o.Contains('d') -and (-not $e.dim -or $e.dim -eq $child)) {
                    $o.d = if ($isTotal) { $e.t } else { $e.a }; $o.dk = if ($isTotal) { 't' } else { 'a' }; $dailyUsed += $nD
                }
                if ($dailyUsed -ge $HtmlDailyBudget) { continue }
                $label = if ($e.dim -and -not $child) { "$($e.m) [$($e.dim)]" } else { $e.m }
                $dm[$label] = if ($isTotal) { [ordered]@{ k = 't'; v = $e.t } } else { [ordered]@{ k = 'a'; v = $e.a; x = $e.x } }
                $dailyUsed += $nD * $(if ($isTotal) { 1 } else { 2 })
            }
            if ($dm.Count) { $o.dm = $dm }
        }
        $jsRes.Add($o)
    }
}

$jsRecs = @(foreach ($r in $RecsSorted) {
        $ri = $resIdxByKey["$($r.TypeKey)|$($r.ResourceId)|$($r.ResourceName)".ToLower()]
        [ordered]@{ r = $(if ($null -ne $ri) { $ri } else { -1 }); t = $(if ($typeIdx.ContainsKey([string]$r.TypeKey)) { $typeIdx[[string]$r.TypeKey] } else { -1 }); pr = [string]$r.Priority; id = [string]$r.RuleId; rc = [string]$r.RuleCode
            cat = [string]$r.Category; act = [string]$r.RecommendedAction; tgt = [string]$r.TargetConfiguration; sa = Jn $r.EstAnnualSavings; sm = Jn $r.EstMonthlySavings; cmc = Jn $r.CurrentMonthlyCost
            ev = [string]$r.Evidence; conf = [string]$r.Confidence; eff = [string]$r.Effort; cmt = [string]$r.Comment; cur = [string]$r.CurrentConfiguration; obs = [string]$r.ObservedUtilization
            n = [string]$r.ResourceName; s = Get-SubIdx ([string]$r.SubscriptionName); g = [string]$r.ResourceGroup; e = [string]$r.Environment; l = [string]$r.Location; rid = [string]$r.ResourceId }
    })
$jsRules = [ordered]@{}
foreach ($code in $script:RuleCatalog.Keys) {
    $rc = $script:RuleCatalog[$code]
    $jsRules[$code] = [ordered]@{ n = $rc.Name; c = (Format-RuleText $rc.Criteria); r = $rc.Rationale; ap = @($(if ($RuleApplies.ContainsKey($code)) { @($RuleApplies[$code]) } else { @() })) }
}

# Log Analytics / Sentinel: daily ingestion per workspace + per table
$laDailyBy = @{}
foreach ($x in $LaDaily) {
    $k = "$($x.SubscriptionName)|$($x.ResourceGroup)|$($x.Workspace)".ToLower(); $di = $dateIdx[[string]$x.Date]
    if ($null -eq $di) { continue }
    if (-not $laDailyBy.ContainsKey($k)) { $laDailyBy[$k] = [object[]]::new($nD) }
    $laDailyBy[$k][$di] = Jn $x.BillableGB
}
$laTblBy = @{}
foreach ($x in $LaByTable) {
    $k = "$($x.SubscriptionName)|$($x.ResourceGroup)|$($x.Workspace)".ToLower(); $di = $dateIdx[[string]$x.Date]
    if ($null -eq $di) { continue }
    if (-not $laTblBy.ContainsKey($k)) { $laTblBy[$k] = [ordered]@{} }
    $tn = [string]$x.DataType
    if (-not $laTblBy[$k].Contains($tn)) { $laTblBy[$k][$tn] = [object[]]::new($nD) }
    $laTblBy[$k][$tn][$di] = Jn $x.BillableGB
}
$jsLa = @(foreach ($a in $script:LaAnalysis) {
        $k = "$($a.SubscriptionName)|$($a.ResourceGroup)|$($a.Workspace)".ToLower()
        $tb = @()
        if ($laTblBy.ContainsKey($k)) {
            $tb = @($laTblBy[$k].Keys | ForEach-Object { [ordered]@{ n = $_; v = $laTblBy[$k][$_]; tot = [double](($laTblBy[$k][$_] | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum) } } |
                Sort-Object { if ($_.n -eq 'Other tables') { -1 } else { $_.tot } } -Descending | ForEach-Object { [ordered]@{ n = $_.n; v = $_.v } })
        }
        [ordered]@{ s = Get-SubIdx ([string]$a.SubscriptionName); g = [string]$a.ResourceGroup; n = [string]$a.Workspace; l = [string]$a.Location; ri = $resIdxBySRN["$($a.SubscriptionName)|$($a.ResourceGroup)|$($a.Workspace)".ToLower()]
            tier = [string]$a.LAPricingTier; com = $(if ([string]$a.LAPricingTier -match '^(\d+)') { [double]$Matches[1] } else { $null }); sen = [bool]$a.SentinelEnabled; senp = [string]$a.SentinelPricing; senc = Jn $a.SentinelCommitmentGBPerDay
            ret = Jn $a.RetentionDays; avg = Jn $a.AvgDailyGB; p95 = Jn $a.P95DailyGB; max = Jn $a.MaxDailyGB; min = Jn $a.MinDailyGB; wd = Jn $a.WeekdayAvgGB; we = Jn $a.WeekendAvgGB
            cu = Jn $a.CommitmentUtilizationPct; below = Jn $a.DaysBelowCommitment; above = Jn $a.DaysAboveCommitment
            cla = Jn $a.CurrentLAMonthlyEst; ola = [string]$a.OptimalLATier; olac = Jn $a.OptimalLAMonthlyEst; cse = Jn $a.CurrentSentinelMonthlyEst; ose = [string]$a.OptimalSentinelTier; osec = Jn $a.OptimalSentinelMonthlyEst
            sav = Jn $a.EstMonthlySavings; st = [string]$a.Status; src = [string]$a.PriceSource; d = $laDailyBy[$k]; tb = $tb }
    })

# Cost cubes (month x subscription x service, day x subscription x service) - the dashboard filters them by subscription
$jsSvc = [System.Collections.Generic.List[string]]::new(); $svcIdx = @{}
function Get-SvcIdx([string]$Name) { if (-not $Name) { $Name = '(unassigned)' }; if (-not $svcIdx.ContainsKey($Name)) { $svcIdx[$Name] = $jsSvc.Count; $jsSvc.Add($Name) }; return $svcIdx[$Name] }
$monthIdx = @{}; for ($i = 0; $i -lt $Months.Count; $i++) { $monthIdx[[string]$Months[$i]] = $i }
$useUsdM2 = @($CostMonthly | Where-Object { $_.CostUSD }).Count -gt 0
$cm = @(foreach ($x in $CostMonthly) {
        $v = Jn $(if ($useUsdM2) { $x.CostUSD } else { $x.Cost }); $mi = $monthIdx[[string]$x.Month]
        if (-not $v -or $null -eq $mi) { continue }
        , @($mi, (Get-SubIdx ([string]$x.SubscriptionName)), (Get-SvcIdx ([string]$x.ServiceName)), $v)
    })
$useUsdD2 = @($CostDaily | Where-Object { $_.CostUSD }).Count -gt 0
$cd = @(foreach ($x in $CostDaily) {
        $v = Jn $(if ($useUsdD2) { $x.CostUSD } else { $x.Cost }); $di = $dateIdx[[string]$x.Date]
        if (-not $v -or $null -eq $di) { continue }
        , @($di, (Get-SubIdx ([string]$x.SubscriptionName)), (Get-SvcIdx ([string]$x.ServiceName)), $v)
    })
$partialMonths = @($MonthlyTotals | Where-Object PartialMonth | ForEach-Object { [string]$_.Month })

$jsOrph = @(foreach ($o in $Orphans) {
        [ordered]@{ f = [string]$o.Finding; ci = [string]$o.CostImpact; s = Get-SubIdx ([string]$o.SubscriptionName); g = [string]$o.ResourceGroup; n = [string]$o.ResourceName; t = [string]$o.ArmType
            l = [string]$o.Location; sku = [string]$o.Sku; d = [string]$o.Detail; id = [string]$o.ResourceId; c = Jn (Get-MonthlyCost $o.ResourceId) }
    })
$jsOther = @(foreach ($o in $OtherPaid) { [ordered]@{ s = Get-SubIdx ([string]$o.SubscriptionName); g = [string]$o.ResourceGroup; n = [string]$o.ResourceName; svc = [string]$o.ServiceName; c = Jn $o.MonthlyCost; t = [string]$o.ResourceType; id = [string]$o.ResourceId } })
$jsAdv = @(foreach ($a in $Advisor) {
        [ordered]@{ cat = [string]$a.Category; s = Get-SubIdx ([string]$a.SubscriptionName); g = [string]$a.ResourceGroup; typ = [string]$a.ImpactedType; res = [string]$a.ImpactedResource; imp = [string]$a.Impact
            prob = [string]$a.Problem; sol = [string]$a.Solution; cur = [string]$a.CurrentSku; tgt = [string]$a.TargetSku; term = [string]$a.Term; as = Jn $a.AnnualSavings; id = [string]$a.ResourceId }
    })
$jsInv = @($Inventory | Group-Object ArmType | ForEach-Object { [ordered]@{ t = [string]$_.Name; n = $_.Count; k = [string](@($_.Group | Where-Object CatalogKey | Select-Object -First 1).CatalogKey) } } | Sort-Object { $_.n } -Descending)
$jsCov = @(foreach ($c in $CollectionSummary) { [ordered]@{ k = [string]$c.Key; n = [string]($script:SheetNames[$c.Key] ?? $c.Display); c = [string]$c.Category; r = [int]((N $c.Rows) ?? 0); ok = [int]((N $c.MetricsOK) ?? 0); nd = [int]((N $c.MetricsNoData) ?? 0); er = [int]((N $c.MetricsError) ?? 0); g = (B $c.Generic) } })
$jsMcat = @(foreach ($m in $MetricCatalog) { [ordered]@{ t = [string]$m.ResourceType; a = [string]$m.Alias; m = [string]$m.AzureMetric; o = [string]$m.Outputs; sub = [string]$m.SubResource; f = [string]$m.DimensionFilter } })
foreach ($g in (Import-RawCsv 'Reference_GenericMetrics.csv')) { if ($g.Alias -and -not ($jsMcat | Where-Object { $_.t -eq $g.Display -and $_.a -eq $g.Alias })) { $jsMcat += [ordered]@{ t = [string]$g.Display; a = [string]$g.Alias; m = [string]$g.Metric; o = "Role: $($g.Role)"; sub = ''; f = '' } } }

$HtmlData = [ordered]@{
    meta  = [ordered]@{
        customer = $CustomerName; generated = (Get-Date).ToString('yyyy-MM-dd HH:mm'); version = '2.0'; start = [string]$Meta.WindowStartLocal; end = [string]$Meta.WindowEndLocal
        days = $LookbackDays; tz = [string]$Meta.TimeZoneId; bizStart = [int]$Meta.BusinessHoursStart; bizEnd = [int]$Meta.BusinessHoursEnd; weekend = @($Meta.WeekendDays)
        subs = [int]$Meta.SubscriptionCount; inventory = [int]$Meta.InventoryCount; analyzed = $AllRows.Count; currency = $script:Currency; costType = [string]$Meta.CostType
        hasCost = [bool]$script:HasCost; totalMonthly = Jn $TotalMonthlyCost; annualSavings = Jn $AnnualSavings; metricRequests = [int]((N $Meta.MetricCallsTotal) ?? 0)
        batch = [bool]$Meta.MetricsBatchApi; collectedBy = [string]$Meta.CollectedBy; collectedAt = [string]$Meta.CollectedAtUtc; tenant = [string]$Meta.TenantId; sources = $Meta.DataSources
        th = [ordered]@{ idle = $CpuIdleThreshold; low = $CpuLowThreshold; high = $CpuHighThreshold; memLow = $MemLowThreshold; memHigh = $MemHighThreshold; target = $TargetUtilization
            genLow = $GenericLow; genHigh = $GenericHigh; steady = $SteadyHoursPct; snap = $SnapshotMaxAge }
        limits = [ordered]@{ props = $propCap; dailyBudget = $HtmlDailyBudget; dailyUsed = $dailyUsed }
        savedPath = $HtmlPath
    }
    dates = @($WindowDates); subs = $jsSubs; sg = @(foreach ($k in $StatusGroups.Keys) { [ordered]@{ n = $k; s = @($StatusGroups[$k]) } })
    types = $jsTypes; res = $jsRes; recs = $jsRecs; rules = $jsRules; la = $jsLa
    cost = [ordered]@{ months = @($Months | ForEach-Object { [string]$_ }); partial = $partialMonths; svcs = $jsSvc; cm = $cm; cd = $cd }
    orph = $jsOrph; other = $jsOther; adv = $jsAdv; ri = @($Reservations); sp = @($SavingsPlans); def = @($Defender); inv = $jsInv; cov = $jsCov; mcat = $jsMcat
    method = @($Methodology); scope = @($Subscriptions | Select-Object SubscriptionName, SubscriptionId, State)
}
$json = $HtmlData | ConvertTo-Json -Depth 12 -Compress -EscapeHandling EscapeHtml
$json = $json.Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')

#endregion

#region ============================== EMBEDDED DASHBOARD TEMPLATE ==============================

# Embedded so the solution needs only its three scripts (see the HTML comment for the data contract)
$script:EmbeddedDashboardTemplate = @'
<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Azure Utilization &amp; Cost Optimization Assessment v2">
<title>__TITLE__</title>
<!--
  Azure Utilization & Cost Optimization Dashboard (template v2)
  Phase2-AnalyzeAndReport.ps1 injects the charting library and the assessment data model.
  Data contract (window.D): meta, dates[], subs[], sg[] (status groups), types[] {k,n,c,a,pfx,pm,mc[],fc[],pc[],ml{},mu{},hw{}},
  res[] {t,s,g,n,l,e,st,u,es,c,cc,x,sv,pv,us,rc,rp,sku,k,ag,cr,pm,ms,id,an,ed,rpm,pat[],tg{},m[],f[],p[],d[],dk,dm{}},
  recs[] {r,t,pr,id,rc,cat,act,tgt,sa,sm,cmc,ev,conf,eff,cmt,cur,obs,n,s,g,e,l,rid}, rules{code:{n,c,r,ap[]}}, la[], cost{months,partial,svcs,cm,cd},
  orph[], other[], adv[], ri[], sp[], def[], inv[], cov[], mcat[], method[], scope[]
-->
<style>
:root{--bg:#F2F5FA;--card:#FFFFFF;--card2:#F6F8FC;--text:#172033;--muted:#5D6B82;--line:#E2E8F1;--brand:#0F4C81;--brand2:#0078D4;--pos:#0B7A0B;--shadow:0 1px 2px rgba(16,24,40,.05),0 8px 24px rgba(16,24,40,.06);--radius:14px;--hero1:#0B2545;--hero2:#13315C;--hero3:#3B2A73}
[data-theme=dark]{--bg:#0A1120;--card:#111B2E;--card2:#16223A;--text:#E4EBF5;--muted:#95A5BE;--line:#24334F;--brand2:#3AA0FF;--pos:#4FD18B;--shadow:0 1px 2px rgba(0,0,0,.35),0 10px 30px rgba(0,0,0,.35);--hero1:#071226;--hero2:#0E2547;--hero3:#2B1F5C}
*{box-sizing:border-box}
html,body{margin:0}
body{font-family:"Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);font-size:13.5px;line-height:1.45;-webkit-font-smoothing:antialiased}
a{color:var(--brand2);text-decoration:none}a:hover{text-decoration:underline}
code{font-family:Consolas,"Cascadia Mono",monospace;font-size:12px;background:var(--card2);border:1px solid var(--line);border-radius:5px;padding:0 5px}
kbd{font-family:inherit;font-size:11px;border:1px solid var(--line);border-bottom-width:2px;border-radius:5px;padding:0 5px;background:var(--card);color:var(--muted)}
.muted{color:var(--muted)}.sm{font-size:12px}.pos{color:var(--pos)}.neg{color:#D13438}.sp{flex:1}.nowrap{white-space:nowrap}
#app{display:flex;min-height:100vh}
#side{width:250px;flex:0 0 250px;background:linear-gradient(180deg,#0B2545 0%,#12305A 55%,#1D3F78 100%);color:#C9D6EA;position:sticky;top:0;height:100vh;display:flex;flex-direction:column;overflow-y:auto;z-index:30;transition:width .2s,flex-basis .2s}
body.collapsed #side{width:66px;flex-basis:66px}
.brand{display:flex;gap:11px;align-items:center;padding:18px 15px 15px;border-bottom:1px solid rgba(255,255,255,.08)}
.logo{width:40px;height:40px;flex:0 0 40px;border-radius:10px;background:#fff;display:grid;place-items:center;box-shadow:0 6px 16px rgba(0,0,0,.25)}
.b1{color:#fff;font-weight:700;font-size:13.5px;line-height:1.2}.b2{font-size:11.5px;opacity:.8;margin-top:2px}
#nav{padding:10px 8px;display:flex;flex-direction:column;gap:2px}
.nav-i{display:flex;align-items:center;gap:12px;padding:9px 12px;border-radius:10px;color:#C9D6EA;cursor:pointer;font-weight:500;white-space:nowrap;text-decoration:none!important}
.nav-i:hover{background:rgba(255,255,255,.07);color:#fff}
.nav-i.on{background:rgba(255,255,255,.14);color:#fff;box-shadow:inset 3px 0 0 #4CC2FF}
.nav-i svg{flex:0 0 18px}
.nav-i .bdg{margin-left:auto;background:rgba(255,255,255,.16);border-radius:10px;padding:0 7px;font-size:11px}
body.collapsed .nav-i span,body.collapsed .brand>div:last-child,body.collapsed .side-foot{display:none}
.side-foot{margin-top:auto;padding:14px 16px;font-size:11px;opacity:.75;border-top:1px solid rgba(255,255,255,.08);line-height:1.6}
#main{flex:1;min-width:0;display:flex;flex-direction:column}
#top{position:sticky;top:0;z-index:20;display:flex;align-items:center;gap:14px;padding:11px 24px;background:color-mix(in srgb,var(--bg) 82%,transparent);backdrop-filter:blur(12px);border-bottom:1px solid var(--line)}
#top h1{margin:0;font-size:19px;font-weight:700;letter-spacing:-.01em}
#top .sub{font-size:12px;color:var(--muted)}
.top-actions{margin-left:auto;display:flex;gap:8px;align-items:center}
.icon-btn{width:36px;height:36px;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--text);display:grid;place-items:center;cursor:pointer}
.icon-btn:hover{border-color:var(--brand2);color:var(--brand2)}
.search-btn{display:flex;align-items:center;gap:8px;min-width:300px;border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:10px;padding:7px 12px;cursor:pointer;font:inherit}
.search-btn kbd{margin-left:auto}
#filters{padding:12px 24px 2px;display:flex;flex-wrap:wrap;gap:7px;align-items:center}
.fl{display:flex;align-items:center;gap:6px;color:var(--muted);font-weight:600;margin-right:4px}
.fbtn{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:999px;padding:5px 11px;font:inherit;font-size:12.5px;cursor:pointer;display:inline-flex;gap:6px;align-items:center;transition:all .12s}
.fbtn:hover{border-color:var(--brand2)}
.fbtn.act{border-color:var(--brand2);background:color-mix(in srgb,var(--brand2) 10%,var(--card));color:var(--brand2);font-weight:600}
.fbtn .n{background:var(--brand2);color:#fff;border-radius:9px;padding:0 6px;font-size:11px}
.fsearch{display:flex;align-items:center;gap:6px;border:1px solid var(--line);background:var(--card);border-radius:999px;padding:4px 12px;color:var(--muted)}
.fsearch input{border:0;outline:0;background:transparent;color:var(--text);font:inherit;width:220px}
.fcount{color:var(--muted);font-size:12.5px;margin-left:4px}
.lnk{border:0;background:none;color:var(--brand2);cursor:pointer;font:inherit;font-weight:600;padding:2px 4px}
#fchips{display:flex;flex-wrap:wrap;gap:6px;width:100%}
#fchips:empty{display:none}
.chip{display:inline-flex;align-items:center;gap:6px;border-radius:8px;background:color-mix(in srgb,var(--brand2) 12%,var(--card));color:var(--text);padding:3px 4px 3px 9px;font-size:12px}
.chip b{font-weight:600;color:var(--muted)}
.chip button{border:0;background:none;color:var(--muted);cursor:pointer;font-size:14px;line-height:1;padding:0 4px}
#page{padding:14px 24px 34px;display:flex;flex-direction:column;gap:16px}
.fade-in{animation:fi .35s ease both}@keyframes fi{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(2,minmax(0,1fr))}.g3{grid-template-columns:repeat(3,minmax(0,1fr))}.g4{grid-template-columns:repeat(4,minmax(0,1fr))}
.g21{grid-template-columns:minmax(0,2fr) minmax(0,1fr)}.g12{grid-template-columns:minmax(0,1fr) minmax(0,2fr)}
@media (max-width:1300px){.g3,.g4{grid-template-columns:repeat(2,minmax(0,1fr))}.g21,.g12{grid-template-columns:1fr}}
@media (max-width:860px){.g2,.g3,.g4{grid-template-columns:1fr}#side{display:none}.search-btn{min-width:0}.search-btn span{display:none}.ms-brand span{display:none}}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);display:flex;flex-direction:column;min-width:0}
.card-h{display:flex;align-items:center;gap:10px;padding:13px 16px 4px}
.card-h h3{margin:0;font-size:14px;font-weight:650;letter-spacing:-.005em}
.card-a{margin-left:auto;display:flex;gap:6px;align-items:center}
.card-b{padding:8px 16px 15px;min-width:0;flex:1}
.info{display:inline-grid;place-items:center;width:16px;height:16px;border-radius:50%;border:1px solid var(--line);color:var(--muted);font-size:10px;font-weight:700;cursor:help;margin-left:4px;vertical-align:1px}
.chart{width:100%;min-height:120px}
.nochart,.empty{color:var(--muted);text-align:center;padding:26px 10px;font-size:12.5px}
.kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(205px,1fr));gap:14px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:14px 15px;box-shadow:var(--shadow);display:flex;gap:12px;position:relative;overflow:hidden}
.kpi::after{content:'';position:absolute;right:-28px;bottom:-34px;width:96px;height:96px;border-radius:50%;background:var(--kc,#0078D4);opacity:.08}
.kpi-i{width:38px;height:38px;flex:0 0 38px;border-radius:11px;background:color-mix(in srgb,var(--kc,#0078D4) 15%,transparent);color:var(--kc,#0078D4);display:grid;place-items:center}
.kpi-l{font-size:11px;color:var(--muted);font-weight:650;text-transform:uppercase;letter-spacing:.05em}
.kpi-v{font-size:23px;font-weight:750;letter-spacing:-.02em;line-height:1.25}
.kpi-s{font-size:12px;color:var(--muted)}
.hero{position:relative;overflow:hidden;background:radial-gradient(900px 280px at 8% -30%,rgba(76,194,255,.38),transparent),radial-gradient(700px 300px at 100% 120%,rgba(135,100,184,.45),transparent),linear-gradient(120deg,var(--hero1),var(--hero2) 48%,var(--hero3));color:#fff;border-radius:18px;padding:24px 28px;display:flex;gap:26px;align-items:center;box-shadow:var(--shadow)}
.hero h2{margin:6px 0 8px;font-size:23px;font-weight:700;letter-spacing:-.015em;max-width:760px}
.hero p{margin:0;opacity:.82;font-size:13px}
.eyebrow{text-transform:uppercase;letter-spacing:.09em;font-size:11px;font-weight:700;opacity:.8}
.hero-k{margin-left:auto;display:flex;gap:28px;align-items:center}
.hero-k>div{display:flex;flex-direction:column}
.hero-k span{font-size:11px;text-transform:uppercase;letter-spacing:.07em;opacity:.75;font-weight:650}
.hero-k b{font-size:30px;font-weight:750;letter-spacing:-.02em;line-height:1.15}
.hero-k b.g{color:#7CF0B0}
.hero-k small{opacity:.75;font-size:12px}
.hero-ring{align-items:center}
@media (max-width:1100px){.hero{flex-direction:column;align-items:flex-start}.hero-k{margin-left:0;flex-wrap:wrap}}
.pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11.5px;font-weight:650;white-space:nowrap;background:color-mix(in srgb,var(--pc) 14%,transparent);color:var(--pc);line-height:1.5}
[data-theme=dark] .pill{background:color-mix(in srgb,var(--pc) 24%,transparent);color:color-mix(in srgb,var(--pc) 70%,#fff)}
.ubar{width:64px;height:6px;border-radius:3px;background:var(--line);display:inline-block;vertical-align:middle;margin-right:7px;overflow:hidden}
.ubar i{display:block;height:100%;border-radius:3px}
.ring{display:inline-block;vertical-align:middle}
.ring.na{display:inline-grid;place-items:center;border-radius:50%;border:2px dashed var(--line);color:var(--muted);font-size:10px}
.spark{display:block}
.btn{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:9px;padding:6px 12px;font:inherit;font-size:12.5px;cursor:pointer;font-weight:600;display:inline-flex;gap:6px;align-items:center;white-space:nowrap;text-decoration:none!important}
.btn:hover{border-color:var(--brand2);color:var(--brand2)}
.btn.pri{background:var(--brand2);border-color:var(--brand2);color:#fff}
.btn.pri:hover{filter:brightness(1.08);color:#fff}
.btn.sm{padding:4px 9px;font-size:12px}
.btn[disabled]{opacity:.45;cursor:default}
select.sel{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:8px;padding:5px 8px;font:inherit;font-size:12.5px;max-width:320px}
.ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:6px;font-size:12.5px;color:var(--muted)}
.tabs{display:flex;gap:2px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.tab{padding:8px 14px;border:0;background:none;color:var(--muted);font:inherit;font-weight:650;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px;text-decoration:none!important}
.tab:hover{color:var(--text)}
.tab.on{color:var(--brand2);border-bottom-color:var(--brand2)}
.tab .n{font-size:11px;background:var(--card2);border:1px solid var(--line);border-radius:9px;padding:0 6px;margin-left:5px;color:var(--muted)}
.dt-bar{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.dt-s{display:flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:8px;padding:4px 9px;color:var(--muted);background:var(--card)}
.dt-s input{border:0;outline:0;background:transparent;color:var(--text);font:inherit;width:220px}
.dt-n{color:var(--muted);font-size:12px}
.tw{overflow:auto;max-height:620px;border:1px solid var(--line);border-radius:10px}
table.t{border-collapse:separate;border-spacing:0;width:100%;font-size:12.5px}
table.t th{position:sticky;top:0;background:var(--card2);text-align:left;font-weight:650;padding:8px 10px;border-bottom:1px solid var(--line);white-space:nowrap;cursor:pointer;user-select:none;z-index:2;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.03em}
table.t th:hover{color:var(--text)}
table.t td{padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:middle}
table.t tbody tr:last-child td{border-bottom:0}
table.t tr.clk{cursor:pointer}
table.t tbody tr:hover td{background:color-mix(in srgb,var(--brand2) 6%,transparent)}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
th.ck,td.ck{width:30px;text-align:center}
.clip{max-width:440px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.wrap{white-space:normal;min-width:260px;max-width:520px}
.rn{min-width:190px}.rn b{display:block;font-weight:650}.rn small{color:var(--muted);font-size:11.5px}
.dt-p{display:flex;gap:10px;align-items:center;justify-content:flex-end;margin-top:8px;font-size:12px;color:var(--muted)}
.dt-p:empty{display:none}
#dd{position:absolute;z-index:70;background:var(--card);border:1px solid var(--line);border-radius:12px;box-shadow:0 18px 50px rgba(0,0,0,.22);width:330px;display:none}
#dd.open{display:block}
.dd-h{padding:10px;border-bottom:1px solid var(--line)}
.dd-h input,.dd-h select{width:100%;border:1px solid var(--line);border-radius:8px;padding:6px 9px;background:var(--card2);color:var(--text);font:inherit;outline:0}
.dd-h select{margin-bottom:7px}
.dd-a{display:flex;justify-content:space-between;margin-top:6px}
.dd-l{max-height:330px;overflow-y:auto;padding:5px}
.dd-o{display:flex;align-items:center;gap:8px;padding:5px 8px;border-radius:7px;cursor:pointer;font-size:12.5px}
.dd-o:hover{background:var(--card2)}
.dd-t{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dd-c{color:var(--muted);font-size:11.5px;font-variant-numeric:tabular-nums}
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:7px;vertical-align:1px}
#drawerMask{position:fixed;inset:0;background:rgba(8,16,32,.38);opacity:0;pointer-events:none;transition:opacity .2s;z-index:45}
#drawerMask.open{opacity:1;pointer-events:auto}
#drawer{position:fixed;top:0;right:0;height:100vh;width:min(800px,96vw);background:var(--bg);box-shadow:-12px 0 44px rgba(0,0,0,.28);transform:translateX(105%);transition:transform .25s ease;z-index:50;overflow-y:auto}
#drawer.open{transform:none}
.dr-h{position:sticky;top:0;z-index:3;display:flex;gap:16px;align-items:flex-start;padding:20px 22px 16px;background:linear-gradient(120deg,var(--hero1),var(--hero2) 60%,var(--hero3));color:#fff}
.dr-h h2{margin:2px 0 8px;font-size:20px;word-break:break-all}
.dr-h .pill{background:rgba(255,255,255,.16);color:#fff}
.dr-t{font-size:11px;text-transform:uppercase;letter-spacing:.08em;opacity:.78;font-weight:700}
.dr-r{margin-left:auto;text-align:center;color:#fff}
.dr-r small{display:block;opacity:.8;font-size:11px}
.dr-x{background:rgba(255,255,255,.12);border-color:rgba(255,255,255,.25);color:#fff}
.dr-b{padding:18px 22px 40px;display:flex;flex-direction:column;gap:16px}
.dr-b h3{margin:6px 0 0;font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.facts{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}
.facts>div{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:8px 11px;min-width:0}
.facts span{display:block;font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
.facts b{display:block;font-weight:650;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pvu{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.pvu>div{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.pvu h4{margin:0 0 4px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.pvu p{margin:0}
.gauge{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.gbar{position:relative;height:12px;border-radius:6px;background:linear-gradient(90deg,#D13438 0%,#F7630C 18%,#13A10E 45%,#13A10E 80%,#8764B8 100%);opacity:.9;margin:22px 0 6px}
.gmark{position:absolute;top:-7px;width:3px;height:26px;background:var(--text);border-radius:2px}
.gmark span{position:absolute;top:-20px;left:50%;transform:translateX(-50%);font-size:11px;font-weight:700;white-space:nowrap}
.gtick{position:absolute;top:13px;font-size:10px;color:var(--muted);transform:translateX(-50%)}
.note{background:color-mix(in srgb,#C19C00 12%,var(--card));border:1px solid color-mix(in srgb,#C19C00 35%,var(--line));border-radius:10px;padding:9px 12px;font-size:12.5px}
.rcard{background:var(--card);border:1px solid var(--line);border-left:4px solid var(--pc);border-radius:12px;padding:11px 14px;display:flex;flex-direction:column;gap:5px}
.rc-h{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.rc-a{font-weight:650}
.rc-k{font-size:12.5px;display:grid;grid-template-columns:130px 1fr;gap:8px}
.rc-k span{color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.04em;padding-top:1px}
.tagc{display:inline-flex;gap:4px;align-items:center;border:1px solid var(--line);background:var(--card);border-radius:7px;padding:3px 8px;font-size:12px;margin:0 6px 6px 0;cursor:pointer}
.tagc b{color:var(--muted);font-weight:600}
.tagc:hover{border-color:var(--brand2)}
.kv{width:100%;border-collapse:collapse;font-size:12.5px}
.kv td{padding:6px 9px;border-bottom:1px solid var(--line);vertical-align:top;word-break:break-word}
.kv td:first-child{color:var(--muted);width:38%;font-family:Consolas,monospace;font-size:11.5px}
.kvw{background:var(--card);border:1px solid var(--line);border-radius:12px;max-height:440px;overflow:auto}
#palette{position:fixed;inset:0;background:rgba(8,16,32,.5);display:none;z-index:80;align-items:flex-start;justify-content:center;padding-top:11vh}
#palette.open{display:flex}
.pal-box{width:min(700px,94vw);background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:0 24px 70px rgba(0,0,0,.4);overflow:hidden}
#palInput{width:100%;border:0;padding:17px 20px;font-size:16px;background:transparent;color:var(--text);outline:none;border-bottom:1px solid var(--line);font-family:inherit}
#palList{max-height:430px;overflow-y:auto;padding:6px}
.pal-g{font-size:11px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);padding:8px 14px 4px;font-weight:700}
.pal-i{padding:8px 14px;display:flex;gap:10px;align-items:center;cursor:pointer;border-radius:9px}
.pal-i small{color:var(--muted);margin-left:auto;font-size:11.5px}
.pal-i.on{background:color-mix(in srgb,var(--brand2) 14%,transparent)}
#toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(30px);background:#172033;color:#fff;padding:9px 16px;border-radius:10px;opacity:0;transition:all .25s;z-index:90;font-size:12.5px;pointer-events:none}
#toast.show{opacity:1;transform:translateX(-50%)}
.tcards{display:grid;grid-template-columns:repeat(auto-fill,minmax(232px,1fr));gap:12px}
.tcard{background:var(--card);border:1px solid var(--line);border-radius:13px;padding:12px 14px;cursor:pointer;transition:transform .12s,box-shadow .12s,border-color .12s;display:flex;flex-direction:column;gap:6px;text-decoration:none!important;color:var(--text)}
.tcard:hover{transform:translateY(-2px);box-shadow:var(--shadow);border-color:var(--brand2)}
.tcard.on{outline:2px solid var(--brand2);border-color:transparent}
.tc-h{display:flex;gap:10px;align-items:center}
.tc-h b{font-size:13.5px;line-height:1.2}
.tc-h small{display:block;color:var(--muted);font-size:11px}
.tc-m{display:flex;justify-content:space-between;font-size:12px;color:var(--muted)}
.tc-m b{color:var(--text);font-weight:650}
.badge{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;border-radius:6px;padding:1px 6px;background:var(--card2);border:1px solid var(--line);color:var(--muted)}
.badge.auto{color:#8764B8;border-color:color-mix(in srgb,#8764B8 40%,var(--line))}
.insights{display:flex;flex-direction:column;gap:9px}
.ins{display:flex;gap:11px;padding:10px 12px;border-radius:11px;background:var(--card2);border:1px solid var(--line);font-size:13px;align-items:flex-start}
.ins i{flex:0 0 28px;height:28px;border-radius:8px;display:grid;place-items:center;background:color-mix(in srgb,var(--ic) 16%,transparent);color:var(--ic)}
.ins b{font-weight:700}
.summary{font-size:13.5px;line-height:1.6;background:color-mix(in srgb,var(--brand2) 6%,var(--card));border:1px solid color-mix(in srgb,var(--brand2) 22%,var(--line));border-radius:12px;padding:12px 15px}
.toplist{display:flex;flex-direction:column;gap:7px}
.tl{display:grid;grid-template-columns:4px 1fr auto;gap:11px;align-items:center;padding:7px 10px 7px 0;border-radius:10px;cursor:pointer}
.tl:hover{background:var(--card2)}
.tl i{height:34px;border-radius:3px;background:var(--pc)}
.tl b{display:block;font-weight:650;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tl small{display:block;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px}
.tl em{font-style:normal;font-weight:750;color:var(--pos)}
.capbar{margin:10px 0 14px}
.capbar .lbl{display:flex;justify-content:space-between;font-size:12.5px;margin-bottom:5px}
.capbar .bar{position:relative;height:14px;border-radius:7px;background:var(--card2);border:1px solid var(--line);overflow:hidden}
.capbar .bar i{position:absolute;left:0;top:0;bottom:0;border-radius:7px}
.whatif{position:sticky;top:62px;z-index:6;display:flex;flex-wrap:wrap;gap:16px;align-items:center;padding:11px 15px;border-radius:13px;background:linear-gradient(110deg,color-mix(in srgb,var(--brand2) 12%,var(--card)),var(--card));border:1px solid color-mix(in srgb,var(--brand2) 30%,var(--line));box-shadow:var(--shadow)}
.whatif b{font-size:16px}
.whatif .lbl{font-size:11px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;display:block}
.tdh{display:flex;align-items:flex-end;gap:12px}
.tdh h2{margin:2px 0 0;font-size:21px}
.legend-s{display:flex;flex-wrap:wrap;gap:12px;font-size:12px;color:var(--muted)}
.legend-s span{display:inline-flex;align-items:center;gap:5px}
.ws-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:10px}
.envc{border-radius:12px;padding:13px 15px;border:1px solid var(--line);background:var(--card)}
.envc h4{margin:0 0 6px;font-size:13px}
.envc .v{font-size:21px;font-weight:750}
.ms-brand{display:flex;align-items:center;gap:8px;padding-left:12px;margin-left:4px;border-left:1px solid var(--line);font-weight:600;font-size:15px;color:var(--text);letter-spacing:-.01em;white-space:nowrap}footer#foot{margin:6px 24px 24px;display:flex;justify-content:space-between;align-items:flex-end;gap:24px;flex-wrap:wrap;background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:14px 18px;color:var(--muted);font-size:12px}.foot-l{min-width:0;flex:1 1 420px}.foot-k{font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;font-weight:700;color:var(--muted)}.foot-path{font-family:Consolas,"Cascadia Mono",monospace;font-size:12px;color:var(--text);word-break:break-all;margin:3px 0 6px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}.foot-meta{font-size:11.5px;line-height:1.5}.ms-credit{flex:0 0 auto;min-width:230px;border-left:4px solid #0078D4;background:var(--card2);border-radius:10px;padding:10px 16px;display:flex;flex-direction:column;gap:2px}.ms-credit .ms-word{display:flex;align-items:center;gap:8px;font-weight:600;font-size:15px;color:var(--text);margin-bottom:6px}.ms-credit .cred-n{font-weight:700;color:var(--text);font-size:13.5px}.ms-credit .cred-o{color:var(--text)}
@media print{#side,#top,#filters,.no-print,#drawer,#drawerMask,#palette,#dd,.whatif{display:none!important}#page{padding:0}.card,.kpi{break-inside:avoid;box-shadow:none}body{background:#fff}.tw{max-height:none;overflow:visible}}
</style>
<!--__ECHARTS__-->
</head>
<body>
<div id="app">
  <aside id="side">
    <div class="brand"><div class="logo" id="logo"><svg viewBox="0 0 23 23" width="24" height="24" role="img" aria-label="Microsoft"><path fill="#F25022" d="M1 1h10v10H1z"/><path fill="#7FBA00" d="M12 1h10v10H12z"/><path fill="#00A4EF" d="M1 12h10v10H1z"/><path fill="#FFB900" d="M12 12h10v10H12z"/></svg></div><div><div class="b1">Azure Utilization &amp;<br>Cost Optimizer</div><div class="b2" id="bCustomer"></div></div></div>
    <nav id="nav"></nav>
    <div class="side-foot" id="sideFoot"></div>
  </aside>
  <main id="main">
    <header id="top">
      <button class="icon-btn" id="btnMenu" title="Collapse / expand the menu"></button>
      <div><h1 id="pageTitle"></h1><div class="sub" id="pageSub"></div></div>
      <div class="top-actions">
        <button class="search-btn" id="btnPalette" title="Search (Ctrl+K)"><i id="icoSearch"></i><span>Search resources, services, pages, rules&hellip;</span><kbd>Ctrl K</kbd></button>
        <button class="icon-btn" id="btnTheme" title="Toggle dark / light theme"></button>
        <button class="icon-btn" id="btnPrint" title="Print or save as PDF"></button>
        <div class="ms-brand" title="Microsoft"><svg viewBox="0 0 23 23" width="20" height="20" role="img" aria-label="Microsoft"><path fill="#F25022" d="M1 1h10v10H1z"/><path fill="#7FBA00" d="M12 1h10v10H12z"/><path fill="#00A4EF" d="M1 12h10v10H1z"/><path fill="#FFB900" d="M12 12h10v10H12z"/></svg><span>Microsoft</span></div>
      </div>
    </header>
    <section id="filters" class="no-print">
      <div class="fl" id="flbl"></div>
      <span id="fbtns" style="display:contents"></span>
      <div class="fsearch"><i id="icoFs"></i><input id="fq" placeholder="Search name, group, SKU, tag value&hellip;" autocomplete="off"></div>
      <span class="fcount" id="fcount"></span>
      <button class="lnk" id="freset" style="display:none">Reset all filters</button>
      <div id="fchips"></div>
    </section>
    <div id="page"></div>
    <footer id="foot"><div class="foot-l"><div class="foot-k">Report saved to</div><div class="foot-path"><span id="footPath"></span><button class="lnk" id="footCopy" title="Copy the file path">Copy path</button></div><div class="foot-meta" id="footMeta"></div></div><div class="ms-credit"><div class="ms-word"><svg viewBox="0 0 23 23" width="20" height="20" role="img" aria-label="Microsoft"><path fill="#F25022" d="M1 1h10v10H1z"/><path fill="#7FBA00" d="M12 1h10v10H12z"/><path fill="#00A4EF" d="M1 12h10v10H1z"/><path fill="#FFB900" d="M12 12h10v10H12z"/></svg><span>Microsoft</span></div><div class="foot-k">Solution Developed:</div><div class="cred-n">Zahir Hussain Shah</div><div class="cred-o">Microsoft Qatar</div></div></footer>
  </main>
</div>
<div id="drawerMask"></div>
<aside id="drawer" aria-label="Resource details"></aside>
<div id="palette"><div class="pal-box"><input id="palInput" placeholder="Search resources, services, pages and rules&hellip;" autocomplete="off"><div id="palList"></div></div></div>
<div id="dd"></div>
<div id="toast"></div>
<script id="dash-data" type="application/json">/*__DATA__*/</script>
<script>
(function () {
'use strict';
let D;
try { D = JSON.parse(document.getElementById('dash-data').textContent); }
catch (e) { document.body.innerHTML = '<div style="padding:40px;font-family:Segoe UI,sans-serif"><h2>Dashboard data could not be loaded</h2><p>' + String(e.message).replace(/</g, '&lt;') + '</p></div>'; return; }
window.D = D;
const M = D.meta, T = D.types || [], R = D.res || [], RC = D.recs || [], DATES = D.dates || [], TH = M.th || {};
const HAS_EC = typeof window.echarts !== 'undefined';
const CUR = (M.currency || 'USD') === 'USD' ? '$' : M.currency + ' ';

// ---------------------------------------------------------------- utilities
const $ = (s, el) => (el || document).querySelector(s);
const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
const esc = s => s == null ? '' : String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const isNum = v => typeof v === 'number' && isFinite(v);
const sum = (a, f) => { let s = 0; for (const x of a) { const v = f ? f(x) : x; if (isNum(v)) s += v; } return s; };
const avg = a => { const v = a.filter(isNum); return v.length ? v.reduce((p, c) => p + c, 0) / v.length : null; };
const pctl = (a, p) => { const v = a.filter(isNum).sort((x, y) => x - y); if (!v.length) return null; return v[Math.min(v.length - 1, Math.max(0, Math.ceil(p / 100 * v.length) - 1))]; };
const groupBy = (a, f) => { const m = new Map(); for (const x of a) { const k = f(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; };
const fmt = (v, d = 0) => isNum(v) ? v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }) : '–';
const fmtA = v => { if (!isNum(v)) return '–'; const a = Math.abs(v); return a >= 1000 ? fmt(v, 0) : a >= 100 ? fmt(v, 1) : a >= 1 ? fmt(v, 2) : a === 0 ? '0' : Number(v.toPrecision(2)).toString(); };
const money = v => isNum(v) ? (v < 0 ? '-' : '') + CUR + fmt(Math.abs(v), 0) : '–';
const moneyK = v => { if (!isNum(v)) return '–'; const a = Math.abs(v); const s = a >= 1e6 ? (a / 1e6).toFixed(a >= 1e7 ? 1 : 2) + 'M' : a >= 1e4 ? (a / 1e3).toFixed(a >= 1e5 ? 0 : 1) + 'K' : fmt(a, 0); return (v < 0 ? '-' : '') + CUR + s; };
const pctS = (v, d = 0) => isNum(v) ? fmt(v, d) + '%' : '–';
const debounce = (f, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => f(...a), ms); }; };
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
let _uid = 0; const uid = p => (p || 'u') + (++_uid);
const wavg = (a, vf, wf) => { let w = 0, sw = 0; const plain = []; for (const x of a) { const v = vf(x); if (!isNum(v)) continue; plain.push(v); const ww = wf ? wf(x) : 0; if (ww > 0) { w += v * ww; sw += ww; } } return sw > 0 ? w / sw : avg(plain); };
const lerpC = (a, b, t) => { const pa = a.match(/\w\w/g).map(h => parseInt(h, 16)), pb = b.match(/\w\w/g).map(h => parseInt(h, 16)); return '#' + pa.map((x, i) => Math.round(x + (pb[i] - x) * t).toString(16).padStart(2, '0')).join(''); };
const scoreColor = s => { if (!isNum(s)) return '#8A8886'; s = Math.max(0, Math.min(100, s)); return s < 50 ? lerpC('D13438', 'F7A30C', s / 50) : lerpC('F7A30C', '13A10E', (s - 50) / 50); };
const utilColor = u => !isNum(u) ? '#8A8886' : u < TH.idle ? '#D13438' : u < TH.low ? '#F7630C' : u > TH.high ? '#8764B8' : '#13A10E';
const linkify = s => esc(s).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
const DAYN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKEND = new Set(M.weekend || []);
const dateDay = d => DAYN[new Date(d + 'T12:00:00').getDay()];
const shortDate = d => d ? d.slice(5) : '';
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2200); }

const SV = p => '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + p + '</svg>';
const ICON = {
  home: SV('<path d="M3 11l9-7 9 7"/><path d="M5 10v10h14V10"/><path d="M10 20v-6h4v6"/>'),
  bulb: SV('<path d="M9 18h6"/><path d="M10 21h4"/><path d="M12 3a6 6 0 0 0-4 10.5c.8.8 1 1.5 1 2.5h6c0-1 .2-1.7 1-2.5A6 6 0 0 0 12 3z"/>'),
  grid: SV('<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>'),
  clock: SV('<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>'),
  shield: SV('<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6l8-3z"/><path d="M9 12l2 2 4-4"/>'),
  coin: SV('<circle cx="12" cy="12" r="9"/><path d="M15 9.5c-.5-1-1.6-1.5-3-1.5-1.7 0-3 .8-3 2s1.3 1.7 3 2 3 .8 3 2-1.3 2-3 2c-1.4 0-2.5-.5-3-1.5"/><path d="M12 6v2M12 16v2"/>'),
  tag: SV('<path d="M3 12V4h8l10 10-8 8L3 12z"/><circle cx="7.5" cy="8.5" r="1.5"/>'),
  trash: SV('<path d="M4 7h16"/><path d="M9 7V4h6v3"/><path d="M6 7l1 13h10l1-13"/>'),
  table: SV('<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>'),
  book: SV('<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2V5z"/><path d="M4 19a2 2 0 0 1 2-2h13"/>'),
  search: SV('<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>'),
  moon: SV('<path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/>'),
  sun: SV('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
  print: SV('<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>'),
  x: SV('<path d="M6 6l12 12M18 6L6 18"/>'),
  dl: SV('<path d="M12 4v11"/><path d="M7 10l5 5 5-5"/><path d="M4 20h16"/>'),
  filter: SV('<path d="M3 5h18l-7 8v6l-4-2v-4L3 5z"/>'),
  ext: SV('<path d="M14 4h6v6"/><path d="M20 4l-9 9"/><path d="M18 14v6H4V6h6"/>'),
  chev: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M6 9l6 6 6-6"/></svg>',
  menu: SV('<path d="M4 6h16M4 12h16M4 18h16"/>'),
  down: SV('<path d="M12 4v16"/><path d="M6 14l6 6 6-6"/>'),
  up: SV('<path d="M12 20V4"/><path d="M6 10l6-6 6 6"/>'),
  check: SV('<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>'),
  pause: SV('<circle cx="12" cy="12" r="9"/><path d="M10 9v6M14 9v6"/>'),
  bolt: SV('<path d="M13 3L5 14h6l-1 7 8-11h-6l1-7z"/>'),
  cpu: SV('<rect x="6" y="6" width="12" height="12" rx="2"/><path d="M9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4"/>'),
  db: SV('<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>'),
  chart: SV('<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>'),
  net: SV('<circle cx="12" cy="5" r="2.5"/><circle cx="5" cy="19" r="2.5"/><circle cx="19" cy="19" r="2.5"/><path d="M12 7.5v4M12 11.5l-6 5.5M12 11.5l6 5.5"/>'),
  box: SV('<path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/>'),
  eye: SV('<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>'),
  spark: SV('<path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5L18 18M6 18l2.5-2.5M15.5 8.5L18 6"/>'),
  layers: SV('<path d="M12 3l9 5-9 5-9-5 9-5z"/><path d="M3 13l9 5 9-5"/>'),
  logo: '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"><path d="M4 18a8 8 0 1 1 16 0"/><path d="M12 18l4-6"/><circle cx="12" cy="18" r="1.6" fill="#fff"/></svg>'
};
const CAT_ICON = { 'Compute': 'cpu', 'Containers': 'box', 'App Service': 'layers', 'Databases': 'db', 'Storage': 'db', 'Networking': 'net', 'Monitoring & Security': 'eye', 'Monitoring': 'eye', 'Integration': 'layers', 'Analytics': 'chart', 'AI & ML': 'spark', 'AI': 'spark' };

// ---------------------------------------------------------------- status / colors
const SGC = ['#D13438', '#F7630C', '#13A10E', '#8764B8', '#C19C00', '#8A8886'];
const SGK = ['idle', 'over', 'ok', 'under', 'stop', 'nd'];
const SG = (D.sg || []).map((g, i) => ({ n: g.n, s: g.s || [], c: SGC[i] || '#8A8886', k: SGK[i] || ('g' + i) }));
const stIdx = {}; SG.forEach((g, i) => g.s.forEach(s => { stIdx[s] = i; }));
const sgOf = st => stIdx[st] != null ? stIdx[st] : SG.length - 1;
const STC = st => (SG[sgOf(st)] || {}).c || '#8A8886';
const CATC = { 'Delete': '#D13438', 'Rightsize': '#F7630C', 'Configure': '#0078D4', 'Performance Risk': '#8764B8' };
const PRIC = { 'High': '#D13438', 'Medium': '#F7630C', 'Low': '#0078D4' };
const PAL = ['#0078D4', '#00B7C3', '#8764B8', '#F7630C', '#13A10E', '#C19C00', '#E3008C', '#4F6BED', '#038387', '#CA5010', '#498205', '#881798', '#69797E', '#B4009E', '#004E8C', '#986F0B'];
const pill = (txt, col) => txt ? '<span class="pill" style="--pc:' + col + '">' + esc(txt) + '</span>' : '';
const stPill = st => pill(st || 'Not assessed', STC(st));
const catPill = c => pill(c, CATC[c] || '#69797E');
const priPill = p => p ? pill(p, PRIC[p] || '#69797E') : '';
function ring(score, size, stroke) {
  size = size || 40; stroke = stroke || 5;
  if (!isNum(score)) return '<span class="ring na" style="width:' + size + 'px;height:' + size + 'px">n/a</span>';
  const r = (size - stroke) / 2, c = 2 * Math.PI * r, off = c * (1 - Math.max(0, Math.min(100, score)) / 100), h = size / 2;
  return '<svg class="ring" width="' + size + '" height="' + size + '" viewBox="0 0 ' + size + ' ' + size + '"><circle cx="' + h + '" cy="' + h + '" r="' + r + '" fill="none" stroke="currentColor" stroke-opacity=".14" stroke-width="' + stroke + '"/><circle cx="' + h + '" cy="' + h + '" r="' + r + '" fill="none" stroke="' + scoreColor(score) + '" stroke-width="' + stroke + '" stroke-dasharray="' + c.toFixed(2) + '" stroke-dashoffset="' + off.toFixed(2) + '" stroke-linecap="round" transform="rotate(-90 ' + h + ' ' + h + ')"/><text x="50%" y="50%" dominant-baseline="central" text-anchor="middle" font-size="' + Math.round(size * 0.3) + '" font-weight="700" fill="currentColor">' + Math.round(score) + '</text></svg>';
}
function ubar(u) { if (!isNum(u)) return '<span class="muted">–</span>'; return '<span class="nowrap"><span class="ubar"><i style="width:' + Math.max(2, Math.min(100, u)) + '%;background:' + utilColor(u) + '"></i></span>' + fmt(u, 1) + '%</span>'; }
function spark(arr, color, w, h) {
  w = w || 92; h = h || 24;
  if (!arr || !arr.some(isNum)) return '<span class="muted">–</span>';
  const nums = arr.filter(isNum), mn = Math.min.apply(null, nums), mx = Math.max.apply(null, nums), rng = (mx - mn) || 1, step = w / Math.max(1, arr.length - 1);
  let d = '', pen = false;
  arr.forEach((x, i) => { if (!isNum(x)) { pen = false; return; } d += (pen ? 'L' : 'M') + (i * step).toFixed(1) + ' ' + (h - 3 - (x - mn) / rng * (h - 6)).toFixed(1); pen = true; });
  return '<svg class="spark" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h + '"><path d="' + d + '" fill="none" stroke="' + (color || '#0078D4') + '" stroke-width="1.7" stroke-linejoin="round" stroke-linecap="round"/></svg>';
}
const kpi = (label, value, sub, color, icon, tip) => '<div class="kpi" style="--kc:' + (color || '#0078D4') + '"' + (tip ? ' title="' + esc(tip) + '"' : '') + '><div class="kpi-i">' + (ICON[icon] || ICON.chart) + '</div><div><div class="kpi-l">' + esc(label) + '</div><div class="kpi-v">' + value + '</div><div class="kpi-s">' + (sub || '') + '</div></div></div>';
const card = (title, body, o) => { o = o || {}; return '<section class="card ' + (o.cls || '') + '"' + (o.id ? ' id="' + o.id + '"' : '') + (o.style ? ' style="' + o.style + '"' : '') + '><header class="card-h"><h3>' + esc(title) + (o.info ? '<span class="info" title="' + esc(o.info) + '">i</span>' : '') + '</h3><div class="card-a">' + (o.actions || '') + '</div></header><div class="card-b">' + body + '</div></section>'; };
const chartBox = (id, h) => '<div class="chart" id="' + id + '" style="height:' + (h || 300) + 'px"></div>';
const emptyMsg = t => '<div class="empty">' + esc(t || 'No data for the current filters.') + '</div>';

// ---------------------------------------------------------------- data preparation
T.forEach((t, i) => {
  t.i = i; t.res = []; t.mIx = {}; t.fIx = {}; t.pIx = {};
  (t.mc || []).forEach((c, j) => { t.mIx[c] = j; }); (t.fc || []).forEach((c, j) => { t.fIx[c] = j; }); (t.pc || []).forEach((c, j) => { t.pIx[c] = j; });
  t.ml = t.ml || {}; t.mu = t.mu || {}; t.hw = t.hw || {};
});
const tByKey = {}; T.forEach(t => { tByKey[t.k] = t; });
R.forEach((r, i) => {
  r.i = i; r.T = T[r.t] || { n: '?', c: 'Other', mIx: {}, fIx: {}, pIx: {}, res: [], ml: {}, mu: {}, hw: {} }; r.T.res.push(r);
  r.cat = r.T.c; r.tn = r.T.n; r.sub = D.subs[r.s] || ''; r.e = r.e || 'Unclassified';
  r.tc = r.x ? 0 : (r.c || 0) + (r.cc || 0); r.sgi = sgOf(r.st);
  r.rcs = r.rc ? r.rc.split(',').map(s => s.trim()).filter(Boolean) : [];
  r.tg = r.tg || {}; r.tgl = {}; Object.keys(r.tg).forEach(k => { r.tgl[k.toLowerCase()] = r.tg[k]; });
  r.recs = [];
  r.hay = [r.n, r.g, r.tn, r.sub, r.sku, r.l, r.k, r.id, r.e].concat(Object.keys(r.tg).map(k => k + '=' + r.tg[k])).join(' ').toLowerCase();
});
RC.forEach((x, i) => { x.i = i; x.R = x.r >= 0 ? R[x.r] : null; if (x.R) x.R.recs.push(x); x.tn = T[x.t] ? T[x.t].n : ''; x.sub = D.subs[x.s] || ''; x.rule = (D.rules || {})[x.rc] || {}; x.e = x.e || 'Unclassified'; });
const CATORD = { 'Delete': 0, 'Rightsize': 1, 'Configure': 2, 'Performance Risk': 3 };
R.forEach(r => r.recs.sort((a, b) => (CATORD[a.cat] - CATORD[b.cat]) || ((b.sa || 0) - (a.sa || 0))));
const TAGK = (() => { const m = new Map(); R.forEach(r => Object.keys(r.tg).forEach(k => { const lk = k.toLowerCase(); const e = m.get(lk) || { k: lk, d: k, n: 0 }; e.n++; m.set(lk, e); })); return Array.from(m.values()).sort((a, b) => b.n - a.n); })();
const tagDisp = lk => (TAGK.find(t => t.k === lk) || { d: lk }).d;
function fv(r, col) { const t = r.T; let j = t.fIx[col]; if (j != null) return r.f ? r.f[j] : null; j = t.mIx[col]; if (j != null) return r.m ? r.m[j] : null; j = t.pIx[col]; if (j != null) return r.p ? r.p[j] : null; return null; }
const AGGN = { Avg: 'avg', P95: 'P95', P99: 'P99', Med: 'median', Max: 'max', Min: 'min', MinAvg: 'lowest hourly avg', P05: 'P05', Total: 'total', MaxT: 'peak-hour total', PeakHour: 'peak hour' };
function mSplit(col) { const m = /^(.*)_(Avg|P95|P99|Med|Max|Min|MinAvg|P05|Total|MaxT|PeakHour)$/.exec(col); return m ? { a: m[1], g: m[2] } : { a: col, g: '' }; }
function mLabel(t, col) { const s = mSplit(col); const lbl = (t && t.ml && t.ml[s.a]) || s.a; return lbl + (s.g ? ' · ' + (AGGN[s.g] || s.g) : ''); }
function wscore(rs) { let w = 0, sw = 0, n = 0, s = 0; for (const r of rs) { if (!isNum(r.es)) continue; s += r.es; n++; if (r.tc > 0) { w += r.es * r.tc; sw += r.tc; } } return sw > 0 ? w / sw : n ? s / n : null; }
const isQW = x => (x.cat === 'Delete' || x.cat === 'Configure') && x.eff === 'Low' && x.conf !== 'Low' && x.sa > 0;
function savSplit(rs) {
  const o = { Delete: 0, Rightsize: 0, Configure: 0 };
  rs.forEach(r => {
    if (!r.sv) return; const recs = r.recs.filter(x => x.cat !== 'Performance Risk');
    if (recs.some(x => x.cat === 'Delete')) { o.Delete += r.sv; return; }
    const a = sum(recs.filter(x => x.cat === 'Rightsize'), x => x.sm), b = sum(recs.filter(x => x.cat === 'Configure'), x => x.sm), t = a + b;
    if (t > 0) { o.Rightsize += r.sv * a / t; o.Configure += r.sv * b / t; } else o.Configure += r.sv;
  });
  return o;
}
function dedupSav(recs) {
  const out = { Delete: 0, Rightsize: 0, Configure: 0, total: 0 };
  const tk = x => (T[x.t] || {}).k;
  const parentDel = new Set(recs.filter(x => x.cat === 'Delete' && (tk(x) === 'StorageAccounts' || tk(x) === 'CognitiveServices') && x.rid).map(x => x.rid.toLowerCase()));
  groupBy(recs.filter(x => x.cat !== 'Performance Risk' && isNum(x.sm)), x => x.R ? 'r' + x.R.i : 'n' + x.t + '|' + x.n).forEach(list => {
    const f = list[0];
    if ((tk(f) === 'PremiumFileShares' || tk(f) === 'AIModelDeployments') && f.rid && parentDel.has(f.rid.toLowerCase())) return;
    const cap = list.some(x => x.rc === 'DEALLOC') ? null : (f.R ? f.R.c : f.cmc); const del = list.filter(x => x.cat === 'Delete');
    if (del.length) { let v = Math.max.apply(null, del.map(x => x.sm)); if (isNum(cap) && cap > 0) v = Math.min(v, cap); out.Delete += v; out.total += v; return; }
    let a = sum(list.filter(x => x.cat === 'Rightsize'), x => x.sm), b = sum(list.filter(x => x.cat === 'Configure'), x => x.sm);
    if (isNum(cap) && cap > 0 && a + b > cap) { const f = cap / (a + b); a *= f; b *= f; }
    out.Rightsize += a; out.Configure += b; out.total += a + b;
  });
  return out;
}

// ---------------------------------------------------------------- global filters
const DIMS = [
  { k: 'sub', n: 'Subscription', get: r => r.sub },
  { k: 'rg', n: 'Resource group', get: r => r.g || '(none)' },
  { k: 'loc', n: 'Location', get: r => r.l || '(none)' },
  { k: 'env', n: 'Environment', get: r => r.e },
  { k: 'cat', n: 'Category', get: r => r.cat },
  { k: 'type', n: 'Resource type', get: r => r.tn },
  { k: 'sg', n: 'Utilization status', get: r => (SG[r.sgi] || {}).n },
  { k: 'act', n: 'Action', multi: true, get: r => r.rcs.length ? r.rcs : ['No action'] },
  { k: 'pri', n: 'Priority', get: r => r.rp || 'None' }
];
const F = { q: '', tagKey: '', tagVals: new Set() }; DIMS.forEach(d => { F[d.k] = new Set(); });
let FR = null, FRC = null;
const qToks = () => F.q.toLowerCase().split(/\s+/).filter(Boolean);
function passDim(r, d) { const s = F[d.k]; if (!s.size) return true; const v = d.get(r); return d.multi ? v.some(x => s.has(x)) : s.has(v); }
function passTag(r) { if (!F.tagKey || !F.tagVals.size) return true; const v = r.tgl[F.tagKey]; return F.tagVals.has(v == null || v === '' ? '(not tagged)' : v); }
function passQ(r) { if (!F.q) return true; const h = r.hay; return qToks().every(t => h.indexOf(t) >= 0); }
function pass(r, skip) { for (const d of DIMS) { if (d.k !== skip && !passDim(r, d)) return false; } if (skip !== 'tag' && !passTag(r)) return false; return passQ(r); }
function fr() { if (!FR) FR = R.filter(r => pass(r)); return FR; }
function frecs() {
  if (!FRC) FRC = RC.filter(x => {
    if (F.act.size && !F.act.has(x.cat)) return false;
    if (F.pri.size && !F.pri.has(x.pr)) return false;
    if (x.R) { for (const d of DIMS) { if (d.k === 'act' || d.k === 'pri') continue; if (!passDim(x.R, d)) return false; } return passTag(x.R) && passQ(x.R); }
    return (!F.sub.size || F.sub.has(x.sub)) && (!F.rg.size || F.rg.has(x.g)) && (!F.type.size || F.type.has(x.tn)) && (!F.env.size || F.env.has(x.e));
  });
  return FRC;
}
function passSR(sub, rg, text) { if (F.sub.size && !F.sub.has(sub)) return false; if (F.rg.size && !F.rg.has(rg || '(none)')) return false; if (F.q) { const h = String(text || '').toLowerCase(); if (!qToks().every(t => h.indexOf(t) >= 0)) return false; } return true; }
const anyFilter = () => !!F.q || !!(F.tagKey && F.tagVals.size) || DIMS.some(d => F[d.k].size);
function applyFilters(keepScroll) { FR = FRC = null; renderFilterBar(); const y = window.scrollY; route(true); if (keepScroll !== false) window.scrollTo(0, y); }
function setFilter(k, vals) { if (k === 'tag') return; F[k] = new Set(vals); applyFilters(); }
function setTag(key, val) { F.tagKey = key.toLowerCase(); F.tagVals = new Set([val]); applyFilters(); }

function renderFilterBar() {
  $('#fbtns').innerHTML = DIMS.map(d => { const n = F[d.k].size; return '<button class="fbtn' + (n ? ' act' : '') + '" data-dim="' + d.k + '">' + esc(d.n) + (n ? ' <span class="n">' + n + '</span>' : '') + ICON.chev + '</button>'; }).join('') +
    '<button class="fbtn' + (F.tagKey && F.tagVals.size ? ' act' : '') + '" data-dim="tag">Tag' + (F.tagKey ? ': ' + esc(tagDisp(F.tagKey)) : '') + (F.tagKey && F.tagVals.size ? ' <span class="n">' + F.tagVals.size + '</span>' : '') + ICON.chev + '</button>';
  const n = fr().length;
  $('#fcount').innerHTML = '<b>' + fmt(n) + '</b> of ' + fmt(R.length) + ' resources';
  $('#freset').style.display = anyFilter() ? '' : 'none';
  const chips = [];
  DIMS.forEach(d => F[d.k].forEach(v => chips.push('<span class="chip"><b>' + esc(d.n) + ':</b> ' + esc(v) + '<button data-cd="' + d.k + '" data-cv="' + esc(v) + '" title="Remove">&times;</button></span>')));
  if (F.tagKey) F.tagVals.forEach(v => chips.push('<span class="chip"><b>Tag ' + esc(tagDisp(F.tagKey)) + ':</b> ' + esc(v) + '<button data-cd="tag" data-cv="' + esc(v) + '" title="Remove">&times;</button></span>'));
  if (F.q) chips.push('<span class="chip"><b>Search:</b> ' + esc(F.q) + '<button data-cd="q" title="Remove">&times;</button></span>');
  $('#fchips').innerHTML = chips.join('');
  updateBadges();
}
function openDD(dimKey) {
  const btn = $('.fbtn[data-dim="' + dimKey + '"]'); if (!btn) return;
  const dd = $('#dd'); const keepQ = dd.dataset.dim === dimKey && $('#ddq') ? $('#ddq').value : ''; const keepScroll = dd.dataset.dim === dimKey && $('.dd-l', dd) ? $('.dd-l', dd).scrollTop : 0;
  let opts, head = '', sel;
  if (dimKey === 'tag') {
    head = '<select id="ddTagKey"><option value="">Choose a tag key…</option>' + TAGK.map(t => '<option value="' + esc(t.k) + '"' + (t.k === F.tagKey ? ' selected' : '') + '>' + esc(t.d) + ' (' + fmt(t.n) + ')</option>').join('') + '</select>';
    sel = F.tagVals; const counts = new Map();
    if (F.tagKey) { for (const r of R) { if (!pass(r, 'tag')) continue; const v = r.tgl[F.tagKey]; const k = v == null || v === '' ? '(not tagged)' : v; counts.set(k, (counts.get(k) || 0) + 1); } }
    sel.forEach(v => { if (!counts.has(v)) counts.set(v, 0); });
    opts = Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
  } else {
    const d = DIMS.find(x => x.k === dimKey); sel = F[d.k]; const counts = new Map();
    for (const r of R) { if (!pass(r, d.k)) continue; const v = d.get(r); (d.multi ? v : [v]).forEach(x => counts.set(x, (counts.get(x) || 0) + 1)); }
    sel.forEach(v => { if (!counts.has(v)) counts.set(v, 0); });
    opts = Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
    if (d.k === 'sg') opts.sort((a, b) => SG.findIndex(g => g.n === a[0]) - SG.findIndex(g => g.n === b[0]));
    if (d.k === 'pri') opts.sort((a, b) => ['High', 'Medium', 'Low', 'None'].indexOf(a[0]) - ['High', 'Medium', 'Low', 'None'].indexOf(b[0]));
  }
  const dot = v => dimKey === 'sg' ? '<i class="dot" style="background:' + ((SG.find(g => g.n === v) || {}).c || '#8A8886') + '"></i>' : dimKey === 'act' ? '<i class="dot" style="background:' + (CATC[v] || '#8A8886') + '"></i>' : dimKey === 'pri' ? '<i class="dot" style="background:' + (PRIC[v] || '#8A8886') + '"></i>' : '';
  dd.dataset.dim = dimKey;
  dd.innerHTML = '<div class="dd-h">' + head + '<input id="ddq" placeholder="Filter values…" value="' + esc(keepQ) + '"><div class="dd-a"><button class="lnk" data-a="all">Select all shown</button><button class="lnk" data-a="none">Clear</button></div></div><div class="dd-l">' +
    (opts.length ? opts.map(o => '<label class="dd-o" data-v="' + esc(o[0]) + '"><input type="checkbox"' + (sel.has(o[0]) ? ' checked' : '') + '><span class="dd-t">' + dot(o[0]) + esc(o[0]) + '</span><span class="dd-c">' + fmt(o[1]) + '</span></label>').join('') : '<div class="empty">' + (dimKey === 'tag' && !F.tagKey ? 'Choose a tag key first' : 'No values') + '</div>') + '</div>';
  const b = btn.getBoundingClientRect(); dd.classList.add('open');
  dd.style.left = Math.max(8, Math.min(window.innerWidth - 340, b.left + window.scrollX)) + 'px'; dd.style.top = (b.bottom + window.scrollY + 6) + 'px';
  const q = $('#ddq'); const filt = () => { const t = q.value.toLowerCase(); $$('.dd-o', dd).forEach(o => { o.style.display = o.dataset.v.toLowerCase().indexOf(t) >= 0 ? '' : 'none'; }); };
  q.addEventListener('input', filt); filt(); $('.dd-l', dd).scrollTop = keepScroll;
  if (!keepQ) setTimeout(() => q.focus(), 10);
  const tk = $('#ddTagKey'); if (tk) tk.addEventListener('change', () => { F.tagKey = tk.value; F.tagVals = new Set(); applyFilters(); openDD('tag'); });
  dd.onclick = e => {
    const a = e.target.closest('[data-a]');
    if (a) { if (a.dataset.a === 'none') sel.clear(); else $$('.dd-o', dd).filter(o => o.style.display !== 'none').forEach(o => sel.add(o.dataset.v)); applyFilters(); openDD(dimKey); return; }
    const o = e.target.closest('.dd-o'); if (!o) return; e.preventDefault();
    const v = o.dataset.v; sel.has(v) ? sel.delete(v) : sel.add(v); applyFilters(); openDD(dimKey);
  };
}
function closeDD() { const dd = $('#dd'); dd.classList.remove('open'); dd.dataset.dim = ''; }
document.addEventListener('mousedown', e => { if (!e.target.closest('#dd') && !e.target.closest('.fbtn')) closeDD(); });

// ---------------------------------------------------------------- charts
let CH = [];
function disposeCharts() { CH.forEach(c => { try { c.dispose(); } catch (e) { /* ignore */ } }); CH = []; }
function themed(o) {
  const txt = css('--text'), mut = css('--muted'), line = css('--line'), cardc = css('--card');
  o = Object.assign({ backgroundColor: 'transparent', animationDuration: 650, color: PAL }, o);
  o.textStyle = Object.assign({ fontFamily: 'Segoe UI, system-ui, sans-serif', color: txt }, o.textStyle || {});
  if (o.tooltip !== false) o.tooltip = Object.assign({ confine: true, backgroundColor: cardc, borderColor: line, textStyle: { color: txt, fontSize: 12 }, extraCssText: 'box-shadow:0 8px 28px rgba(0,0,0,.18);border-radius:10px;' }, o.tooltip || {});
  if (o.legend) [].concat(o.legend).forEach(l => { l.textStyle = Object.assign({ color: mut }, l.textStyle || {}); if (l.type == null) l.type = 'scroll'; l.pageTextStyle = { color: mut }; l.pageIconColor = mut; });
  ['xAxis', 'yAxis'].forEach(k => { if (!o[k]) return; [].concat(o[k]).forEach(a => { a.axisLabel = Object.assign({ color: mut }, a.axisLabel || {}); a.axisLine = Object.assign({ lineStyle: { color: line } }, a.axisLine || {}); a.splitLine = Object.assign({ lineStyle: { color: line, type: 'dashed' } }, a.splitLine || {}); a.nameTextStyle = Object.assign({ color: mut }, a.nameTextStyle || {}); a.axisTick = Object.assign({ show: false }, a.axisTick || {}); }); });
  if (o.title) [].concat(o.title).forEach(t => { t.textStyle = Object.assign({ color: txt, fontSize: 13 }, t.textStyle || {}); t.subtextStyle = Object.assign({ color: mut }, t.subtextStyle || {}); });
  if (o.visualMap) [].concat(o.visualMap).forEach(v => { v.textStyle = Object.assign({ color: mut }, v.textStyle || {}); });
  return o;
}
function ec(id, opt, onClick, store) {
  const el = typeof id === 'string' ? document.getElementById(id) : id; if (!el) return null;
  if (!HAS_EC) { el.innerHTML = '<div class="nochart">Charts require the ECharts library (not embedded and the CDN is unreachable).</div>'; return null; }
  const c = echarts.init(el, null, { renderer: 'canvas' });
  c.setOption(themed(opt)); if (onClick) c.on('click', onClick); (store || CH).push(c); return c;
}
window.addEventListener('resize', debounce(() => { CH.concat(DRCH).forEach(c => { try { c.resize(); } catch (e) { /* ignore */ } }); }, 120));
const axisMoney = { formatter: v => moneyK(v) };
const hbarOpt = (labels, values, colors, fmtFn, extra) => Object.assign({
  grid: { left: 8, right: 70, top: 10, bottom: 10, containLabel: true },
  xAxis: { type: 'value', axisLabel: { formatter: v => fmtFn ? fmtFn(v) : fmtA(v), hideOverlap: true } },
  yAxis: { type: 'category', data: labels, inverse: true, axisLabel: { width: 190, overflow: 'truncate', interval: 0 } },
  tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => fmtFn ? fmtFn(v) : fmtA(v) },
  series: [{ type: 'bar', data: values.map((v, i) => ({ value: v, itemStyle: { color: Array.isArray(colors) ? colors[i % colors.length] : (colors || PAL[0]), borderRadius: [0, 6, 6, 0] } })), barMaxWidth: 20, label: { show: true, position: 'right', formatter: p => fmtFn ? fmtFn(p.value) : fmtA(p.value), color: css('--muted'), fontSize: 11 } }]
}, extra || {});
const donutOpt = (data, centerTop, centerSub, fmtFn) => ({
  tooltip: { trigger: 'item', formatter: p => '<b>' + esc(p.name) + '</b><br>' + (fmtFn ? fmtFn(p.value) : fmt(p.value)) + ' (' + p.percent + '%)' },
  legend: { bottom: 0, icon: 'circle', itemWidth: 9, itemHeight: 9 },
  title: centerTop != null ? { text: centerTop, subtext: centerSub || '', left: 'center', top: '34%', textStyle: { fontSize: 20, fontWeight: 750 } } : undefined,
  series: [{ type: 'pie', radius: ['50%', '74%'], center: ['50%', '44%'], avoidLabelOverlap: true, itemStyle: { borderColor: css('--card'), borderWidth: 3, borderRadius: 6 }, label: { show: false }, emphasis: { scale: true, scaleSize: 6 }, data }]
});

// ---------------------------------------------------------------- data grid
function exportCsv(name, cols, rows) {
  const q = v => { if (v == null) return ''; const s = Array.isArray(v) ? v.join(' ') : String(v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const lines = [cols.map(c => q(c.n)).join(',')].concat(rows.map(r => cols.map(c => q(c.v ? c.v(r) : r[c.k])).join(',')));
  const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = (name || 'export').replace(/[\\/:*?"<>|]/g, '-') + '.csv';
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 800);
  toast('Exported ' + fmt(rows.length) + ' rows');
}
function DT(host, cfg) {
  if (typeof host === 'string') host = document.getElementById(host);
  if (!host) return null;
  const st = { q: cfg.q || '', page: 0, ps: cfg.pageSize || 25, sk: cfg.sortK || null, sd: cfg.sortD || -1, rows: cfg.rows || [] };
  const id = uid('dt');
  host.innerHTML = '<div class="dt" id="' + id + '"><div class="dt-bar">' + (cfg.search !== false ? '<div class="dt-s">' + ICON.search + '<input placeholder="Search in this table…" value="' + esc(st.q) + '"></div>' : '') + '<span class="dt-n"></span><span class="sp"></span>' + (cfg.tools || '') + (cfg.exportName ? '<button class="btn sm dt-x">' + ICON.dl + ' Export CSV</button>' : '') + '</div><div class="tw"' + (cfg.maxH ? ' style="max-height:' + cfg.maxH + 'px"' : '') + '><table class="t"><thead></thead><tbody></tbody></table></div><div class="dt-p"></div></div>';
  const root = document.getElementById(id), thead = $('thead', root), tbody = $('tbody', root);
  const getv = (c, r) => c.v ? c.v(r) : r[c.k];
  function view() {
    let rows = st.rows;
    if (st.q) { const toks = st.q.toLowerCase().split(/\s+/).filter(Boolean); rows = rows.filter(r => { const h = cfg.cols.map(c => { const v = getv(c, r); return v == null ? '' : String(v); }).join(' ').toLowerCase(); return toks.every(t => h.indexOf(t) >= 0); }); }
    if (st.sk) { const c = cfg.cols.find(x => x.k === st.sk); if (c) { rows = rows.slice().sort((a, b) => { const x = getv(c, a), y = getv(c, b); const xn = x == null || x === '', yn = y == null || y === ''; if (xn && yn) return 0; if (xn) return 1; if (yn) return -1; return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true })) * st.sd; }); } }
    return rows;
  }
  function draw() {
    const rows = view(), pages = Math.max(1, Math.ceil(rows.length / st.ps)); if (st.page >= pages) st.page = pages - 1;
    const slice = rows.slice(st.page * st.ps, (st.page + 1) * st.ps);
    thead.innerHTML = '<tr>' + (cfg.sel ? '<th class="ck"><input type="checkbox" class="ck-all" title="Select all rows in the table"></th>' : '') + cfg.cols.map(c => '<th class="' + (c.num ? 'num ' : '') + (c.cls || '') + '" data-k="' + esc(c.k) + '"' + (c.tip ? ' title="' + esc(c.tip) + '"' : '') + '>' + esc(c.n) + (st.sk === c.k ? (st.sd > 0 ? ' ▲' : ' ▼') : '') + '</th>').join('') + '</tr>';
    tbody.innerHTML = slice.length ? slice.map((r, i) => '<tr data-i="' + (st.page * st.ps + i) + '"' + (cfg.onRow ? ' class="clk"' : '') + '>' + (cfg.sel ? '<td class="ck"><input type="checkbox"' + (cfg.selSet.has(r) ? ' checked' : '') + '></td>' : '') + cfg.cols.map(c => '<td class="' + (c.num ? 'num ' : '') + (c.cls || '') + '">' + (c.f ? c.f(r) : esc(getv(c, r))) + '</td>').join('') + '</tr>').join('')
      : '<tr><td colspan="' + (cfg.cols.length + (cfg.sel ? 1 : 0)) + '" class="empty">' + esc(cfg.emptyText || 'No rows match the current filters.') + '</td></tr>';
    $('.dt-n', root).textContent = fmt(rows.length) + ' row' + (rows.length === 1 ? '' : 's');
    $('.dt-p', root).innerHTML = rows.length > 25 ? '<button class="btn sm" data-p="-1"' + (st.page ? '' : ' disabled') + '>‹ Prev</button><span>Page ' + (st.page + 1) + ' of ' + pages + '</span><button class="btn sm" data-p="1"' + (st.page < pages - 1 ? '' : ' disabled') + '>Next ›</button><select class="sel" data-ps>' + [25, 50, 100, 250, 1000].map(n => '<option' + (n === st.ps ? ' selected' : '') + '>' + n + '</option>').join('') + '</select><span>rows / page</span>' : '';
    root._rows = rows;
  }
  root.addEventListener('click', e => {
    if (e.target.matches('.ck-all')) { const on = e.target.checked; view().forEach(r => { if (on) cfg.selSet.add(r); else cfg.selSet.delete(r); }); draw(); if (cfg.onSel) cfg.onSel(); return; }
    const th = e.target.closest('th[data-k]');
    if (th) { const k = th.dataset.k; if (st.sk === k) st.sd = -st.sd; else { st.sk = k; const c = cfg.cols.find(x => x.k === k); st.sd = c && c.num ? -1 : 1; } draw(); return; }
    const pb = e.target.closest('[data-p]'); if (pb) { st.page += +pb.dataset.p; draw(); return; }
    if (e.target.closest('.dt-x')) { exportCsv(cfg.exportName, cfg.cols.filter(c => c.x !== false), view()); return; }
    const tr = e.target.closest('tbody tr[data-i]'); if (!tr) return;
    const row = root._rows[+tr.dataset.i]; if (!row) return;
    if (e.target.closest('td.ck')) { const cb = tr.querySelector('input'); if (e.target !== cb) cb.checked = !cb.checked; if (cb.checked) cfg.selSet.add(row); else cfg.selSet.delete(row); if (cfg.onSel) cfg.onSel(); return; }
    if (e.target.closest('a,button,select,input')) return;
    if (cfg.onRow) cfg.onRow(row);
  });
  root.addEventListener('change', e => { if (e.target.matches('[data-ps]')) { st.ps = +e.target.value; st.page = 0; draw(); } });
  const si = $('.dt-s input', root); if (si) si.addEventListener('input', debounce(() => { st.q = si.value; st.page = 0; draw(); }, 180));
  draw();
  return { set(rows) { st.rows = rows; st.page = 0; draw(); }, draw, view, st };
}
const RCOL = {
  name: { k: 'n', n: 'Resource', f: r => '<div class="rn"><b>' + esc(r.n) + '</b><small>' + esc(r.tn) + '</small></div>', v: r => r.n },
  type: { k: 'tn', n: 'Type', v: r => r.tn },
  rg: { k: 'g', n: 'Resource group', v: r => r.g },
  sub: { k: 'sub', n: 'Subscription', v: r => r.sub },
  loc: { k: 'l', n: 'Location', v: r => r.l },
  env: { k: 'e', n: 'Environment', v: r => r.e },
  sku: { k: 'sku', n: 'SKU / size', v: r => r.sku },
  st: { k: 'st', n: 'Status', f: r => stPill(r.st), v: r => r.st },
  u: { k: 'u', n: 'Utilization', num: true, f: r => ubar(r.u), v: r => r.u, tip: 'Primary utilization of the provisioned SKU (P95 %)' },
  es: { k: 'es', n: 'Score', num: true, f: r => ring(r.es, 30, 4), v: r => r.es, tip: 'Efficiency score 0-100' },
  c: { k: 'c', n: 'Monthly cost', num: true, f: r => money(r.c), v: r => r.c },
  sv: { k: 'sv', n: 'Savings / yr', num: true, f: r => r.sv ? '<b class="pos">' + money(r.sv) + '</b>' : '<span class="muted">–</span>', v: r => r.sv },
  rec: { k: 'rec', n: 'Suggested action', f: r => r.recs.length ? '<div class="clip" title="' + esc(r.recs.map(x => '[' + x.cat + '] ' + x.act).join('\n')) + '">' + catPill(r.recs[0].cat) + ' ' + esc(r.recs[0].act) + (r.recs.length > 1 ? ' <span class="muted">+' + (r.recs.length - 1) + '</span>' : '') + '</div>' : '<span class="muted">No action</span>', v: r => r.recs.map(x => '[' + x.cat + '] ' + x.act).join(' | ') },
  trend: { k: 'trend', n: 'Daily trend', f: r => spark(r.d, STC(r.st)), v: r => r.d ? r.d.map(x => x == null ? '' : x).join(' ') : '', x: false },
  pv: { k: 'pv', n: 'Provisioned', f: r => '<div class="clip" title="' + esc(r.pv) + '">' + esc(r.pv) + '</div>', v: r => r.pv },
  us: { k: 'us', n: 'Utilized', f: r => '<div class="clip" title="' + esc(r.us) + '">' + esc(r.us) + '</div>', v: r => r.us },
  age: { k: 'ag', n: 'Age (days)', num: true, v: r => r.ag },
  cr: { k: 'cr', n: 'Created', v: r => r.cr },
  pri: { k: 'rp', n: 'Priority', f: r => priPill(r.rp), v: r => r.rp },
  id: { k: 'id', n: 'Resource ID', f: r => '<div class="clip muted" style="max-width:260px" title="' + esc(r.id) + '">' + esc(r.id) + '</div>', v: r => r.id }
};
const resCols = keys => keys.map(k => RCOL[k]);

// ---------------------------------------------------------------- router & navigation
const PAGES = [
  { k: 'overview', n: 'Executive overview', ic: 'home', r: renderOverview },
  { k: 'recs', n: 'Recommendations', ic: 'bulb', r: renderRecs, badge: () => frecs().length },
  { k: 'services', n: 'Service dashboards', ic: 'grid', r: renderServices, badge: () => new Set(fr().map(r => r.t)).size },
  { k: 'patterns', n: 'Usage patterns', ic: 'clock', r: renderPatterns },
  { k: 'commit', n: 'Commitments & Log Analytics', ic: 'shield', r: renderCommit },
  { k: 'cost', n: 'Cost analysis', ic: 'coin', r: renderCost },
  { k: 'tags', n: 'Tags & environments', ic: 'tag', r: renderTags },
  { k: 'waste', n: 'Idle & orphaned', ic: 'trash', r: renderWaste, badge: () => fr().filter(r => r.sgi === 0).length },
  { k: 'explorer', n: 'Resource explorer', ic: 'table', r: renderExplorer, badge: () => fr().length },
  { k: 'method', n: 'Methodology & rules', ic: 'book', r: renderMethod }
];
function updateBadges() { PAGES.forEach(p => { if (!p.badge) return; const b = $('.bdg[data-b="' + p.k + '"]'); if (b) b.textContent = fmt(p.badge()); }); }
let CUR_PAGE = '', CUR_ARG = '';
function route(refresh) {
  const h = (location.hash || '').replace(/^#\/?/, '').split('/');
  const p = PAGES.find(x => x.k === h[0]) || PAGES[0];
  let arg = ''; try { arg = decodeURIComponent(h.slice(1).join('/')); } catch (e) { arg = h.slice(1).join('/'); }
  const changed = p.k !== CUR_PAGE || arg !== CUR_ARG;
  CUR_PAGE = p.k; CUR_ARG = arg;
  $$('.nav-i').forEach(a => a.classList.toggle('on', a.dataset.k === p.k));
  $('#pageTitle').textContent = p.n;
  $('#pageSub').textContent = M.customer + ' · ' + M.start + ' → ' + M.end + ' (' + M.days + ' days, ' + M.tz + ')';
  disposeCharts(); closeDD();
  const el = $('#page'); el.className = refresh ? '' : 'fade-in'; el.innerHTML = '';
  try { p.r(el, arg); }
  catch (e) { console.error(e); el.innerHTML = card('This view could not be rendered', '<p>' + esc(e.message) + '</p><pre class="sm muted" style="white-space:pre-wrap">' + esc(e.stack || '') + '</pre>'); }
  if (changed && !refresh) window.scrollTo(0, 0);
}
window.addEventListener('hashchange', () => route(false));
const go = h => { if (location.hash === '#/' + h) route(false); else location.hash = '#/' + h; };

// ---------------------------------------------------------------- resource drawer
let DRCH = [];
function disposeDrawerCharts() { DRCH.forEach(c => { try { c.dispose(); } catch (e) { /* ignore */ } }); DRCH = []; }
function closeDrawer() { $('#drawer').classList.remove('open'); $('#drawerMask').classList.remove('open'); disposeDrawerCharts(); }
function copyText(t) {
  const done = () => toast('Copied to clipboard');
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(done, () => fallback());
  else fallback();
  function fallback() { const a = document.createElement('textarea'); a.value = t; document.body.appendChild(a); a.select(); try { document.execCommand('copy'); done(); } catch (e) { toast('Copy failed'); } a.remove(); }
}
const weekendArea = () => ({ silent: true, itemStyle: { color: 'rgba(135,100,184,0.08)' }, data: DATES.map((d, i) => WEEKEND.has(dateDay(d)) ? [{ xAxis: shortDate(DATES[i]) }, { xAxis: shortDate(DATES[i]) }] : null).filter(Boolean) });
function recCard(x) {
  return '<div class="rcard" style="--pc:' + (CATC[x.cat] || '#69797E') + '"><div class="rc-h">' + catPill(x.cat) + ' <code>' + esc(x.id) + '</code> <span class="muted sm">' + esc(x.rule.n || '') + '</span><span class="sp"></span>' + (x.sa ? '<b class="pos">' + money(x.sa) + '/yr</b>' : '') + '</div>' +
    '<div class="rc-a">' + esc(x.act) + '</div>' +
    (x.tgt ? '<div class="rc-k"><span>Target</span><div>' + esc(x.tgt) + '</div></div>' : '') +
    (x.ev ? '<div class="rc-k"><span>Evidence</span><div>' + esc(x.ev) + '</div></div>' : '') +
    (x.rule.c ? '<div class="rc-k"><span>Rule criteria</span><div>' + esc(x.rule.c) + '</div></div>' : '') +
    (x.rule.r ? '<div class="rc-k"><span>Why</span><div>' + esc(x.rule.r) + '</div></div>' : '') +
    '<div class="rc-k"><span>Confidence / effort</span><div>' + esc(x.conf || '–') + ' / ' + esc(x.eff || '–') + (x.pr ? ' · ' + priPill(x.pr) : '') + '</div></div>' +
    (x.cmt ? '<div class="rc-k"><span>Guidance</span><div>' + linkify(x.cmt) + '</div></div>' : '') + '</div>';
}
function openRes(ri) {
  const r = typeof ri === 'number' ? R[ri] : ri; if (!r) return;
  disposeDrawerCharts();
  const t = r.T, dr = $('#drawer');
  const facts = [['Subscription', r.sub], ['Resource group', r.g], ['Location', r.l], ['Environment', r.e], ['SKU / size', r.sku], ['Kind', r.k], ['Created', r.cr], ['Age', isNum(r.ag) ? fmt(r.ag) + ' days' : ''],
    ['Monthly cost', isNum(r.c) ? money(r.c) + (r.x ? ' (allocated share)' : '') : ''], ['Cluster cost', isNum(r.cc) ? money(r.cc) : ''], ['Retail PAYG / month', isNum(r.rpm) ? money(r.rpm) : ''],
    ['Effective discount', isNum(r.ed) ? pctS(r.ed, 1) + ' vs retail' : ''], ['Est. savings', r.sv ? money(r.sv) + ' / year' : ''], ['Metrics status', r.ms]].filter(f => f[1] !== '' && f[1] != null);
  const aggOrder = ['Avg', 'Med', 'P95', 'P99', 'Max', 'MinAvg', 'Min', 'P05', 'Total', 'MaxT', 'PeakHour'];
  const byA = new Map();
  (t.mc || []).forEach((c, j) => { const v = r.m ? r.m[j] : null; if (!isNum(v)) return; const s = mSplit(c); if (!byA.has(s.a)) byA.set(s.a, {}); byA.get(s.a)[s.g] = v; });
  const used = aggOrder.filter(g => Array.from(byA.values()).some(o => o[g] != null));
  const mt = byA.size ? '<div class="tw" style="max-height:360px"><table class="t"><thead><tr><th>Metric</th>' + used.map(g => '<th class="num">' + esc(AGGN[g]) + '</th>').join('') + '</tr></thead><tbody>' +
    Array.from(byA.entries()).map(([a, o]) => '<tr><td><b>' + esc(t.ml[a] || a) + '</b>' + (t.mu[a] ? ' <span class="muted sm">(' + esc(t.mu[a]) + ')</span>' : '') + ((t.ml[a] || a) !== a ? '<div class="muted sm">' + esc(a) + '</div>' : '') + '</td>' + used.map(g => '<td class="num">' + fmtA(o[g]) + '</td>').join('') + '</tr>').join('') + '</tbody></table></div>' : emptyMsg('No metric values were collected for this resource.');
  const fcRows = (t.fc || []).map((c, j) => [c, r.f ? r.f[j] : null]).filter(x => x[1] != null && x[1] !== '');
  const pcRows = (t.pc || []).map((c, j) => [c.replace(/^Prop\./, ''), r.p ? r.p[j] : null]).filter(x => x[1] != null && x[1] !== '');
  const tagKeys = Object.keys(r.tg);
  const series = r.dm ? Object.keys(r.dm).sort((a, b) => (a === r.pm ? -1 : b === r.pm ? 1 : a.localeCompare(b))) : (r.d ? [r.pm || 'Primary metric'] : []);
  const pat = r.pat;
  const gauge = isNum(r.u) ? '<div class="gauge"><div class="sm muted">Utilization of the provisioned capacity – primary metric, P95 (thresholds: idle &lt; ' + TH.idle + '%, over-provisioned &lt; ' + TH.low + '%, target ' + TH.target + '%, hot &gt; ' + TH.high + '%)</div><div class="gbar" style="margin-bottom:24px"><div class="gmark" style="left:calc(' + Math.min(100, Math.max(0, r.u)) + '% - 1px)"><span>' + fmt(r.u, 1) + '%</span></div>' +
    [[TH.low, 'low'], [TH.target, 'target'], [TH.high, 'high']].map(x => '<span class="gtick" style="left:' + x[0] + '%">' + x[0] + '% ' + x[1] + '</span>').join('') + '</div></div>' : '';
  dr.innerHTML = '<div class="dr-h"><div style="min-width:0"><div class="dr-t">' + esc(r.tn) + ' · ' + esc(r.cat) + '</div><h2>' + esc(r.n) + '</h2><div class="rc-h">' + stPill(r.st) + (r.rp ? ' ' + pill(r.rp + ' priority', '#fff') : '') + ' ' + r.recs.map(x => catPill(x.cat)).join(' ') + '</div></div>' +
    '<div class="dr-r">' + ring(r.es, 66, 7) + '<small>Efficiency</small></div><button class="icon-btn dr-x" id="drClose" title="Close (Esc)">' + ICON.x + '</button></div>' +
    '<div class="dr-b">' +
    '<div class="facts">' + facts.map(f => '<div><span>' + esc(f[0]) + '</span><b title="' + esc(f[1]) + '">' + esc(f[1]) + '</b></div>').join('') + '</div>' +
    '<div class="pvu"><div><h4>Provisioned</h4><p>' + esc(r.pv || '–') + '</p></div><div><h4>Utilized</h4><p>' + esc(r.us || '–') + '</p></div></div>' + gauge +
    (r.an ? '<div class="note">' + esc(r.an) + '</div>' : '') +
    '<h3>Suggested recommendations (' + r.recs.length + ')</h3>' + (r.recs.length ? r.recs.map(recCard).join('') : '<div class="muted">No recommendation – utilization is within the thresholds.</div>') +
    (series.length ? '<h3>Daily utilization</h3><div class="card"><div class="card-b"><div class="ctl">Metric <select class="sel" id="drM">' + series.map(s => '<option' + (s === r.pm ? ' selected' : '') + '>' + esc(s) + '</option>').join('') + '</select><span class="muted sm">Shaded columns = weekend (' + esc((M.weekend || []).join('/')) + ')</span></div>' + chartBox('drChart', 250) + '</div></div>' : '') +
    (pat ? '<h3>Usage pattern (' + esc(r.pm || 'primary metric') + ', ' + esc(M.tz) + ')</h3><div class="facts">' +
      [['Active hours', pctS(pat[0], 0)], ['Idle hours', pctS(pat[1], 0)], ['Business-hours avg', fmtA(pat[2])], ['Off-hours avg', fmtA(pat[3]) + (isNum(pat[2]) && pat[2] > 0 && isNum(pat[3]) ? ' (' + pctS(100 * pat[3] / pat[2]) + ' of business hours)' : '')], ['Weekday avg', fmtA(pat[4])], ['Weekend avg', fmtA(pat[5])], ['Peak', (pat[7] || '') + (isNum(pat[6]) ? ' ' + String(pat[6]).padStart(2, '0') + ':00' : '')]]
        .map(f => '<div><span>' + esc(f[0]) + '</span><b>' + esc(f[1]) + '</b></div>').join('') + '</div>' : '') +
    '<h3>Metrics – window statistics</h3>' + mt +
    (fcRows.length ? '<h3>Configuration</h3><div class="kvw"><table class="kv">' + fcRows.map(x => '<tr><td>' + esc(x[0]) + '</td><td>' + esc(x[1]) + '</td></tr>').join('') + '</table></div>' : '') +
    '<h3>Tags (' + tagKeys.length + ')</h3><div>' + (tagKeys.length ? tagKeys.map(k => '<span class="tagc" data-tk="' + esc(k) + '" data-tv="' + esc(r.tg[k]) + '" title="Filter the dashboard by this tag"><b>' + esc(k) + '</b> ' + esc(r.tg[k]) + '</span>').join('') : '<span class="muted">No tags</span>') + '</div>' +
    (pcRows.length ? '<h3>All properties (' + pcRows.length + ')</h3><div class="dt-s" style="margin-bottom:6px">' + ICON.search + '<input id="drPq" placeholder="Search properties…"></div><div class="kvw"><table class="kv" id="drPt">' + pcRows.map(x => '<tr><td>' + esc(x[0]) + '</td><td>' + esc(x[1]) + '</td></tr>').join('') + '</table></div>' : '') +
    '<div class="rc-h" style="margin-top:6px">' + (r.id ? '<a class="btn pri" target="_blank" rel="noopener" href="https://portal.azure.com/#@' + esc(M.tenant) + '/resource' + esc(r.id) + '">' + ICON.ext + ' Open in Azure portal</a><button class="btn" id="drCopy">Copy resource ID</button>' : '') + '</div>' +
    '<div class="muted sm" style="word-break:break-all">' + esc(r.id) + '</div></div>';
  dr.classList.add('open'); $('#drawerMask').classList.add('open'); dr.scrollTop = 0;
  $('#drClose').onclick = closeDrawer;
  const cp = $('#drCopy'); if (cp) cp.onclick = () => copyText(r.id);
  $$('.tagc', dr).forEach(c => { c.onclick = () => { closeDrawer(); setTag(c.dataset.tk, c.dataset.tv); toast('Filtered by tag ' + c.dataset.tk + ' = ' + c.dataset.tv); }; });
  const pq = $('#drPq'); if (pq) pq.addEventListener('input', () => { const q = pq.value.toLowerCase(); $$('#drPt tr').forEach(tr => { tr.style.display = tr.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none'; }); });
  const draw = name => {
    disposeDrawerCharts();
    const e = r.dm ? r.dm[name] : { k: r.dk || 'a', v: r.d }; if (!e) return;
    const pctLike = /pct|percent|cpu|util|dtu/i.test(name);
    const base = { name: name, type: e.k === 't' ? 'bar' : 'line', data: e.v, smooth: true, showSymbol: false, markArea: weekendArea(), itemStyle: { color: '#0078D4', borderRadius: e.k === 't' ? [4, 4, 0, 0] : 0 }, areaStyle: e.k === 't' ? undefined : { opacity: .12 }, lineStyle: { width: 2.2 } };
    if (pctLike && e.k !== 't') base.markLine = { silent: true, symbol: 'none', lineStyle: { type: 'dashed' }, data: [{ yAxis: TH.low, label: { formatter: 'low ' + TH.low + '%', position: 'insideEndTop' }, lineStyle: { color: '#F7630C' } }, { yAxis: TH.high, label: { formatter: 'high ' + TH.high + '%', position: 'insideEndTop' }, lineStyle: { color: '#8764B8' } }] };
    const s = [Object.assign(base, { name: name + (e.k === 't' ? ' · daily total' : ' · daily average') })];
    if (e.x && e.x.some(isNum)) s.push({ name: name + ' · daily max', type: 'line', data: e.x, smooth: true, showSymbol: false, itemStyle: { color: '#8764B8' }, lineStyle: { width: 1.4, type: 'dashed' } });
    ec('drChart', { grid: { left: 8, right: 20, top: 34, bottom: 8, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', valueFormatter: v => fmtA(v) }, xAxis: { type: 'category', data: DATES.map(shortDate) }, yAxis: { type: 'value', axisLabel: { formatter: v => fmtA(v) } }, series: s }, null, DRCH);
  };
  const sel = $('#drM'); if (sel) { sel.onchange = () => draw(sel.value); setTimeout(() => draw(sel.value), 260); }
}
$('#drawerMask').addEventListener('click', closeDrawer);

// ---------------------------------------------------------------- command palette
let PAL_ITEMS = [], PAL_SEL = 0;
function openPalette() { const p = $('#palette'); p.classList.add('open'); const i = $('#palInput'); i.value = ''; palRender(''); setTimeout(() => i.focus(), 20); }
function closePalette() { $('#palette').classList.remove('open'); }
function palRender(q) {
  q = (q || '').trim().toLowerCase();
  const pg = PAGES.filter(p => !q || p.n.toLowerCase().indexOf(q) >= 0).map(p => ({ g: 'Pages', ic: p.ic, t: p.n, s: '', a: () => go(p.k) }));
  const ty = T.filter(t => !q || t.n.toLowerCase().indexOf(q) >= 0 || t.c.toLowerCase().indexOf(q) >= 0).sort((a, b) => b.res.length - a.res.length).slice(0, 8).map(t => ({ g: 'Service dashboards', ic: CAT_ICON[t.c] || 'grid', t: t.n, s: fmt(t.res.length) + ' resources · ' + t.c, a: () => go('services/' + encodeURIComponent(t.k)) }));
  const rs = (q ? R.filter(r => r.hay.indexOf(q) >= 0).sort((a, b) => b.tc - a.tc) : R.filter(r => r.sv > 0).sort((a, b) => b.sv - a.sv)).slice(0, q ? 14 : 6)
    .map(r => ({ g: q ? 'Resources' : 'Top savings opportunities', ic: CAT_ICON[r.cat] || 'box', t: r.n, s: r.tn + ' · ' + (r.sv ? money(r.sv) + '/yr' : money(r.c) + '/mo'), a: () => openRes(r) }));
  const ru = Object.keys(D.rules || {}).filter(k => !q || k.toLowerCase().indexOf(q) >= 0 || (D.rules[k].n || '').toLowerCase().indexOf(q) >= 0).slice(0, q ? 8 : 4).map(k => ({ g: 'Recommendation rules', ic: 'bulb', t: k + ' – ' + D.rules[k].n, s: fmt(RC.filter(x => x.rc === k).length) + ' recommendations', a: () => go('recs/' + encodeURIComponent(k)) }));
  PAL_ITEMS = pg.concat(ty, rs, ru); PAL_SEL = 0;
  let html = '', lastG = '';
  PAL_ITEMS.forEach((it, i) => { if (it.g !== lastG) { html += '<div class="pal-g">' + esc(it.g) + '</div>'; lastG = it.g; } html += '<div class="pal-i' + (i === PAL_SEL ? ' on' : '') + '" data-i="' + i + '">' + (ICON[it.ic] || '') + '<span>' + esc(it.t) + '</span><small>' + esc(it.s) + '</small></div>'; });
  $('#palList').innerHTML = html || emptyMsg('No matches');
}
function palMove(d) { const items = $$('.pal-i'); if (!items.length) return; PAL_SEL = (PAL_SEL + d + items.length) % items.length; items.forEach((x, i) => x.classList.toggle('on', i === PAL_SEL)); items[PAL_SEL].scrollIntoView({ block: 'nearest' }); }
function palRun(i) { const it = PAL_ITEMS[i]; if (!it) return; closePalette(); it.a(); }
$('#palInput').addEventListener('input', e => palRender(e.target.value));
$('#palInput').addEventListener('keydown', e => { if (e.key === 'ArrowDown') { e.preventDefault(); palMove(1); } else if (e.key === 'ArrowUp') { e.preventDefault(); palMove(-1); } else if (e.key === 'Enter') { e.preventDefault(); palRun(PAL_SEL); } else if (e.key === 'Escape') closePalette(); });
$('#palList').addEventListener('click', e => { const it = e.target.closest('.pal-i'); if (it) palRun(+it.dataset.i); });
$('#palette').addEventListener('mousedown', e => { if (e.target.id === 'palette') closePalette(); });

// ---------------------------------------------------------------- shared calculations
function costTrend() {
  const cs = D.cost || {}; if (!cs.months || !cs.months.length) return null;
  const tot = cs.months.map(() => 0); (cs.cm || []).forEach(x => { if (F.sub.size && !F.sub.has(D.subs[x[1]])) return; tot[x[0]] += x[3]; });
  const partial = new Set(cs.partial || []); const all = cs.months.map((m, i) => ({ m, v: tot[i], p: partial.has(m) }));
  return { all, full: all.filter(x => !x.p && x.v > 0) };
}
function trendChart(id, topN) {
  const cs = D.cost || {}; if (!cs.months || !cs.months.length || !(cs.cm || []).length) { $('#' + id).innerHTML = emptyMsg('No cost history available (Cost Management data not collected).'); return; }
  const nM = cs.months.length, bySvc = new Map();
  cs.cm.forEach(x => { if (F.sub.size && !F.sub.has(D.subs[x[1]])) return; const s = cs.svcs[x[2]]; if (!bySvc.has(s)) bySvc.set(s, new Array(nM).fill(0)); bySvc.get(s)[x[0]] += x[3]; });
  const ranked = Array.from(bySvc.entries()).map(e => ({ s: e[0], v: e[1], t: sum(e[1]) })).sort((a, b) => b.t - a.t);
  const top = ranked.slice(0, topN || 7), rest = ranked.slice(topN || 7);
  if (rest.length) top.push({ s: 'Other services', v: cs.months.map((m, i) => sum(rest, x => x.v[i])) });
  const partial = new Set(cs.partial || []), totals = cs.months.map((m, i) => sum(top, x => x.v[i]));
  ec(id, {
    grid: { left: 8, right: 14, top: 44, bottom: 6, containLabel: true }, legend: { top: 0 },
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => money(v) },
    xAxis: { type: 'category', data: cs.months.map(m => partial.has(m) ? m + ' *' : m) }, yAxis: { type: 'value', axisLabel: axisMoney },
    series: top.map((x, i) => ({ name: x.s, type: 'bar', stack: 'c', barMaxWidth: 46, emphasis: { focus: 'series' }, itemStyle: { color: PAL[i % PAL.length] }, data: x.v.map((v, j) => ({ value: Math.round(v), itemStyle: { opacity: partial.has(cs.months[j]) ? 0.45 : 1 } })) }))
      .concat([{ name: 'Total', type: 'line', data: totals.map(v => Math.round(v)), symbolSize: 7, itemStyle: { color: css('--text') }, lineStyle: { width: 2 }, label: { show: true, position: 'top', formatter: p => moneyK(p.value), color: css('--muted'), fontSize: 11 } }])
  });
}
function capacityHtml(rs) {
  const vt = tByKey.VirtualMachines; if (!vt) return '';
  const vms = rs.filter(r => r.T === vt && String(fv(r, 'PowerState') || '').toLowerCase() === 'running' && isNum(fv(r, 'vCPUs')));
  if (!vms.length) return '';
  const vc = sum(vms, r => fv(r, 'vCPUs')), cpuAvg = sum(vms, r => fv(r, 'vCPUs') * (fv(r, 'CPU_Avg') || 0) / 100), cpuP95 = sum(vms, r => fv(r, 'vCPUs') * (fv(r, 'CPU_P95') || 0) / 100);
  const memV = vms.filter(r => isNum(fv(r, 'MemUsedPct_Avg')) && isNum(fv(r, 'MemoryGB')));
  const memM = sum(memV, r => fv(r, 'MemoryGB')), memAvg = sum(memV, r => fv(r, 'MemoryGB') * fv(r, 'MemUsedPct_Avg') / 100), memPk = sum(memV, r => fv(r, 'MemoryGB') * (isNum(fv(r, 'MemUsedPct_Max')) ? fv(r, 'MemUsedPct_Max') : fv(r, 'MemUsedPct_Avg')) / 100);
  const bar = (lbl, used, prov, col, note) => '<div class="capbar"><div class="lbl"><span>' + lbl + '</span><span><b>' + fmt(used, 0) + '</b> of ' + fmt(prov, 0) + ' (' + pctS(prov ? 100 * used / prov : null) + ')' + (note ? ' <span class="muted sm">' + note + '</span>' : '') + '</span></div><div class="bar"><i style="width:' + (prov ? Math.min(100, 100 * used / prov) : 0) + '%;background:' + col + '"></i></div></div>';
  return card('Compute capacity – provisioned vs. used (' + fmt(vms.length) + ' running VMs)', bar('vCPUs used at P95', cpuP95, vc, '#0078D4') + bar('vCPUs used on average', cpuAvg, vc, '#00B7C3') +
    (memV.length ? bar('Memory GB at peak', memPk, memM, '#8764B8', '(' + memV.length + ' VMs with memory metrics)') + bar('Memory GB on average', memAvg, memM, '#B4009E') : '<div class="muted sm">Memory metrics were not available for these VMs.</div>') +
    '<div class="sm muted">Unused capacity at P95: <b>' + fmt(vc - cpuP95, 0) + ' vCPUs</b>' + (memV.length ? ' and <b>' + fmt(memM - memPk, 0) + ' GB</b> of memory' : '') + '.</div>', { info: 'Running VMs only: vCPUs × CPU % and memory GB × used memory % (Azure Monitor platform metrics).' });
}
function insights(rs, recs) {
  const out = [], tot = sum(rs, r => r.tc);
  const lowU = rs.filter(r => isNum(r.u) && r.u < TH.low && r.tc > 0), lowC = sum(lowU, r => r.tc);
  if (tot > 0 && lowU.length) out.push(['#F7630C', 'down', '<b>' + pctS(100 * lowC / tot) + '</b> of the monthly run-rate (' + moneyK(lowC) + ') sits on <b>' + fmt(lowU.length) + '</b> resources whose P95 utilization is below ' + TH.low + '% of their provisioned SKU.']);
  const idle = rs.filter(r => r.sgi === 0); if (idle.length) out.push(['#D13438', 'trash', '<b>' + fmt(idle.length) + '</b> resources are idle, unused or orphaned, costing <b>' + moneyK(sum(idle, r => r.tc)) + '/month</b> – the fastest savings (remove or consolidate).']);
  const top5 = recs.filter(x => x.sa > 0).sort((a, b) => b.sa - a.sa).slice(0, 5), allS = dedupSav(recs).total * 12, topS = Math.min(allS, dedupSav(top5).total * 12);
  if (top5.length && allS > 0) out.push(['#13A10E', 'star', 'The top ' + top5.length + ' recommendations represent <b>' + pctS(100 * topS / allS) + '</b> of the identified savings (' + moneyK(topS) + '/yr), led by <b>' + esc(top5[0].n) + '</b> (' + esc(top5[0].tn) + ').']);
  const pt = rs.filter(r => r.pat && isNum(r.pat[4]) && isNum(r.pat[5]) && r.pat[4] > 0);
  if (pt.length >= 3) { const wd = avg(pt.map(r => r.pat[4])), we = avg(pt.map(r => r.pat[5])); if (wd > 0) { const ch = 100 * (we - wd) / wd; out.push(['#8764B8', 'clock', 'Weekend (' + esc((M.weekend || []).join('/')) + ') load is <b>' + (ch < 0 ? fmt(-ch) + '% lower' : fmt(ch) + '% higher') + '</b> than on weekdays across ' + fmt(pt.length) + ' monitored resources' + (ch < -30 ? ' – a strong signal for schedules and auto-scaling.' : '.')]); } }
  const offs = rs.filter(r => r.pat && isNum(r.pat[2]) && isNum(r.pat[3]) && r.pat[2] > 0 && r.pat[3] / r.pat[2] < 0.35 && r.e !== 'Production');
  if (offs.length) out.push(['#00B7C3', 'pause', '<b>' + fmt(offs.length) + '</b> non-production resources drop below 35% of their business-hours load outside ' + M.bizStart + ':00–' + M.bizEnd + ':00 – candidates for start/stop schedules.']);
  const la = (D.la || []).filter(w => passSR(D.subs[w.s], w.g, w.n));
  if (la.length) { const gb = sum(la, w => w.avg), com = la.filter(w => w.com || w.senc); out.push(['#0078D4', 'shield', fmt(la.length) + ' Log Analytics workspace' + (la.length > 1 ? 's' : '') + ' ingest <b>' + fmtA(gb) + ' GB/day</b> on average' + (com.length ? '; ' + com.length + ' on commitment tiers (avg commitment utilization ' + pctS(avg(com.map(w => w.cu))) + ')' : '; none uses a commitment tier') + '.']); }
  const cats = Array.from(groupBy(rs.filter(r => isNum(r.es)), r => r.cat).entries()).map(e => ({ c: e[0], s: wscore(e[1]), cost: sum(e[1], r => r.tc) })).filter(x => isNum(x.s) && x.cost > 0).sort((a, b) => a.s - b.s);
  if (cats.length >= 2) out.push(['#C19C00', 'chart', 'Efficiency is lowest in <b>' + esc(cats[0].c) + '</b> (score ' + Math.round(cats[0].s) + ') and highest in <b>' + esc(cats[cats.length - 1].c) + '</b> (' + Math.round(cats[cats.length - 1].s) + ').']);
  const under = rs.filter(r => r.sgi === 3); if (under.length) out.push(['#8764B8', 'bolt', '<b>' + fmt(under.length) + '</b> resources run hot, throttle or exceed their commitment – address these performance risks alongside the savings.']);
  const np = rs.filter(r => r.e === 'Non-Production'), npC = sum(np, r => r.tc);
  if (tot > 0 && np.length) out.push(['#4F6BED', 'layers', 'Non-production accounts for <b>' + pctS(100 * npC / tot) + '</b> of spend (' + moneyK(npC) + '/month across ' + fmt(np.length) + ' resources).']);
  const tr = costTrend();
  if (tr && tr.full.length >= 2) { const a = tr.full[0], b = tr.full[tr.full.length - 1]; if (a.v > 0) { const ch = 100 * (b.v - a.v) / a.v; out.push([ch <= 0 ? '#13A10E' : '#D13438', ch <= 0 ? 'down' : 'up', 'Monthly cost moved <b>' + (ch <= 0 ? '' : '+') + fmt(ch, 1) + '%</b> from ' + a.m + ' (' + moneyK(a.v) + ') to ' + b.m + ' (' + moneyK(b.v) + ')' + (ch <= -5 ? ' – the impact of the previous optimization is visible.' : ch >= 5 ? ' – spend is growing again; review new workloads.' : ' – broadly flat.')]); } }
  return out;
}

// ---------------------------------------------------------------- page: executive overview
function renderOverview(el) {
  const rs = fr(), recs = frecs();
  if (!R.length) { el.innerHTML = card('No resources', emptyMsg('The assessment did not return any analyzable resources.')); return; }
  const tot = sum(rs, r => r.tc), savY = sum(rs, r => r.sv), score = wscore(rs);
  const by = SG.map((g, i) => rs.filter(r => r.sgi === i)), costBy = by.map(l => sum(l, r => r.tc));
  const qw = recs.filter(isQW), qwSav = dedupSav(qw).total * 12, nTypes = new Set(rs.map(r => r.t)).size;
  const withM = rs.filter(r => r.ms === 'OK' || r.ms === 'Partial').length;
  const typeRows = Array.from(groupBy(rs.filter(r => isNum(r.u)), r => r.t).entries()).map(e => ({ t: T[e[0]], u: wavg(e[1], r => Math.min(r.u, 100), r => r.tc), cost: sum(e[1], r => r.tc), n: e[1].length })).sort((a, b) => b.cost - a.cost || b.n - a.n).slice(0, 18);
  const sp = savSplit(rs);
  const savRatio = tot > 0 && savY / 12 <= tot ? pctS(100 * savY / 12 / tot) + ' of run-rate' : (savY > 0 && tot > 0 ? 'includes dependent disks / IPs' : '');
  const pl = (n, w) => fmt(n) + ' ' + w + (n === 1 ? '' : 's');
  const summary = 'Across <b>' + fmt(rs.length) + '</b> analyzed resources in <b>' + pl(nTypes, 'service') + '</b>, the monthly run-rate is <b>' + (M.hasCost ? money(tot) : 'n/a') + '</b> and the cost-weighted efficiency score is <b>' + (isNum(score) ? Math.round(score) : 'n/a') + '/100</b> (P95 utilization against a ' + TH.target + '% target). ' +
    '<b>' + moneyK(savY) + '/year</b> of savings are identified' + (savRatio ? ' (' + savRatio + ')' : '') + ': ' + moneyK(sp.Delete) + ' from removing idle/orphaned resources, ' + moneyK(sp.Rightsize) + ' from right-sizing and ' + moneyK(sp.Configure) + ' from configuration and commitment changes. ' +
    '<b>' + fmt(qw.length) + '</b> low-effort quick wins are worth ' + moneyK(qwSav) + '/year; <b>' + fmt(by[3] ? by[3].length : 0) + '</b> resources show performance risk.';
  el.innerHTML =
    '<div class="hero"><div><div class="eyebrow">' + esc(M.customer) + ' · Azure Utilization &amp; Cost Optimizer</div><h2>Provisioned vs. utilized across ' + pl(rs.length, 'resource') + ' and ' + pl(nTypes, 'service') + '</h2>' +
    '<p>' + esc(M.start) + ' → ' + esc(M.end) + ' · ' + M.days + ' days · ' + esc(M.tz) + ' · ' + fmt(M.subs) + ' subscriptions · ' + fmt(M.metricRequests) + ' Azure Monitor requests' + (anyFilter() ? ' · <b>filtered view</b>' : '') + '</p></div>' +
    '<div class="hero-k"><div><span>Monthly run-rate</span><b>' + (M.hasCost ? moneyK(tot) : 'n/a') + '</b><small>' + (M.hasCost ? 'analyzed resources · ' + esc(M.costType) : 'cost data unavailable') + '</small></div>' +
    '<div><span>Savings identified</span><b class="g">' + moneyK(savY) + '</b><small>per year' + (savRatio ? ' · ' + savRatio : '') + '</small></div>' +
    '<div class="hero-ring" title="Cost-weighted efficiency score (0-100)">' + ring(score, 98, 10) + '<small>Efficiency score</small></div></div></div>' +
    '<div class="summary">' + summary + '</div>' +
    '<div class="kpis">' +
    kpi('Idle · unused · orphaned', fmt(by[0].length), moneyK(costBy[0]) + ' / month paid for nothing', SG[0].c, 'trash') +
    kpi('Over-provisioned', fmt(by[1].length), moneyK(costBy[1]) + ' / month at low utilization', SG[1].c, 'down') +
    kpi('Right-sized · in use', fmt(by[2].length), pctS(rs.length ? 100 * by[2].length / rs.length : null) + ' of resources', SG[2].c, 'check') +
    kpi('Under-provisioned', fmt(by[3].length), 'performance / throttling risk', SG[3].c, 'up') +
    kpi('Stopped · deallocated', fmt(by[4].length), moneyK(costBy[4]) + ' / month still billed (disks, IPs)', SG[4].c, 'pause') +
    kpi('Recommendations', fmt(recs.length), fmt(qw.length) + ' quick wins · ' + moneyK(qwSav) + '/yr', '#0078D4', 'bulb') +
    kpi('Services analyzed', fmt(nTypes), fmt(new Set(rs.filter(r => r.T.a).map(r => r.t)).size) + ' with auto-discovered metrics', '#00B7C3', 'grid') +
    kpi('Metric coverage', pctS(rs.length ? 100 * withM / rs.length : null), fmt(withM) + ' resources with utilization data', '#4F6BED', 'chart') + '</div>' +
    '<div class="grid g3">' + card('Where the money goes – by utilization status', chartBox('ovStatus', 300), { info: 'Monthly cost of the filtered resources grouped by utilization status.' }) +
    card('From current to optimized run-rate', chartBox('ovWater', 300), { info: 'Monthly run-rate and the de-duplicated monthly savings per action type (Delete supersedes other actions on the same resource).' }) +
    card('Efficiency score by service category', chartBox('ovCat', 300), { info: 'Cost-weighted efficiency (P95 utilization vs. the ' + TH.target + '% target; idle = 0).' }) + '</div>' +
    '<div class="grid g21">' + card('Cost vs. utilization – every resource', '<div class="legend-s" id="ovScLeg"></div>' + chartBox('ovScatter', 420), { info: 'X: utilization of the provisioned SKU (P95 %). Y: monthly cost (log scale). Bubble size: savings. Red zone = expensive and under-utilized. Click a bubble for details.' }) +
    card('Spend map – category › service', chartBox('ovTree', 440), { info: 'Area = monthly cost; color = efficiency score (red = wasteful, green = efficient). Click a service to open its dashboard.' }) + '</div>' +
    '<div class="grid g2">' + card('Utilization of provisioned capacity by service', chartBox('ovUtil', Math.max(240, 26 * typeRows.length + 50)), { info: 'Cost-weighted average utilization % of the current SKU per service (top 18 by cost).' }) +
    card('Top savings opportunities', '<div class="toplist" id="ovTop"></div>', { actions: '<a class="btn sm" href="#/recs">All recommendations</a>' }) + '</div>' +
    '<div class="grid g2">' + card('Monthly cost trend', chartBox('ovTrend', 320), { info: 'Cost Management history (' + M.costType + '); * = partial month. Respects the subscription filter.' }) +
    card('Insights', '<div class="insights" id="ovIns"></div>') + '</div>' + capacityHtml(rs);

  const useCost = tot > 0;
  ec('ovStatus', donutOpt(SG.map((g, i) => ({ name: g.n, value: useCost ? Math.round(costBy[i]) : by[i].length, itemStyle: { color: g.c } })).filter(x => x.value > 0), useCost ? moneyK(tot) : fmt(rs.length), useCost ? 'per month' : 'resources', useCost ? money : fmt),
    p => { const g = SG.find(x => x.n === p.name); if (g) setFilter('sg', [g.n]); });
  const wcur = tot, w1 = sp.Delete / 12, w2 = sp.Rightsize / 12, w3 = sp.Configure / 12, wopt = Math.max(0, wcur - w1 - w2 - w3);
  ec('ovWater', {
    grid: { left: 8, right: 10, top: 26, bottom: 6, containLabel: true }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, formatter: ps => { const p = ps.find(x => x.seriesName === 'v'); return p ? '<b>' + p.name + '</b><br>' + money(p.value) + ' / month' : ''; } },
    xAxis: { type: 'category', data: ['Current', 'Delete', 'Rightsize', 'Configure', 'Optimized'], axisLabel: { interval: 0 } }, yAxis: { type: 'value', axisLabel: axisMoney },
    series: [{ name: 'b', type: 'bar', stack: 'w', silent: true, itemStyle: { color: 'transparent' }, data: [0, wcur - w1, wcur - w1 - w2, wopt, 0] },
      { name: 'v', type: 'bar', stack: 'w', barMaxWidth: 54, label: { show: true, position: 'top', formatter: p => moneyK(p.value), color: css('--muted'), fontSize: 11 }, data: [[wcur, '#4F6BED'], [w1, CATC.Delete], [w2, CATC.Rightsize], [w3, '#00B7C3'], [wopt, '#13A10E']].map(x => ({ value: Math.round(x[0]), itemStyle: { color: x[1], borderRadius: 6 } })) }]
  });
  const cats = Array.from(groupBy(rs, r => r.cat).entries()).map(e => ({ c: e[0], s: wscore(e[1]), cost: sum(e[1], r => r.tc), n: e[1].length })).filter(x => isNum(x.s)).sort((a, b) => b.cost - a.cost);
  if (cats.length) ec('ovCat', hbarOpt(cats.map(x => x.c), cats.map(x => Math.round(x.s)), cats.map(x => scoreColor(x.s)), v => fmt(v), { xAxis: { type: 'value', max: 100 }, tooltip: { trigger: 'item', formatter: p => { const x = cats[p.dataIndex]; return '<b>' + esc(x.c) + '</b><br>Efficiency ' + Math.round(x.s) + '/100<br>' + money(x.cost) + '/month · ' + fmt(x.n) + ' resources'; } } }), p => { const x = cats[p.dataIndex]; if (x) setFilter('cat', [x.c]); });
  else $('#ovCat').innerHTML = emptyMsg();
  const pts = rs.filter(r => isNum(r.u) && r.tc > 0);
  if (pts.length) {
    const medC = pctl(pts.map(r => r.tc), 50) || 1, maxC = Math.max.apply(null, pts.map(r => r.tc)), maxSv = Math.max(1, Math.max.apply(null, pts.map(r => r.sv || 0)));
    $('#ovScLeg').innerHTML = SG.map(g => '<span><i class="dot" style="background:' + g.c + '"></i>' + esc(g.n) + '</span>').join('');
    ec('ovScatter', {
      grid: { left: 10, right: 24, top: 18, bottom: 34, containLabel: true },
      tooltip: { trigger: 'item', formatter: p => { const r = R[p.value[3]]; return '<b>' + esc(r.n) + '</b><br>' + esc(r.tn) + ' · ' + esc(r.sku || '') + '<br>Utilization ' + fmt(r.u, 1) + '% · ' + money(r.tc) + '/month' + (r.sv ? '<br><b>Savings ' + money(r.sv) + '/yr</b>' : '') + '<br>' + esc(r.st); } },
      xAxis: { type: 'value', name: 'Utilization of provisioned capacity (P95 %)', nameLocation: 'middle', nameGap: 26, min: 0, max: v => Math.max(100, Math.ceil(v.max / 10) * 10) },
      yAxis: { type: 'log', logBase: 10, min: 1, name: 'Monthly cost', axisLabel: axisMoney },
      series: SG.map((g, gi) => ({ name: g.n, type: 'scatter', data: pts.filter(r => r.sgi === gi).map(r => [Math.min(r.u, 150), Math.max(1, r.tc), r.sv || 0, r.i]), symbolSize: v => 8 + 30 * Math.sqrt((v[2] || 0) / maxSv), itemStyle: { color: g.c, opacity: 0.78, borderColor: '#fff', borderWidth: 0.6 },
        markArea: gi === 0 ? { silent: true, itemStyle: { color: 'rgba(209,52,56,0.07)' }, label: { show: true, position: 'insideTopLeft', formatter: 'Waste zone: under-utilized and costly', color: '#D13438', fontSize: 11 }, data: [[{ coord: [0, Math.max(1, medC)] }, { coord: [TH.low, Math.max(maxC * 1.2, medC * 2)] }]] } : undefined,
        markLine: gi === 0 ? { silent: true, symbol: 'none', lineStyle: { type: 'dashed', color: '#8A8886' }, label: { formatter: p => p.value + '%', color: css('--muted') }, data: [{ xAxis: TH.low }, { xAxis: TH.target }, { xAxis: TH.high }] } : undefined }))
    }, p => openRes(p.value[3]));
  } else $('#ovScatter').innerHTML = emptyMsg('No resources with both utilization and cost data in the current filter.');
  const tdata = Array.from(groupBy(rs.filter(r => r.tc > 0), r => r.cat).entries()).map(e => ({ name: e[0], value: sum(e[1], r => r.tc), itemStyle: { color: '#33415C' }, children: Array.from(groupBy(e[1], r => r.t).entries()).map(f => { const s = wscore(f[1]); return { name: T[f[0]].n, k: T[f[0]].k, value: sum(f[1], r => r.tc), s: s, sv: sum(f[1], r => r.sv), n: f[1].length, itemStyle: { color: scoreColor(s) } }; }) }));
  if (tdata.length) ec('ovTree', {
    tooltip: { formatter: p => { const d = p.data; return d.children ? '<b>' + esc(p.name) + '</b><br>' + money(p.value) + '/month' : '<b>' + esc(p.name) + '</b><br>' + money(p.value) + '/month · ' + fmt(d.n) + ' resources<br>Efficiency ' + (isNum(d.s) ? Math.round(d.s) : 'n/a') + ' · savings ' + money(d.sv) + '/yr'; } },
    series: [{ type: 'treemap', data: tdata, roam: false, nodeClick: false, breadcrumb: { show: false }, top: 0, left: 0, right: 0, bottom: 0, upperLabel: { show: true, height: 22, color: '#fff', fontWeight: 600 }, label: { formatter: p => p.name + '\n' + moneyK(p.value), color: '#fff', fontSize: 11, lineHeight: 15 },
      levels: [{ itemStyle: { borderColor: css('--card'), borderWidth: 3, gapWidth: 3 } }, { itemStyle: { borderColor: 'rgba(255,255,255,.35)', borderWidth: 1, gapWidth: 1 } }] }]
  }, p => { if (p.data && p.data.k) go('services/' + encodeURIComponent(p.data.k)); });
  else $('#ovTree').innerHTML = emptyMsg('No cost data in the current filter.');
  if (typeRows.length) ec('ovUtil', hbarOpt(typeRows.map(x => x.t.n), typeRows.map(x => Math.round(x.u * 10) / 10), typeRows.map(x => utilColor(x.u)), v => fmt(v, 0) + '%', {
    xAxis: { type: 'value', max: 100, axisLabel: { formatter: v => v + '%' } },
    tooltip: { trigger: 'item', formatter: p => { const x = typeRows[p.dataIndex]; return '<b>' + esc(x.t.n) + '</b><br>Avg utilization ' + fmt(x.u, 1) + '% (cost-weighted)<br>' + fmt(x.n) + ' resources · ' + money(x.cost) + '/month'; } }
  }), p => { const x = typeRows[p.dataIndex]; if (x) go('services/' + encodeURIComponent(x.t.k)); });
  else $('#ovUtil').innerHTML = emptyMsg();
  const top = recs.filter(x => x.sa > 0).sort((a, b) => b.sa - a.sa).slice(0, 9);
  $('#ovTop').innerHTML = top.length ? top.map(x => '<div class="tl" data-r="' + x.r + '" style="--pc:' + (CATC[x.cat] || '#69797E') + '"><i></i><div style="min-width:0"><b>' + esc(x.n) + '</b><small>' + esc(x.tn) + ' · ' + esc(x.act) + '</small></div><em>' + moneyK(x.sa) + '/yr</em></div>').join('') : emptyMsg('No savings opportunities in the current filter.');
  $$('#ovTop .tl').forEach(e => { e.onclick = () => { const i = +e.dataset.r; if (i >= 0) openRes(i); }; });
  trendChart('ovTrend', 6);
  const ins = insights(rs, recs);
  $('#ovIns').innerHTML = ins.length ? ins.map(x => '<div class="ins" style="--ic:' + x[0] + '"><i>' + (ICON[x[1]] || '') + '</i><div>' + x[2] + '</div></div>').join('') : emptyMsg();
}

// ---------------------------------------------------------------- page: recommendations
const SEL = new Set();
function recCols(withSel) {
  return [
    { k: 'pr', n: 'Priority', f: x => priPill(x.pr), v: x => x.pr },
    { k: 'id', n: 'Rule ID', f: x => '<code title="' + esc((x.rule.n || '') + ' – ' + (x.rule.c || '')) + '">' + esc(x.id) + '</code>', v: x => x.id },
    { k: 'cat', n: 'Action', f: x => catPill(x.cat), v: x => x.cat },
    { k: 'n', n: 'Resource', f: x => '<div class="rn"><b>' + esc(x.n) + '</b><small>' + esc(x.tn) + '</small></div>', v: x => x.n },
    { k: 'act', n: 'Recommended action', f: x => '<div class="wrap">' + esc(x.act) + (x.tgt ? '<div class="muted sm">→ ' + esc(x.tgt) + '</div>' : '') + '</div>', v: x => x.act },
    { k: 'sa', n: 'Savings / yr', num: true, f: x => x.sa ? '<b class="pos">' + money(x.sa) + '</b>' : '<span class="muted">–</span>', v: x => x.sa },
    { k: 'cmc', n: 'Monthly cost', num: true, f: x => money(x.cmc), v: x => x.cmc },
    { k: 'conf', n: 'Confidence', v: x => x.conf },
    { k: 'eff', n: 'Effort', v: x => x.eff },
    { k: 'ev', n: 'Evidence', f: x => '<div class="clip" style="max-width:320px" title="' + esc(x.ev) + '">' + esc(x.ev) + '</div>', v: x => x.ev },
    { k: 'sub', n: 'Subscription', v: x => x.sub },
    { k: 'g', n: 'Resource group', v: x => x.g },
    { k: 'e', n: 'Environment', v: x => x.e },
    { k: 'crit', n: 'Rule criteria', f: x => '<div class="clip" style="max-width:300px" title="' + esc(x.rule.c) + '">' + esc(x.rule.c) + '</div>', v: x => x.rule.c },
    { k: 'tgt', n: 'Target', v: x => x.tgt, f: x => '<div class="clip" style="max-width:220px">' + esc(x.tgt) + '</div>' },
    { k: 'rid', n: 'Resource ID', v: x => x.rid, f: x => '<div class="clip muted" style="max-width:220px">' + esc(x.rid) + '</div>' }
  ];
}
function renderRecs(el, arg) {
  const tab = arg === 'advisor' ? 'adv' : 'rb';
  const recs = frecs(), adv = (D.adv || []).filter(a => passSR(D.subs[a.s], a.g, [a.res, a.prob, a.sol, a.typ].join(' ')));
  el.innerHTML = '<div class="tabs"><a class="tab' + (tab === 'rb' ? ' on' : '') + '" href="#/recs">Suggested recommendations (rule-based)<span class="n">' + fmt(recs.length) + '</span></a><a class="tab' + (tab === 'adv' ? ' on' : '') + '" href="#/recs/advisor">Azure Advisor (cost &amp; performance)<span class="n">' + fmt(adv.length) + '</span></a></div><div id="recBody" style="display:flex;flex-direction:column;gap:16px"></div>';
  if (tab === 'adv') return renderAdvisor($('#recBody'), adv);
  const body = $('#recBody'), ds = dedupSav(recs), qw = recs.filter(isQW);
  const cnt = c => recs.filter(x => x.cat === c).length;
  body.innerHTML = '<div class="kpis">' +
    kpi('Recommendations', fmt(recs.length), fmt(new Set(recs.map(x => x.r)).size) + ' resources affected', '#0078D4', 'bulb') +
    kpi('Estimated savings', moneyK(ds.total * 12), 'per year · de-duplicated per resource', '#13A10E', 'coin') +
    kpi('Delete', fmt(cnt('Delete')), moneyK(ds.Delete * 12) + '/yr', CATC.Delete, 'trash') +
    kpi('Rightsize', fmt(cnt('Rightsize')), moneyK(ds.Rightsize * 12) + '/yr', CATC.Rightsize, 'down') +
    kpi('Configure', fmt(cnt('Configure')), moneyK(ds.Configure * 12) + '/yr', CATC.Configure, 'layers') +
    kpi('Performance risk', fmt(cnt('Performance Risk')), 'scale up / out to protect workloads', CATC['Performance Risk'], 'bolt') +
    kpi('Quick wins', fmt(qw.length), moneyK(dedupSav(qw).total * 12) + '/yr · low effort', '#00B7C3', 'star') + '</div>' +
    '<div class="grid g3">' + card('Savings by rule family', chartBox('rcRule', 320)) + card('Savings by resource type', chartBox('rcType', 320)) + card('Effort vs. confidence (bubble = savings)', chartBox('rcMatrix', 320), { info: 'Bubbles in the bottom-left corner (low effort, high confidence) are the quick wins.' }) + '</div>' +
    '<div class="grid g12">' + card('Actions by priority', chartBox('rcPri', 300)) + card('Grouped view', '<div class="ctl">Group by <select class="sel" id="rcGrp"><option value="rule">Rule family</option><option value="type">Resource type</option><option value="sub">Subscription</option><option value="rg">Resource group</option><option value="env">Environment</option></select></div><div id="rcGroup"></div>') + '</div>' +
    '<div class="whatif no-print"><div><span class="lbl">What-if: selected</span><b id="wiN">0</b> actions</div><div><span class="lbl">Savings</span><b id="wiS" class="pos">$0</b>/yr</div><div><span class="lbl">Run-rate after</span><b id="wiR">–</b> <span id="wiP" class="sm muted"></span></div><div><span class="lbl">Efficiency after</span><b id="wiE">–</b></div><span class="sp"></span>' +
    '<button class="btn sm" data-wi="qw">Select quick wins</button><button class="btn sm" data-wi="hi">High confidence</button><button class="btn sm" data-wi="all">All</button><button class="btn sm" data-wi="none">Clear</button><button class="btn sm pri" data-wi="exp">' + ICON.dl + ' Export selection</button></div>' +
    card('All suggested recommendations', '<div id="rcTable"></div>', { info: 'Tick rows to simulate the impact. Click a row for the full resource analysis.' });
  const byRule = Array.from(groupBy(recs.filter(x => x.sa > 0), x => x.rc).entries()).map(e => ({ k: e[0], n: (D.rules[e[0]] || {}).n || e[0], v: dedupSav(e[1]).total * 12, c: e[1].length })).filter(x => x.v > 0).sort((a, b) => b.v - a.v).slice(0, 12);
  if (byRule.length) ec('rcRule', hbarOpt(byRule.map(x => x.n), byRule.map(x => Math.round(x.v)), PAL, moneyK, { tooltip: { trigger: 'item', formatter: p => { const x = byRule[p.dataIndex]; return '<b>' + esc(x.k) + ' – ' + esc(x.n) + '</b><br>' + money(x.v) + '/yr · ' + x.c + ' recommendations'; } } }), p => { const x = byRule[p.dataIndex]; if (x) { const s = $('#rcTable .dt-s input'); if (s) { s.value = x.k; s.dispatchEvent(new Event('input')); } } });
  else $('#rcRule').innerHTML = emptyMsg('No savings in the current filter.');
  const byType = Array.from(groupBy(recs.filter(x => x.sa > 0), x => x.tn).entries()).map(e => ({ n: e[0], v: dedupSav(e[1]).total * 12, c: e[1].length })).filter(x => x.v > 0).sort((a, b) => b.v - a.v).slice(0, 12);
  if (byType.length) ec('rcType', hbarOpt(byType.map(x => x.n), byType.map(x => Math.round(x.v)), '#0078D4', moneyK), p => { const x = byType[p.dataIndex]; if (x) setFilter('type', [x.n]); });
  else $('#rcType').innerHTML = emptyMsg('No savings in the current filter.');
  const lv = ['Low', 'Medium', 'High'], mx = {};
  recs.forEach(x => { const k = (lv.indexOf(x.eff) < 0 ? 1 : lv.indexOf(x.eff)) + '|' + (lv.indexOf(x.conf) < 0 ? 1 : lv.indexOf(x.conf)); mx[k] = mx[k] || { s: 0, n: 0 }; mx[k].s += x.sa || 0; mx[k].n++; });
  const mpts = Object.keys(mx).map(k => { const p = k.split('|').map(Number); return [p[0], p[1], mx[k].s, mx[k].n]; }), mxS = Math.max(1, Math.max.apply(null, mpts.map(p => p[2]).concat([1])));
  ec('rcMatrix', { grid: { left: 10, right: 20, top: 16, bottom: 30, containLabel: true }, tooltip: { formatter: p => 'Effort <b>' + lv[p.value[0]] + '</b> · confidence <b>' + lv[p.value[1]] + '</b><br>' + p.value[3] + ' recommendations · ' + money(p.value[2]) + '/yr' },
    xAxis: { type: 'category', data: lv.map(x => x + ' effort'), splitLine: { show: true }, axisLabel: { interval: 0 } }, yAxis: { type: 'category', data: lv.map(x => x + ' confidence'), splitLine: { show: true } },
    series: [{ type: 'scatter', data: mpts, symbolSize: v => 18 + 60 * Math.sqrt(v[2] / mxS), itemStyle: { color: p => p.value[0] === 0 && p.value[1] === 2 ? '#13A10E' : p.value[0] === 2 && p.value[1] === 0 ? '#D13438' : '#0078D4', opacity: 0.8 }, label: { show: true, formatter: p => p.value[3] + '\n' + moneyK(p.value[2]), color: '#fff', fontSize: 11, fontWeight: 600 } }] });
  const cats = ['Delete', 'Rightsize', 'Configure', 'Performance Risk'], pris = ['High', 'Medium', 'Low'];
  ec('rcPri', { grid: { left: 8, right: 14, top: 34, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, xAxis: { type: 'category', data: cats }, yAxis: { type: 'value' },
    series: pris.map(p => ({ name: p, type: 'bar', stack: 'p', barMaxWidth: 46, itemStyle: { color: PRIC[p] }, data: cats.map(c => recs.filter(x => x.cat === c && x.pr === p).length) })) }, p => setFilter('act', [p.name]));
  const drawGroup = () => {
    const g = $('#rcGrp').value, keyf = { rule: x => x.rc + ' – ' + ((D.rules[x.rc] || {}).n || ''), type: x => x.tn, sub: x => x.sub, rg: x => x.g, env: x => x.e }[g];
    const rows = Array.from(groupBy(recs, keyf).entries()).map(e => ({ k: e[0], n: e[1].length, sa: dedupSav(e[1]).total * 12, del: e[1].filter(x => x.cat === 'Delete').length, rs: e[1].filter(x => x.cat === 'Rightsize').length, cf: e[1].filter(x => x.cat === 'Configure').length, pf: e[1].filter(x => x.cat === 'Performance Risk').length }));
    DT('rcGroup', { rows, pageSize: 10, sortK: 'sa', sortD: -1, exportName: 'Recommendations by ' + g, cols: [{ k: 'k', n: 'Group' }, { k: 'n', n: 'Recs', num: true }, { k: 'del', n: 'Delete', num: true }, { k: 'rs', n: 'Rightsize', num: true }, { k: 'cf', n: 'Configure', num: true }, { k: 'pf', n: 'Perf.', num: true }, { k: 'sa', n: 'Savings / yr', num: true, f: x => '<b class="pos">' + money(x.sa) + '</b>' }] });
  };
  $('#rcGrp').onchange = drawGroup; drawGroup();
  const tbl = DT('rcTable', { rows: recs, cols: recCols(), sel: true, selSet: SEL, onSel: updWhatIf, pageSize: 50, sortK: 'sa', sortD: -1, q: arg && arg !== 'advisor' ? arg : '', exportName: M.customer + ' - Suggested recommendations', onRow: x => { if (x.R) openRes(x.R); } });
  function updWhatIf() {
    const sel = frecs().filter(x => SEL.has(x)), d = dedupSav(sel), rs = fr(), tot = sum(rs, r => r.tc);
    const del = new Set(sel.filter(x => x.cat === 'Delete' && x.R).map(x => x.R.i)), rsav = new Map();
    sel.filter(x => x.cat !== 'Delete' && x.cat !== 'Performance Risk' && x.R && isNum(x.sm)).forEach(x => rsav.set(x.R.i, (rsav.get(x.R.i) || 0) + x.sm));
    let w = 0, sw = 0;
    rs.forEach(r => { if (del.has(r.i) || !isNum(r.es)) return; let c = r.tc, s = r.es; const v = rsav.get(r.i); if (v && c > 0) { const nc = Math.max(c * 0.1, c - v); s = Math.min(100, s * c / nc); c = nc; } if (c > 0) { w += s * c; sw += c; } });
    $('#wiN').textContent = fmt(sel.length); $('#wiS').textContent = money(d.total * 12);
    $('#wiR').textContent = tot > 0 ? money(tot - d.total) + '/mo' : '–'; $('#wiP').textContent = tot > 0 && d.total > 0 ? '(−' + fmt(100 * d.total / tot, 1) + '%)' : '';
    const before = wscore(rs), after = sw > 0 ? w / sw : before; $('#wiE').innerHTML = isNum(after) ? Math.round(after) + ' <span class="sm muted">(now ' + (isNum(before) ? Math.round(before) : 'n/a') + ')</span>' : '–';
  }
  $$('[data-wi]', body).forEach(b => { b.onclick = () => {
    const a = b.dataset.wi, vis = tbl.view();
    if (a === 'none') SEL.clear(); else if (a === 'all') vis.forEach(x => SEL.add(x)); else if (a === 'qw') vis.filter(isQW).forEach(x => SEL.add(x)); else if (a === 'hi') vis.filter(x => x.conf === 'High' && x.sa > 0).forEach(x => SEL.add(x));
    else if (a === 'exp') { const rows = frecs().filter(x => SEL.has(x)); if (!rows.length) { toast('Select recommendations first'); return; } exportCsv(M.customer + ' - Selected recommendations', recCols().filter(c => c.x !== false), rows); return; }
    tbl.draw(); updWhatIf();
  }; });
  updWhatIf();
}
function renderAdvisor(body, adv) {
  const cats = Array.from(groupBy(adv, a => a.cat || 'Other').entries());
  body.innerHTML = '<div class="kpis">' + kpi('Advisor recommendations', fmt(adv.length), cats.map(e => e[0] + ': ' + e[1].length).join(' · '), '#0078D4', 'bulb') +
    kpi('Advisor annual savings', moneyK(sum(adv, a => a.as)), 'as reported by Azure Advisor', '#13A10E', 'coin') +
    kpi('High impact', fmt(adv.filter(a => a.imp === 'High').length), 'Advisor impact rating', '#D13438', 'bolt') + '</div>' +
    '<div class="grid g2">' + card('By impact and category', chartBox('advImp', 280)) + card('By resource type', chartBox('advType', 280)) + '</div>' +
    card('Azure Advisor recommendations', '<div id="advTable"></div>', { info: 'Native Azure Advisor cost and performance recommendations, shown alongside the rule-based analysis for cross-validation.' });
  const imps = ['High', 'Medium', 'Low'];
  ec('advImp', { grid: { left: 8, right: 14, top: 34, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, xAxis: { type: 'category', data: imps }, yAxis: { type: 'value' },
    series: cats.map((e, i) => ({ name: e[0], type: 'bar', stack: 'a', barMaxWidth: 50, itemStyle: { color: PAL[i % PAL.length] }, data: imps.map(m => e[1].filter(a => a.imp === m).length) })) });
  const bt = Array.from(groupBy(adv, a => (a.typ || '').split('/').pop() || '(subscription)').entries()).map(e => ({ n: e[0], c: e[1].length })).sort((a, b) => b.c - a.c).slice(0, 12);
  if (bt.length) ec('advType', hbarOpt(bt.map(x => x.n), bt.map(x => x.c), '#4F6BED', v => fmt(v))); else $('#advType').innerHTML = emptyMsg();
  DT('advTable', { rows: adv, pageSize: 50, sortK: 'as', sortD: -1, exportName: M.customer + ' - Azure Advisor', cols: [
    { k: 'cat', n: 'Category' }, { k: 'imp', n: 'Impact', f: a => pill(a.imp, PRIC[a.imp] || '#69797E'), v: a => a.imp }, { k: 'res', n: 'Impacted resource', f: a => '<div class="rn"><b>' + esc(a.res) + '</b><small>' + esc(a.typ) + '</small></div>', v: a => a.res },
    { k: 'prob', n: 'Problem', f: a => '<div class="wrap">' + esc(a.prob) + '</div>', v: a => a.prob }, { k: 'sol', n: 'Solution', f: a => '<div class="wrap">' + esc(a.sol) + '</div>', v: a => a.sol },
    { k: 'as', n: 'Annual savings', num: true, f: a => a.as ? '<b class="pos">' + money(a.as) + '</b>' : '–', v: a => a.as }, { k: 'cur', n: 'Current SKU', v: a => a.cur }, { k: 'tgt', n: 'Target SKU', v: a => a.tgt }, { k: 'term', n: 'Term', v: a => a.term },
    { k: 'sub', n: 'Subscription', v: a => D.subs[a.s] }, { k: 'g', n: 'Resource group', v: a => a.g }, { k: 'id', n: 'Resource ID', v: a => a.id, f: a => '<div class="clip muted" style="max-width:220px">' + esc(a.id) + '</div>' }] });
}

// ---------------------------------------------------------------- page: service dashboards
let SVC_CAT = 'All';
const HIST_B = [[0, 5, '0–5%'], [5, 20, '5–20%'], [20, 40, '20–40%'], [40, 60, '40–60%'], [60, 80, '60–80%'], [80, 100.0001, '80–100%'], [100.0001, 1e12, '>100%']];
const HIST_C = ['#D13438', '#F7630C', '#F7A30C', '#13A10E', '#13A10E', '#8764B8', '#5C2D91'];
function renderServices(el, arg) {
  const rs = fr(), byT = groupBy(rs, r => r.t), active = arg ? tByKey[arg] : null;
  const cats = ['All'].concat(Array.from(new Set(Array.from(byT.keys()).map(i => T[i].c))).sort());
  if (cats.indexOf(SVC_CAT) < 0) SVC_CAT = 'All';
  const cards = Array.from(byT.entries()).map(e => ({ t: T[e[0]], l: e[1], cost: sum(e[1], r => r.tc), sv: sum(e[1], r => r.sv), s: wscore(e[1]), u: wavg(e[1], r => r.u, r => r.tc) }))
    .filter(x => SVC_CAT === 'All' || x.t.c === SVC_CAT).sort((a, b) => b.cost - a.cost || b.l.length - a.l.length);
  el.innerHTML = '<div class="tabs" id="svCats">' + cats.map(c => '<button class="tab' + (c === SVC_CAT ? ' on' : '') + '" data-c="' + esc(c) + '">' + esc(c) + '<span class="n">' + (c === 'All' ? byT.size : Array.from(byT.keys()).filter(i => T[i].c === c).length) + '</span></button>').join('') + '</div>' +
    '<div class="tcards">' + cards.map(x => '<a class="tcard' + (active && active === x.t ? ' on' : '') + '" href="#/services/' + encodeURIComponent(x.t.k) + '"><div class="tc-h"><span class="kpi-i" style="--kc:' + scoreColor(x.s) + ';width:34px;height:34px;flex-basis:34px">' + (ICON[CAT_ICON[x.t.c]] || ICON.grid) + '</span><div style="min-width:0"><b>' + esc(x.t.n) + '</b><small>' + esc(x.t.c) + ' · ' + fmt(x.l.length) + ' resources</small></div><span class="sp"></span>' + ring(x.s, 38, 4) + '</div>' +
      '<div class="tc-m"><span>Cost <b>' + moneyK(x.cost) + '</b>/mo</span><span>Util <b>' + pctS(x.u) + '</b></span><span>Save <b class="pos">' + moneyK(x.sv) + '</b></span></div>' + (x.t.a ? '<div><span class="badge auto">auto-discovered metrics</span></div>' : '') + '</a>').join('') + '</div><div id="svDash" style="display:flex;flex-direction:column;gap:16px"></div>';
  $$('#svCats .tab').forEach(b => { b.onclick = () => { SVC_CAT = b.dataset.c; route(true); }; });
  const host = $('#svDash');
  if (active) { typeDash(host, active, byT.get(active.i) || []); setTimeout(() => host.scrollIntoView({ behavior: 'smooth', block: 'start' }), 60); return; }
  host.innerHTML = card('Service landscape – utilization vs. cost', chartBox('svBubble', 440), { info: 'Each bubble is a service: X = cost-weighted utilization, Y = monthly cost (log), size = resources, color = efficiency score. Click a bubble to open the service dashboard.' });
  const pts = cards.filter(x => isNum(x.u) && x.cost > 0);
  if (pts.length) ec('svBubble', { grid: { left: 10, right: 30, top: 20, bottom: 34, containLabel: true }, tooltip: { formatter: p => { const x = pts[p.dataIndex]; return '<b>' + esc(x.t.n) + '</b><br>' + fmt(x.l.length) + ' resources · ' + money(x.cost) + '/month<br>Utilization ' + pctS(x.u, 1) + ' · efficiency ' + (isNum(x.s) ? Math.round(x.s) : 'n/a') + '<br>Savings ' + money(x.sv) + '/yr'; } },
    xAxis: { type: 'value', name: 'Utilization (cost-weighted P95 %)', nameLocation: 'middle', nameGap: 26, min: 0, max: v => Math.max(100, Math.ceil(v.max / 10) * 10) }, yAxis: { type: 'log', min: 1, axisLabel: axisMoney },
    series: [{ type: 'scatter', data: pts.map(x => [Math.min(x.u, 150), Math.max(1, x.cost), x.l.length]), symbolSize: v => 14 + 7 * Math.sqrt(v[2]), itemStyle: { color: p => scoreColor(pts[p.dataIndex].s), opacity: 0.85, borderColor: '#fff', borderWidth: 1 }, label: { show: true, position: 'right', formatter: p => pts[p.dataIndex].t.n, color: css('--muted'), fontSize: 11 },
      markLine: { silent: true, symbol: 'none', lineStyle: { type: 'dashed', color: '#8A8886' }, data: [{ xAxis: TH.low }, { xAxis: TH.target }] } }] }, p => { const x = pts[p.dataIndex]; if (x) go('services/' + encodeURIComponent(x.t.k)); });
  else $('#svBubble').innerHTML = emptyMsg('Select a service above.');
}
function numOptions(t, rs) {
  const o = [['core:u', 'Utilization % (primary, P95)'], ['core:c', 'Monthly cost'], ['core:es', 'Efficiency score'], ['core:ag', 'Age (days)']];
  (t.mc || []).forEach(c => { if (rs.some(r => isNum(fv(r, c)))) o.push(['m:' + c, mLabel(t, c)]); });
  (t.fc || []).forEach(c => { if (rs.some(r => isNum(fv(r, c)))) o.push(['f:' + c, c]); });
  return o;
}
function numVal(r, id) { const k = id.slice(id.indexOf(':') + 1); if (id.indexOf('core:') === 0) return r[k]; return fv(r, k); }
function typeDash(host, t, rs) {
  const tot = sum(rs, r => r.tc), sv = sum(rs, r => r.sv), score = wscore(rs), us = rs.filter(r => isNum(r.u)).map(r => r.u), by = SG.map((g, i) => rs.filter(r => r.sgi === i).length);
  const pmLbl = t.pm ? (t.ml[t.pm] || t.pm) : 'primary metric', hwKeys = Object.keys(t.hw || {});
  const opts = numOptions(t, rs);
  const defX = t.k === 'VirtualMachines' && opts.some(o => o[0] === 'm:CPU_P95') ? 'm:CPU_P95' : 'core:u';
  const defY = t.k === 'VirtualMachines' && opts.some(o => o[0] === 'm:MemUsedPct_Max') ? 'm:MemUsedPct_Max' : 'core:c';
  host.innerHTML = '<div class="tdh"><div><div class="eyebrow muted">' + esc(t.c) + ' · ' + (t.a ? 'auto-discovered Azure Monitor metrics' : 'dedicated right-sizing analysis') + '</div><h2>' + esc(t.n) + '</h2></div><span class="sp"></span><a class="btn" href="#/explorer/' + encodeURIComponent(t.k) + '">' + ICON.table + ' Open in explorer</a></div>' +
    '<div class="kpis">' + kpi('Resources', fmt(rs.length), fmt(by[0]) + ' idle · ' + fmt(by[1]) + ' over · ' + fmt(by[3]) + ' under', '#0078D4', CAT_ICON[t.c] || 'grid') +
    kpi('Monthly cost', moneyK(tot), 'run-rate · ' + esc(M.costType), '#4F6BED', 'coin') + kpi('Savings', moneyK(sv), 'per year', '#13A10E', 'star') +
    kpi('Median utilization', pctS(pctl(us, 50), 1), 'P95 across resources: ' + pctS(pctl(us, 95), 1), utilColor(pctl(us, 50)), 'chart') +
    kpi('Efficiency score', isNum(score) ? Math.round(score) + '<span class="sm muted"> / 100</span>' : 'n/a', 'cost-weighted', scoreColor(score), 'check') + '</div>' +
    (t.k === 'VirtualMachines' ? '<div class="grid g3">' + capacityHtml(rs) + card('Power state & OS', chartBox('vmPower', 260)) + card('Effective discount vs. retail PAYG', chartBox('vmDisc', 260), { info: 'Billed cost vs. retail pay-as-you-go for the hours run. ~0% = pure PAYG (reservation / savings plan candidates).' }) + '</div>' : '') +
    '<div class="grid g3">' + card('Utilization distribution', chartBox('tdHist', 280), { info: 'Resources by utilization of their provisioned SKU (primary metric, P95).' }) + card('Utilization status', chartBox('tdStatus', 280)) + card('Top 12 by monthly cost', chartBox('tdTop', 280)) + '</div>' +
    '<div class="grid g2">' + card('Metric explorer', '<div class="ctl">X <select class="sel" id="tdX">' + opts.map(o => '<option value="' + esc(o[0]) + '"' + (o[0] === defX ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') + '</select> Y <select class="sel" id="tdY">' + opts.map(o => '<option value="' + esc(o[0]) + '"' + (o[0] === defY ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') + '</select><label><input type="checkbox" id="tdLog"' + (defY === 'core:c' ? ' checked' : '') + '> log Y</label></div>' + chartBox('tdScatter', 340), { info: 'Plot any two collected metrics against each other. Bubble size = monthly cost; color = utilization status. Click a point for details.' }) +
    card('Daily trend – ' + pmLbl + ' across resources', chartBox('tdDaily', 372), { info: 'Fleet average, P95 across resources and maximum per day for the primary metric.' }) + '</div>' +
    '<div class="grid g2">' + card('Hour-of-week pattern (' + esc(M.tz) + ')', hwKeys.length ? '<div class="ctl">Metric <select class="sel" id="tdHwM">' + hwKeys.map(k => '<option value="' + esc(k) + '"' + (k === t.pm ? ' selected' : '') + '>' + esc(t.ml[k] || k) + ' (' + fmt(t.hw[k].r) + ' resources)</option>').join('') + '</select></div>' + chartBox('tdHw', 290) : emptyMsg('No hourly pattern collected for this service.')) +
    card('Metric statistics across resources', '<div id="tdStats"></div>') + '</div>' +
    card('Resources', '<div id="tdTable"></div>');
  if (t.k === 'VirtualMachines') {
    const ps = Array.from(groupBy(rs, r => String(fv(r, 'PowerState') || 'unknown') + ' · ' + String(fv(r, 'OsType') || '?')).entries()).map((e, i) => ({ name: e[0], value: e[1].length, itemStyle: { color: PAL[i % PAL.length] } }));
    ec('vmPower', donutOpt(ps, fmt(rs.length), 'VMs'));
    const ed = rs.filter(r => isNum(r.ed)), bk = [[-1e9, 5, '< 5% (PAYG)'], [5, 20, '5–20%'], [20, 40, '20–40%'], [40, 60, '40–60%'], [60, 1e9, '> 60%']];
    if (ed.length) ec('vmDisc', { grid: { left: 8, right: 10, top: 20, bottom: 6, containLabel: true }, tooltip: { trigger: 'axis' }, xAxis: { type: 'category', data: bk.map(b => b[2]) }, yAxis: { type: 'value', name: 'VMs' }, series: [{ type: 'bar', barMaxWidth: 46, data: bk.map((b, i) => ({ value: ed.filter(r => r.ed >= b[0] && r.ed < b[1]).length, itemStyle: { color: i === 0 ? '#F7630C' : '#13A10E', borderRadius: [6, 6, 0, 0] } })), label: { show: true, position: 'top', color: css('--muted') } }] });
    else $('#vmDisc').innerHTML = emptyMsg('Requires cost and retail price data for running VMs.');
  }
  ec('tdHist', { grid: { left: 8, right: 10, top: 20, bottom: 6, containLabel: true }, tooltip: { trigger: 'axis' }, xAxis: { type: 'category', data: HIST_B.map(b => b[2]) }, yAxis: { type: 'value', name: 'resources' },
    series: [{ type: 'bar', barMaxWidth: 46, data: HIST_B.map((b, i) => ({ value: us.filter(u => u >= b[0] && u < b[1]).length, itemStyle: { color: HIST_C[i], borderRadius: [6, 6, 0, 0] } })), label: { show: true, position: 'top', color: css('--muted') } }] });
  ec('tdStatus', donutOpt(SG.map((g, i) => ({ name: g.n, value: by[i], itemStyle: { color: g.c } })).filter(x => x.value > 0), fmt(rs.length), 'resources'), p => { const g = SG.find(x => x.n === p.name); if (g) setFilter('sg', [g.n]); });
  const top = rs.filter(r => r.tc > 0).sort((a, b) => b.tc - a.tc).slice(0, 12);
  if (top.length) ec('tdTop', hbarOpt(top.map(r => r.n), top.map(r => Math.round(r.tc)), top.map(r => STC(r.st)), moneyK, { tooltip: { trigger: 'item', formatter: p => { const r = top[p.dataIndex]; return '<b>' + esc(r.n) + '</b><br>' + money(r.tc) + '/month · utilization ' + pctS(r.u, 1) + '<br>' + esc(r.st); } } }), p => openRes(top[p.dataIndex]));
  else $('#tdTop').innerHTML = emptyMsg('No cost data.');
  const drawScatter = () => {
    const x = $('#tdX').value, y = $('#tdY').value, logY = $('#tdLog').checked, maxC = Math.max(1, Math.max.apply(null, rs.map(r => r.tc).concat([1])));
    const pts = rs.map(r => [numVal(r, x), numVal(r, y), r.tc, r.i]).filter(p => isNum(p[0]) && isNum(p[1]) && (!logY || p[1] > 0));
    const lx = (opts.find(o => o[0] === x) || [0, x])[1], ly = (opts.find(o => o[0] === y) || [0, y])[1];
    const pct = id => /core:u|Pct|CPU|Percent/i.test(id);
    ec('tdScatter', { grid: { left: 10, right: 24, top: 20, bottom: 36, containLabel: true }, legend: { show: false },
      tooltip: { formatter: p => { const r = R[p.value[3]]; return '<b>' + esc(r.n) + '</b><br>' + esc(lx) + ': ' + fmtA(p.value[0]) + '<br>' + esc(ly) + ': ' + fmtA(p.value[1]) + '<br>' + money(r.tc) + '/month · ' + esc(r.st); } },
      xAxis: { type: 'value', name: lx, nameLocation: 'middle', nameGap: 26, axisLabel: { formatter: v => fmtA(v) } }, yAxis: { type: logY ? 'log' : 'value', name: ly, axisLabel: { formatter: v => y === 'core:c' ? moneyK(v) : fmtA(v) } },
      series: SG.map((g, gi) => ({ name: g.n, type: 'scatter', data: pts.filter(p => R[p[3]].sgi === gi), symbolSize: v => 8 + 26 * Math.sqrt((v[2] || 0) / maxC), itemStyle: { color: g.c, opacity: 0.8, borderColor: '#fff', borderWidth: 0.6 },
        markLine: gi === 0 && pct(x) ? { silent: true, symbol: 'none', lineStyle: { type: 'dashed', color: '#8A8886' }, label: { formatter: p => p.value + '%' }, data: [{ xAxis: TH.low }, { xAxis: TH.high }] } : undefined })) }, p => openRes(p.value[3]));
  };
  ['tdX', 'tdY', 'tdLog'].forEach(id => { $('#' + id).onchange = () => { CH = CH.filter(c => { if (c.getDom().id === 'tdScatter') { c.dispose(); return false; } return true; }); drawScatter(); }; });
  drawScatter();
  const dr = rs.filter(r => r.d && r.pm === t.pm);
  if (dr.length && DATES.length) {
    const av = [], p95 = [], mx = [];
    DATES.forEach((d, i) => { const v = dr.map(r => r.d[i]).filter(isNum); av.push(v.length ? avg(v) : null); p95.push(v.length ? pctl(v, 95) : null); mx.push(v.length ? Math.max.apply(null, v) : null); });
    ec('tdDaily', { grid: { left: 8, right: 16, top: 34, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', valueFormatter: v => fmtA(v) }, xAxis: { type: 'category', data: DATES.map(shortDate) }, yAxis: { type: 'value', axisLabel: { formatter: v => fmtA(v) } },
      series: [{ name: 'Average across ' + dr.length + ' resources', type: 'line', data: av, smooth: true, showSymbol: false, areaStyle: { opacity: 0.15 }, lineStyle: { width: 2.4 }, itemStyle: { color: '#0078D4' }, markArea: weekendArea() },
        { name: 'P95 across resources', type: 'line', data: p95, smooth: true, showSymbol: false, lineStyle: { width: 1.6, type: 'dashed' }, itemStyle: { color: '#F7630C' } },
        { name: 'Maximum', type: 'line', data: mx, smooth: true, showSymbol: false, lineStyle: { width: 1.2, type: 'dotted' }, itemStyle: { color: '#8764B8' } }] });
  } else $('#tdDaily').innerHTML = emptyMsg('No daily series for this service.');
  if (hwKeys.length) { const draw = () => { CH = CH.filter(c => { if (c.getDom().id === 'tdHw') { c.dispose(); return false; } return true; }); heatmap('tdHw', t.hw[$('#tdHwM').value].z, t.ml[$('#tdHwM').value] || $('#tdHwM').value); }; $('#tdHwM').onchange = draw; draw(); }
  const stats = (t.mc || []).map(c => { const v = rs.map(r => fv(r, c)).filter(isNum); return v.length ? { c, l: mLabel(t, c), n: v.length, min: Math.min.apply(null, v), avg: avg(v), med: pctl(v, 50), p95: pctl(v, 95), max: Math.max.apply(null, v) } : null; }).filter(Boolean);
  DT('tdStats', { rows: stats, pageSize: 10, maxH: 300, exportName: t.n + ' - metric statistics', cols: [{ k: 'l', n: 'Metric', f: s => '<b>' + esc(s.l) + '</b><div class="muted sm">' + esc(s.c) + '</div>', v: s => s.l }, { k: 'n', n: 'Res.', num: true }, { k: 'min', n: 'Min', num: true, f: s => fmtA(s.min) }, { k: 'avg', n: 'Avg', num: true, f: s => fmtA(s.avg) }, { k: 'med', n: 'Median', num: true, f: s => fmtA(s.med) }, { k: 'p95', n: 'P95', num: true, f: s => fmtA(s.p95) }, { k: 'max', n: 'Max', num: true, f: s => fmtA(s.max) }] });
  const keyM = (t.mc || []).filter(c => /_P95$/.test(c)).slice(0, 4).concat((t.mc || []).filter(c => /_Total$/.test(c)).slice(0, 2));
  DT('tdTable', { rows: rs, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - ' + t.n, onRow: r => openRes(r),
    cols: resCols(['name', 'rg', 'sku', 'st', 'u', 'es', 'c', 'sv', 'trend', 'rec']).concat(keyM.map(c => ({ k: 'm:' + c, n: mLabel(t, c), num: true, v: r => fv(r, c), f: r => fmtA(fv(r, c)) })), resCols(['sub', 'env', 'loc', 'pv', 'us'])) });
}
function heatmap(id, z, label) {
  const data = [], vals = (z || []).filter(isNum);
  for (let d = 0; d < 7; d++) for (let h = 0; h < 24; h++) { const v = z[d * 24 + h]; data.push([h, d, isNum(v) ? v : '-']); }
  const biz = h => h >= M.bizStart && h < M.bizEnd;
  ec(id, { grid: { left: 8, right: 10, top: 10, bottom: 54, containLabel: true }, tooltip: { formatter: p => DAYN[p.value[1]] + ' ' + String(p.value[0]).padStart(2, '0') + ':00' + (biz(p.value[0]) && !WEEKEND.has(DAYN[p.value[1]]) ? ' (business hours)' : '') + '<br>' + esc(label) + ': <b>' + fmtA(p.value[2]) + '</b>' },
    xAxis: { type: 'category', data: Array.from({ length: 24 }, (x, h) => String(h).padStart(2, '0')), splitArea: { show: false }, axisLabel: { formatter: v => biz(+v) ? '{b|' + v + '}' : v, rich: { b: { fontWeight: 700, color: css('--brand2') } } } },
    yAxis: { type: 'category', data: DAYN.map(d => d.slice(0, 3)), inverse: true, axisLabel: { formatter: v => WEEKEND.has(DAYN.find(d => d.slice(0, 3) === v)) ? '{w|' + v + '}' : v, rich: { w: { color: '#8764B8', fontWeight: 700 } } } },
    visualMap: { min: vals.length ? Math.min.apply(null, vals) : 0, max: vals.length ? Math.max.apply(null, vals) : 1, calculable: true, orient: 'horizontal', left: 'center', bottom: 0, itemHeight: 160, inRange: { color: ['#EAF3FC', '#9ECAE1', '#3A8DD0', '#08519C', '#4B1F86'] }, formatter: v => fmtA(v) },
    series: [{ type: 'heatmap', data, itemStyle: { borderColor: css('--card'), borderWidth: 1.5, borderRadius: 3 }, emphasis: { itemStyle: { borderColor: '#000', borderWidth: 1 } } }] });
}

// ---------------------------------------------------------------- page: usage patterns
const STOPPABLE = new Set(['VirtualMachines', 'VMScaleSets', 'AKSNodePools', 'AzureVirtualDesktop', 'SynapseSqlPools', 'DataExplorerClusters', 'PowerBIFabricCapacities', 'MLComputes', 'SqlManagedInstances', 'AnalysisServices', 'ContainerInstances', 'MySQLFlexibleServers', 'PostgreSQLFlexibleServers', 'AppServicePlans']);
let PAT_T = '', PAT_M = '';
function renderPatterns(el) {
  const rs = fr(), withHw = T.filter(t => Object.keys(t.hw || {}).length && rs.some(r => r.T === t));
  if (!withHw.some(t => t.k === PAT_T)) { PAT_T = (withHw.find(t => t.k === 'VirtualMachines') || withHw[0] || {}).k || ''; PAT_M = ''; }
  const t = tByKey[PAT_T];
  if (t && !(PAT_M in t.hw)) PAT_M = (t.pm in t.hw) ? t.pm : Object.keys(t.hw)[0];
  const pt = rs.filter(r => r.pat && isNum(r.pat[2]) && isNum(r.pat[3]));
  const bizH = (7 - WEEKEND.size) * (M.bizEnd - M.bizStart), offFrac = Math.max(0, (168 - bizH) / 168);
  const cand = pt.filter(r => r.pat[2] > 0 && (r.pat[3] / r.pat[2] < 0.35 || (isNum(r.pat[1]) && r.pat[1] >= 60)) && r.sgi !== 0 && r.sgi !== 4)
    .map(r => ({ r, ratio: 100 * r.pat[3] / r.pat[2], pot: STOPPABLE.has(r.T.k) && r.tc > 0 ? r.tc * offFrac * 0.9 * 12 : null })).sort((a, b) => (b.pot || 0) - (a.pot || 0) || b.r.tc - a.r.tc);
  el.innerHTML = '<div class="kpis">' + kpi('Resources with hourly patterns', fmt(pt.length), 'business hours ' + M.bizStart + ':00–' + M.bizEnd + ':00 · weekend ' + esc((M.weekend || []).join('/')), '#0078D4', 'clock') +
    kpi('Off-hours vs. business hours', pctS(avg(pt.filter(r => r.pat[2] > 0).map(r => 100 * r.pat[3] / r.pat[2]))), 'average load outside business hours', '#8764B8', 'pause') +
    kpi('Weekend vs. weekday', pctS(avg(pt.filter(r => isNum(r.pat[4]) && r.pat[4] > 0 && isNum(r.pat[5])).map(r => 100 * r.pat[5] / r.pat[4]))), 'average weekend load', '#00B7C3', 'chart') +
    kpi('Mostly idle hours', fmt(pt.filter(r => isNum(r.pat[1]) && r.pat[1] >= 60).length), 'resources idle ≥ 60% of hours', '#D13438', 'trash') +
    kpi('Schedule candidates', fmt(cand.length), moneyK(sum(cand, c => c.pot)) + '/yr indicative (stoppable services)', '#13A10E', 'bolt') + '</div>' +
    '<div class="grid g21">' + card('Hour-of-week heatmap (' + esc(M.tz) + ')', withHw.length ? '<div class="ctl">Service <select class="sel" id="paT">' + withHw.map(x => '<option value="' + esc(x.k) + '"' + (x.k === PAT_T ? ' selected' : '') + '>' + esc(x.n) + '</option>').join('') + '</select> Metric <select class="sel" id="paM">' + (t ? Object.keys(t.hw).map(k => '<option value="' + esc(k) + '"' + (k === PAT_M ? ' selected' : '') + '>' + esc(t.ml[k] || k) + ' (' + fmt(t.hw[k].r) + ' resources)</option>').join('') : '') + '</select><span class="muted sm">Bold hours = business hours · purple days = weekend</span></div>' + chartBox('paHeat', 330) : emptyMsg('No hourly usage patterns collected.'), { info: 'Average of the hourly values by local day and hour across all resources of the service (whole window; not affected by the resource filters).' }) +
    card('Peak hour of each resource', chartBox('paPeak', 360), { info: 'Distribution of the hour with the highest load of the primary metric (local time).' }) + '</div>' +
    '<div class="grid g2">' + card('Off-hours load as % of business hours – by service', chartBox('paOff', 340), { info: 'Low values = load concentrated in business hours → start/stop schedules or auto-scale.' }) + card('Weekend load as % of weekday – by service', chartBox('paWk', 340)) + '</div>' +
    '<div class="grid g2">' + card('Idle hours distribution', chartBox('paIdle', 300), { info: 'Share of hours in which the primary metric was idle, per resource.' }) + card('Daily utilization index by service (100 = window average)', chartBox('paIdx', 300), { info: 'Fleet average of the primary metric per day, normalized per service, to compare day-to-day dynamics.' }) + '</div>' +
    card('Schedule & auto-pause candidates', '<div id="paTable"></div>', { info: 'Resources (not idle, not stopped) whose off-hours load is below 35% of business hours or idle ≥ 60% of hours. Indicative savings assume stopping ' + fmt(offFrac * 100) + '% of hours at 90% efficiency for stoppable services; not included in the savings totals.' });
  if (t && PAT_M) heatmap('paHeat', t.hw[PAT_M].z, t.ml[PAT_M] || PAT_M);
  const ta = $('#paT'); if (ta) ta.onchange = () => { PAT_T = ta.value; PAT_M = ''; route(true); };
  const ma = $('#paM'); if (ma) ma.onchange = () => { PAT_M = ma.value; route(true); };
  const ph = new Array(24).fill(0); pt.forEach(r => { if (isNum(r.pat[6]) && r.pat[6] >= 0 && r.pat[6] < 24) ph[r.pat[6]]++; });
  ec('paPeak', { grid: { left: 8, right: 10, top: 16, bottom: 6, containLabel: true }, tooltip: { trigger: 'axis' }, xAxis: { type: 'category', data: ph.map((x, h) => String(h).padStart(2, '0')) }, yAxis: { type: 'value', name: 'resources' },
    series: [{ type: 'bar', barMaxWidth: 18, data: ph.map((v, h) => ({ value: v, itemStyle: { color: h >= M.bizStart && h < M.bizEnd ? '#0078D4' : '#8764B8', borderRadius: [4, 4, 0, 0] } })) }] });
  const byT2 = Array.from(groupBy(pt, r => r.t).entries()).map(e => { const l = e[1].filter(r => r.pat[2] > 0); const lw = e[1].filter(r => isNum(r.pat[4]) && r.pat[4] > 0 && isNum(r.pat[5])); return { t: T[e[0]], off: l.length ? avg(l.map(r => 100 * r.pat[3] / r.pat[2])) : null, wk: lw.length ? avg(lw.map(r => 100 * r.pat[5] / r.pat[4])) : null, n: e[1].length }; });
  const offL = byT2.filter(x => isNum(x.off)).sort((a, b) => a.off - b.off).slice(0, 16), wkL = byT2.filter(x => isNum(x.wk)).sort((a, b) => a.wk - b.wk).slice(0, 16);
  const pc = v => v < 35 ? '#13A10E' : v < 70 ? '#F7A30C' : '#8A8886';
  const capBar = (id, list, f) => { if (!list.length) { $('#' + id).innerHTML = emptyMsg(); return; } const o = hbarOpt(list.map(x => x.t.n), list.map(x => Math.min(200, Math.round(x[f]))), list.map(x => pc(x[f])), v => fmt(v) + '%', { xAxis: { type: 'value', max: 200, axisLabel: { formatter: v => v + '%' } }, tooltip: { trigger: 'item', formatter: p => '<b>' + esc(list[p.dataIndex].t.n) + '</b><br>' + fmt(list[p.dataIndex][f], 1) + '% (' + list[p.dataIndex].n + ' resources)' } }); o.series[0].label.formatter = p => fmt(list[p.dataIndex][f]) + '%'; ec(id, o); };
  capBar('paOff', offL, 'off'); capBar('paWk', wkL, 'wk');
  const ib = [[0, 20, '0–20%'], [20, 40, '20–40%'], [40, 60, '40–60%'], [60, 80, '60–80%'], [80, 101, '80–100%']], idl = pt.filter(r => isNum(r.pat[1]));
  ec('paIdle', { grid: { left: 8, right: 10, top: 16, bottom: 6, containLabel: true }, tooltip: { trigger: 'axis' }, xAxis: { type: 'category', data: ib.map(b => b[2] + ' idle') }, yAxis: { type: 'value', name: 'resources' },
    series: [{ type: 'bar', barMaxWidth: 50, data: ib.map((b, i) => ({ value: idl.filter(r => r.pat[1] >= b[0] && r.pat[1] < b[1]).length, itemStyle: { color: ['#13A10E', '#7FBA00', '#F7A30C', '#F7630C', '#D13438'][i], borderRadius: [6, 6, 0, 0] } })), label: { show: true, position: 'top', color: css('--muted') } }] });
  const idxS = Array.from(groupBy(rs.filter(r => r.d && r.pm === r.T.pm), r => r.t).entries()).map(e => { const t2 = T[e[0]]; const av = DATES.map((d, i) => avg(e[1].map(r => r.d[i]))); const m = avg(av); return { n: t2.n, c: e[1].length, v: m ? av.map(x => isNum(x) ? Math.round(1000 * x / m) / 10 : null) : null }; }).filter(x => x.v).sort((a, b) => b.c - a.c).slice(0, 10);
  if (idxS.length) ec('paIdx', { grid: { left: 8, right: 14, top: 40, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis' }, xAxis: { type: 'category', data: DATES.map(shortDate) }, yAxis: { type: 'value', scale: true }, series: idxS.map((x, i) => ({ name: x.n, type: 'line', smooth: true, showSymbol: false, data: x.v, lineStyle: { width: 1.8 }, itemStyle: { color: PAL[i % PAL.length] }, markArea: i === 0 ? weekendArea() : undefined })) });
  else $('#paIdx').innerHTML = emptyMsg();
  DT('paTable', { rows: cand, pageSize: 25, exportName: M.customer + ' - Schedule candidates', onRow: c => openRes(c.r), sortK: 'pot', sortD: -1, cols: [
    { k: 'n', n: 'Resource', f: c => RCOL.name.f(c.r), v: c => c.r.n }, { k: 'e', n: 'Environment', v: c => c.r.e }, { k: 'biz', n: 'Business-hours avg', num: true, v: c => c.r.pat[2], f: c => fmtA(c.r.pat[2]) }, { k: 'off', n: 'Off-hours avg', num: true, v: c => c.r.pat[3], f: c => fmtA(c.r.pat[3]) },
    { k: 'ratio', n: 'Off / business', num: true, f: c => pctS(c.ratio), v: c => c.ratio }, { k: 'idle', n: 'Idle hours', num: true, v: c => c.r.pat[1], f: c => pctS(c.r.pat[1]) }, { k: 'peak', n: 'Peak', v: c => (c.r.pat[7] || '') + ' ' + (isNum(c.r.pat[6]) ? String(c.r.pat[6]).padStart(2, '0') + ':00' : '') },
    { k: 'c', n: 'Monthly cost', num: true, v: c => c.r.c, f: c => money(c.r.c) }, { k: 'pot', n: 'Indicative schedule savings / yr', num: true, v: c => c.pot, f: c => c.pot ? '<b class="pos">' + money(c.pot) + '</b>' : '<span class="muted">n/a</span>' }, { k: 'rec', n: 'Current recommendation', f: c => RCOL.rec.f(c.r), v: c => RCOL.rec.v(c.r) }] });
}

// ---------------------------------------------------------------- page: commitments & Log Analytics / Sentinel
let LA_SEL = -1;
function renderCommit(el) {
  const la = (D.la || []).map((w, i) => Object.assign(w, { i })).filter(w => passSR(D.subs[w.s], w.g, w.n));
  const laSorted = la.slice().sort((a, b) => (b.avg || 0) - (a.avg || 0));
  if (!laSorted.some(w => w.i === LA_SEL)) LA_SEL = laSorted.length ? laSorted[0].i : -1;
  const vmc = frecs().filter(x => x.rc === 'COMMIT-RI'), commitWs = la.filter(w => w.com || w.senc);
  const ri = D.ri || [], sp = D.sp || [];
  el.innerHTML = '<div class="kpis">' + kpi('Log Analytics workspaces', fmt(la.length), fmt(la.filter(w => w.sen).length) + ' with Microsoft Sentinel', '#0078D4', 'shield') +
    kpi('Average ingestion', fmtA(sum(la, w => w.avg)) + ' GB/day', 'P95 day ' + fmtA(sum(la, w => w.p95)) + ' GB (sum of workspaces)', '#4F6BED', 'chart') +
    kpi('Commitment tiers', fmt(commitWs.length), commitWs.length ? 'avg utilization ' + pctS(avg(commitWs.map(w => w.cu))) : 'all pay-as-you-go', '#8764B8', 'layers') +
    kpi('LA / Sentinel savings', moneyK(sum(la, w => w.sav) * 12), 'per year from tier optimization', '#13A10E', 'coin') +
    kpi('Reservations · savings plans', fmt(ri.length) + ' · ' + fmt(sp.length), 'in scope', '#00B7C3', 'shield') +
    kpi('VM commitment candidates', fmt(vmc.length), moneyK(dedupSav(vmc).total * 12) + '/yr indicative', '#C19C00', 'cpu') + '</div>' +
    (laSorted.length ? '<div class="grid g21">' + card('Daily ingestion vs. commitment', '<div class="ctl">Workspace <select class="sel" id="laSel">' + laSorted.map(w => '<option value="' + w.i + '"' + (w.i === LA_SEL ? ' selected' : '') + '>' + esc(w.n) + ' – ' + fmtA(w.avg) + ' GB/day</option>').join('') + '</select></div>' + chartBox('laDaily', 360), { info: 'Billable GB per day by table (stacked) with the Log Analytics and Sentinel commitment levels.' }) +
      card('Workspace analysis', '<div id="laCard"></div>') + '</div>' +
      '<div class="grid g2">' + card('Ingestion by workspace – average, P95, max (GB/day)', chartBox('laBars', Math.max(260, 30 * Math.min(15, laSorted.length) + 60))) + card('Current vs. optimal monthly cost', chartBox('laCost', Math.max(260, 30 * Math.min(15, laSorted.length) + 60)), { info: 'Estimated monthly cost at the current pricing tier vs. the cheapest tier for the observed daily series (Log Analytics + Sentinel).' }) + '</div>' +
      card('Log Analytics & Sentinel – tier analysis for every workspace', '<div id="laTable"></div>') : card('Log Analytics & Sentinel', emptyMsg('No Log Analytics workspaces in the current filter.'))) +
    card('Steady VMs – reservation / savings plan candidates', '<div id="vmcTable"></div>', { info: 'VMs running ≥ ' + TH.steady + '% of hours, right-sized and billed close to pay-as-you-go.' }) +
    '<div class="grid g2">' + card('Reservations – utilization', '<div id="riTable"></div>') + card('Savings plans – utilization', '<div id="spTable"></div>') + '</div>';
  if (laSorted.length) {
    $('#laSel').onchange = e => { LA_SEL = +e.target.value; route(true); };
    const w = (D.la || [])[LA_SEL];
    if (w) {
      const ser = (w.tb || []).map((tb, i) => ({ name: tb.n, type: 'bar', stack: 't', barMaxWidth: 26, data: tb.v, itemStyle: { color: tb.n === 'Other tables' ? '#A0AEC0' : PAL[i % PAL.length] }, emphasis: { focus: 'series' } }));
      const ml = []; if (w.com) ml.push({ yAxis: w.com, name: 'LA commitment', label: { formatter: 'Log Analytics commitment ' + w.com + ' GB/day', position: 'insideEndTop', fontWeight: 600 }, lineStyle: { color: '#D13438', type: 'dashed', width: 2 } });
      if (w.senc) ml.push({ yAxis: w.senc, name: 'Sentinel commitment', label: { formatter: 'Sentinel commitment ' + w.senc + ' GB/day', position: 'insideEndTop', fontWeight: 600 }, lineStyle: { color: '#8764B8', type: 'dashed', width: 2 } });
      ser.push({ name: 'Total billable GB', type: 'line', data: w.d || [], smooth: true, symbolSize: 5, itemStyle: { color: css('--text') }, lineStyle: { width: 2 }, markLine: ml.length ? { symbol: 'none', data: ml } : undefined, markArea: weekendArea() });
      ec('laDaily', { grid: { left: 8, right: 16, top: 44, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', valueFormatter: v => fmtA(v) + ' GB' }, xAxis: { type: 'category', data: DATES.map(shortDate) }, yAxis: { type: 'value', name: 'GB', max: v => Math.max(v.max, w.com || 0, w.senc || 0) * 1.08 }, series: ser });
      const row = (k, v) => '<div><span>' + esc(k) + '</span><b>' + v + '</b></div>';
      $('#laCard').innerHTML = '<div class="rc-h" style="margin-bottom:8px">' + pill(w.st || 'n/a', STC(w.st)) + (w.sen ? ' ' + pill('Sentinel', '#8764B8') : '') + '</div><div class="facts" style="grid-template-columns:1fr 1fr">' +
        row('Pricing tier', esc(w.tier)) + row('Sentinel', w.sen ? esc(w.senp || 'enabled') + (w.senc ? ' · ' + w.senc + ' GB/day' : '') : 'not enabled') + row('Average / day', fmtA(w.avg) + ' GB') + row('P95 / max day', fmtA(w.p95) + ' / ' + fmtA(w.max) + ' GB') +
        row('Weekday / weekend', fmtA(w.wd) + ' / ' + fmtA(w.we) + ' GB') + row('Commitment utilization', pctS(w.cu)) + row('Days below / above', (isNum(w.below) ? w.below : '–') + ' / ' + (isNum(w.above) ? w.above : '–')) + row('Retention', isNum(w.ret) ? w.ret + ' days' : '–') +
        row('Current LA est.', money(w.cla) + '/mo') + row('Optimal LA', esc(w.ola || '–') + ' · ' + money(w.olac)) + row('Current Sentinel est.', money(w.cse) + '/mo') + row('Optimal Sentinel', esc(w.ose || '–') + ' · ' + money(w.osec)) + '</div>' +
        '<div class="summary" style="margin-top:10px">Estimated saving: <b class="pos">' + money(w.sav) + '/month</b> (' + money((w.sav || 0) * 12) + '/year). <span class="muted sm">Prices: ' + esc(w.src || 'retail') + '</span></div>' + (w.ri >= 0 && w.ri != null ? '<div style="margin-top:10px"><button class="btn sm" id="laOpen">' + ICON.eye + ' Open workspace details</button></div>' : '');
      const lo = $('#laOpen'); if (lo) lo.onclick = () => openRes(w.ri);
    }
    const top = laSorted.slice(0, 15);
    ec('laBars', { grid: { left: 8, right: 16, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => fmtA(v) + ' GB' }, xAxis: { type: 'value' }, yAxis: { type: 'category', data: top.map(w => w.n), inverse: true, axisLabel: { width: 170, overflow: 'truncate' } },
      series: [['Average', 'avg', '#0078D4'], ['P95', 'p95', '#F7630C'], ['Max', 'max', '#8764B8']].map(s => ({ name: s[0], type: 'bar', barMaxWidth: 10, data: top.map(w => w[s[1]]), itemStyle: { color: s[2], borderRadius: [0, 4, 4, 0] } })) }, p => { LA_SEL = top[p.dataIndex].i; route(true); });
    ec('laCost', { grid: { left: 8, right: 16, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => money(v) }, xAxis: { type: 'value', axisLabel: axisMoney }, yAxis: { type: 'category', data: top.map(w => w.n), inverse: true, axisLabel: { width: 170, overflow: 'truncate' } },
      series: [{ name: 'Current (LA + Sentinel)', type: 'bar', barMaxWidth: 12, data: top.map(w => (w.cla || 0) + (w.cse || 0)), itemStyle: { color: '#A0AEC0', borderRadius: [0, 4, 4, 0] } }, { name: 'Optimal', type: 'bar', barMaxWidth: 12, data: top.map(w => (w.olac || 0) + (w.osec || 0)), itemStyle: { color: '#13A10E', borderRadius: [0, 4, 4, 0] } }] });
    DT('laTable', { rows: laSorted, pageSize: 25, exportName: M.customer + ' - Log Analytics tier analysis', onRow: w => { LA_SEL = w.i; route(true); window.scrollTo({ top: 0, behavior: 'smooth' }); }, cols: [
      { k: 'n', n: 'Workspace', f: w => '<div class="rn"><b>' + esc(w.n) + '</b><small>' + esc(w.g) + '</small></div>', v: w => w.n }, { k: 'tier', n: 'LA tier' }, { k: 'sen', n: 'Sentinel', v: w => w.sen ? (w.senp || 'Yes') + (w.senc ? ' (' + w.senc + ' GB)' : '') : 'No' },
      { k: 'avg', n: 'Avg GB/day', num: true, f: w => fmtA(w.avg) }, { k: 'p95', n: 'P95', num: true, f: w => fmtA(w.p95) }, { k: 'max', n: 'Max', num: true, f: w => fmtA(w.max) }, { k: 'cu', n: 'Commit util.', num: true, f: w => pctS(w.cu) },
      { k: 'cur', n: 'Current est./mo', num: true, v: w => (w.cla || 0) + (w.cse || 0), f: w => money((w.cla || 0) + (w.cse || 0)) }, { k: 'opt', n: 'Optimal tier', v: w => [w.ola, w.sen ? w.ose : ''].filter(Boolean).join(' + ') },
      { k: 'sav', n: 'Savings/mo', num: true, f: w => w.sav ? '<b class="pos">' + money(w.sav) + '</b>' : '–' }, { k: 'st', n: 'Status', f: w => stPill(w.st), v: w => w.st }, { k: 'sub', n: 'Subscription', v: w => D.subs[w.s] }, { k: 'ret', n: 'Retention', num: true }] });
  }
  DT('vmcTable', { rows: vmc, pageSize: 10, emptyText: 'No steady-state VMs qualify as commitment candidates in the current filter.', exportName: M.customer + ' - VM commitment candidates', onRow: x => { if (x.R) openRes(x.R); }, sortK: 'sa', sortD: -1,
    cols: [recCols()[3], { k: 'obs', n: 'Observed', f: x => '<div class="clip">' + esc(x.obs) + '</div>', v: x => x.obs }, recCols()[4], { k: 'cmc', n: 'Monthly cost', num: true, f: x => money(x.cmc) }, recCols()[5], recCols()[9]] });
  const rawTable = (id, rows, empty) => { if (!rows.length) { $('#' + id).innerHTML = emptyMsg(empty); return; } const keys = Object.keys(rows[0]); DT(id, { rows, pageSize: 10, exportName: id, cols: keys.map(k => ({ k, n: k, num: rows.some(r => isNum(r[k])), f: r => /pct|percent|util/i.test(k) && isNum(Number(r[k])) && r[k] !== '' ? ubar(Number(r[k])) : esc(r[k]) })) }); };
  rawTable('riTable', ri, 'No reservations found in scope (or no permission to read reservation utilization).');
  rawTable('spTable', sp, 'No savings plans found in scope (or no permission to read benefit utilization).');
}

// ---------------------------------------------------------------- page: cost analysis
function renderCost(el) {
  const rs = fr(), tot = sum(rs, r => r.tc), tr = costTrend(), cs = D.cost || {};
  const full = tr ? tr.full : [], last = full.length ? full[full.length - 1] : null, prev = full.length > 1 ? full[full.length - 2] : null;
  const other = (D.other || []).filter(o => passSR(D.subs[o.s], o.g, [o.n, o.svc, o.t].join(' ')));
  el.innerHTML = '<div class="kpis">' + kpi('Run-rate – analyzed resources', moneyK(tot), 'per month · ' + esc(M.costType), '#0078D4', 'coin') +
    kpi('Run-rate – all billed resources', isNum(M.totalMonthly) ? moneyK(M.totalMonthly) : 'n/a', 'whole scope, unfiltered', '#4F6BED', 'layers') +
    kpi('Last full month', last ? moneyK(last.v) : 'n/a', last ? last.m + (prev && prev.v > 0 ? ' · ' + (last.v >= prev.v ? '+' : '') + fmt(100 * (last.v - prev.v) / prev.v, 1) + '% MoM' : '') : '', '#00B7C3', 'chart') +
    kpi('Change over history', full.length > 1 && full[0].v > 0 ? (last.v >= full[0].v ? '+' : '') + fmt(100 * (last.v - full[0].v) / full[0].v, 1) + '%' : 'n/a', full.length > 1 ? full[0].m + ' → ' + last.m : 'needs 2+ full months', full.length > 1 && last.v < full[0].v ? '#13A10E' : '#D13438', 'down') +
    kpi('Paid, not analyzed', fmt(other.length), moneyK(sum(other, o => o.c)) + '/month (other services)', '#C19C00', 'eye') + '</div>' +
    '<div class="grid g2">' + card('Monthly cost by service', chartBox('coMonth', 340), { info: 'Cost Management history; * = partial month. Respects the subscription filter.' }) + card('Daily cost by service (assessment window)', chartBox('coDaily', 340), { info: 'Respects the subscription filter.' }) + '</div>' +
    '<div class="grid g3">' + card('Spend composition – category › service › status', chartBox('coSun', 380), { info: 'Monthly run-rate of the filtered resources. Inner ring: category; middle: service; outer: utilization status.' }) + card('Cost by subscription', chartBox('coSub', 380)) + card('Cost by location', chartBox('coLoc', 380)) + '</div>' +
    '<div class="grid g2">' + card('Top 20 resources by cost (color = utilization status)', chartBox('coTop', 520)) + card('Cost by resource group (top 20)', chartBox('coRg', 520)) + '</div>' +
    card('Billed resources outside the analyzed types', '<div id="coOther"></div>', { info: 'Every billed resource is visible: these have cost in Cost Management but no dedicated utilization analysis (e.g. marketplace, support, child resources).' });
  trendChart('coMonth', 9);
  if ((cs.cd || []).length) {
    const bySvc = new Map(); cs.cd.forEach(x => { if (F.sub.size && !F.sub.has(D.subs[x[1]])) return; const s = cs.svcs[x[2]]; if (!bySvc.has(s)) bySvc.set(s, new Array(DATES.length).fill(0)); bySvc.get(s)[x[0]] += x[3]; });
    const ranked = Array.from(bySvc.entries()).map(e => ({ s: e[0], v: e[1], t: sum(e[1]) })).sort((a, b) => b.t - a.t), top = ranked.slice(0, 8), rest = ranked.slice(8);
    if (rest.length) top.push({ s: 'Other services', v: DATES.map((d, i) => sum(rest, x => x.v[i])) });
    ec('coDaily', { grid: { left: 8, right: 14, top: 44, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', valueFormatter: v => money(v) }, xAxis: { type: 'category', data: DATES.map(shortDate), boundaryGap: false }, yAxis: { type: 'value', axisLabel: axisMoney },
      series: top.map((x, i) => ({ name: x.s, type: 'line', stack: 'd', areaStyle: { opacity: 0.75 }, showSymbol: false, smooth: true, lineStyle: { width: 0.8 }, data: x.v.map(v => Math.round(v * 100) / 100), itemStyle: { color: PAL[i % PAL.length] }, emphasis: { focus: 'series' }, markArea: i === 0 ? weekendArea() : undefined })) });
  } else $('#coDaily').innerHTML = emptyMsg('No daily cost data.');
  const sun = Array.from(groupBy(rs.filter(r => r.tc > 0), r => r.cat).entries()).map((e, i) => ({ name: e[0], itemStyle: { color: PAL[i % PAL.length] }, children: Array.from(groupBy(e[1], r => r.tn).entries()).map(f => ({ name: f[0], children: Array.from(groupBy(f[1], r => r.sgi).entries()).map(g => ({ name: SG[g[0]].n, value: Math.round(sum(g[1], r => r.tc)), itemStyle: { color: SG[g[0]].c } })) })) }));
  if (sun.length) ec('coSun', { tooltip: { formatter: p => '<b>' + esc(p.treePathInfo.slice(1).map(x => x.name).join(' › ')) + '</b><br>' + money(p.value) + '/month' }, series: [{ type: 'sunburst', data: sun, radius: ['12%', '96%'], sort: (a, b) => b.getValue() - a.getValue(), itemStyle: { borderColor: css('--card'), borderWidth: 1.5 }, label: { show: false }, levels: [{}, { r0: '12%', r: '42%', label: { show: true, rotate: 'tangential', fontSize: 10, color: '#fff', minAngle: 12 } }, { r0: '42%', r: '78%', label: { show: true, fontSize: 9, color: '#fff', minAngle: 14, overflow: 'truncate', width: 70 } }, { r0: '78%', r: '96%', label: { show: false } }] }] });
  else $('#coSun').innerHTML = emptyMsg('No cost data in the current filter.');
  const bySub = Array.from(groupBy(rs, r => r.sub).entries()).map(e => ({ n: e[0], v: sum(e[1], r => r.tc), s: sum(e[1], r => r.sv) })).filter(x => x.v > 0).sort((a, b) => b.v - a.v);
  if (bySub.length) ec('coSub', hbarOpt(bySub.map(x => x.n), bySub.map(x => Math.round(x.v)), PAL, moneyK, { tooltip: { trigger: 'item', formatter: p => { const x = bySub[p.dataIndex]; return '<b>' + esc(x.n) + '</b><br>' + money(x.v) + '/month<br>Savings ' + money(x.s) + '/yr'; } } }), p => setFilter('sub', [bySub[p.dataIndex].n])); else $('#coSub').innerHTML = emptyMsg();
  const byLoc = Array.from(groupBy(rs, r => r.l || '(none)').entries()).map((e, i) => ({ name: e[0], value: Math.round(sum(e[1], r => r.tc)), itemStyle: { color: PAL[i % PAL.length] } })).filter(x => x.value > 0).sort((a, b) => b.value - a.value);
  if (byLoc.length) ec('coLoc', donutOpt(byLoc, moneyK(tot), 'per month', money), p => setFilter('loc', [p.name])); else $('#coLoc').innerHTML = emptyMsg();
  const top = rs.filter(r => r.tc > 0).sort((a, b) => b.tc - a.tc).slice(0, 20);
  if (top.length) ec('coTop', hbarOpt(top.map(r => r.n), top.map(r => Math.round(r.tc)), top.map(r => STC(r.st)), moneyK, { tooltip: { trigger: 'item', formatter: p => { const r = top[p.dataIndex]; return '<b>' + esc(r.n) + '</b><br>' + esc(r.tn) + '<br>' + money(r.tc) + '/month · utilization ' + pctS(r.u, 1) + '<br>' + esc(r.st) + (r.sv ? '<br><b>Savings ' + money(r.sv) + '/yr</b>' : ''); } } }), p => openRes(top[p.dataIndex])); else $('#coTop').innerHTML = emptyMsg();
  const byRg = Array.from(groupBy(rs, r => r.sub + ' / ' + r.g).entries()).map(e => ({ n: e[1][0].g, f: e[0], v: sum(e[1], r => r.tc), s: wscore(e[1]) })).filter(x => x.v > 0).sort((a, b) => b.v - a.v).slice(0, 20);
  if (byRg.length) ec('coRg', hbarOpt(byRg.map(x => x.n), byRg.map(x => Math.round(x.v)), byRg.map(x => scoreColor(x.s)), moneyK, { tooltip: { trigger: 'item', formatter: p => { const x = byRg[p.dataIndex]; return '<b>' + esc(x.f) + '</b><br>' + money(x.v) + '/month · efficiency ' + (isNum(x.s) ? Math.round(x.s) : 'n/a'); } } }), p => setFilter('rg', [byRg[p.dataIndex].n])); else $('#coRg').innerHTML = emptyMsg();
  DT('coOther', { rows: other, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - Other billed resources', cols: [{ k: 'n', n: 'Resource', f: o => '<div class="rn"><b>' + esc(o.n) + '</b><small>' + esc(o.t) + '</small></div>', v: o => o.n }, { k: 'svc', n: 'Service' }, { k: 'c', n: 'Monthly cost', num: true, f: o => money(o.c) }, { k: 'sub', n: 'Subscription', v: o => D.subs[o.s] }, { k: 'g', n: 'Resource group' }, { k: 'id', n: 'Resource ID', f: o => '<div class="clip muted" style="max-width:300px">' + esc(o.id) + '</div>' }] });
}

// ---------------------------------------------------------------- page: tags & environments
let TAG_K = '';
function renderTags(el) {
  const rs = fr(), tot = sum(rs, r => r.tc);
  const keys = TAGK.map(t => { const l = rs.filter(r => r.tgl[t.k] != null && r.tgl[t.k] !== ''); return { k: t.k, d: t.d, n: l.length, c: sum(l, r => r.tc) }; }).filter(x => x.n > 0).sort((a, b) => b.n - a.n);
  if (!keys.some(k => k.k === TAG_K)) TAG_K = (keys.find(k => /^env/.test(k.k)) || keys[0] || {}).k || '';
  const tagged = rs.filter(r => Object.keys(r.tg).length), untagged = rs.filter(r => !Object.keys(r.tg).length);
  const envs = ['Production', 'Non-Production', 'Unclassified'].map(e => { const l = rs.filter(r => r.e === e); return { e, n: l.length, c: sum(l, r => r.tc), sv: sum(l, r => r.sv), s: wscore(l), idle: l.filter(r => r.sgi === 0).length }; });
  const envCol = { 'Production': '#0078D4', 'Non-Production': '#8764B8', 'Unclassified': '#8A8886' };
  el.innerHTML = '<div class="kpis">' + kpi('Tag keys in use', fmt(keys.length), 'across the filtered resources', '#0078D4', 'tag') +
    kpi('Resources tagged', pctS(rs.length ? 100 * tagged.length / rs.length : null), fmt(tagged.length) + ' of ' + fmt(rs.length), '#13A10E', 'check') +
    kpi('Spend tagged', pctS(tot > 0 ? 100 * sum(tagged, r => r.tc) / tot : null), moneyK(sum(tagged, r => r.tc)) + '/month', '#4F6BED', 'coin') +
    kpi('Untagged resources', fmt(untagged.length), moneyK(sum(untagged, r => r.tc)) + '/month without any tag', '#D13438', 'eye') + '</div>' +
    '<div class="grid g3">' + envs.map(x => '<div class="envc" style="border-top:4px solid ' + envCol[x.e] + '"><h4>' + esc(x.e) + '</h4><div class="v">' + moneyK(x.c) + '<span class="sm muted"> /month</span></div><div class="tc-m"><span>Resources <b>' + fmt(x.n) + '</b></span><span>Idle <b>' + fmt(x.idle) + '</b></span><span>Efficiency <b>' + (isNum(x.s) ? Math.round(x.s) : 'n/a') + '</b></span><span>Savings <b class="pos">' + moneyK(x.sv) + '</b></span></div></div>').join('') + '</div>' +
    '<div class="grid g2">' + card('Tag coverage – top 20 keys', chartBox('tgCov', 520), { info: 'Share of resources and of monthly cost carrying each tag key. Environment classification uses the Environment/Env tag first, then naming patterns.' }) +
    card('Cost, savings and efficiency by tag value', '<div class="ctl">Tag key <select class="sel" id="tgKey">' + keys.map(k => '<option value="' + esc(k.k) + '"' + (k.k === TAG_K ? ' selected' : '') + '>' + esc(k.d) + ' (' + fmt(k.n) + ')</option>').join('') + '</select><span class="muted sm">Click a bar to filter</span></div>' + chartBox('tgVal', 484)) + '</div>' +
    card('Untagged resources', '<div id="tgTable"></div>');
  const top = keys.slice(0, 20);
  if (top.length) ec('tgCov', { grid: { left: 8, right: 20, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => fmt(v, 1) + '%' }, xAxis: { type: 'value', max: 100, axisLabel: { formatter: v => v + '%' } }, yAxis: { type: 'category', data: top.map(k => k.d), inverse: true, axisLabel: { width: 160, overflow: 'truncate' } },
    series: [{ name: '% of resources', type: 'bar', barMaxWidth: 10, data: top.map(k => Math.round(1000 * k.n / Math.max(1, rs.length)) / 10), itemStyle: { color: '#0078D4', borderRadius: [0, 4, 4, 0] } }, { name: '% of cost', type: 'bar', barMaxWidth: 10, data: top.map(k => tot > 0 ? Math.round(1000 * k.c / tot) / 10 : 0), itemStyle: { color: '#13A10E', borderRadius: [0, 4, 4, 0] } }] }, p => { TAG_K = top[p.dataIndex].k; route(true); });
  else $('#tgCov').innerHTML = emptyMsg('No tags on the filtered resources.');
  const drawVal = () => {
    if (!TAG_K) { $('#tgVal').innerHTML = emptyMsg('No tags on the filtered resources.'); return; }
    const vals = Array.from(groupBy(rs, r => { const v = r.tgl[TAG_K]; return v == null || v === '' ? '(not tagged)' : v; }).entries()).map(e => ({ v: e[0], n: e[1].length, c: sum(e[1], r => r.tc), sv: sum(e[1], r => r.sv), s: wscore(e[1]) })).sort((a, b) => b.c - a.c || b.n - a.n).slice(0, 18);
    ec('tgVal', { grid: { left: 8, right: 20, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, formatter: ps => { const x = vals[ps[0].dataIndex]; return '<b>' + esc(x.v) + '</b><br>' + fmt(x.n) + ' resources · ' + money(x.c) + '/month<br>Savings ' + money(x.sv) + '/yr · efficiency ' + (isNum(x.s) ? Math.round(x.s) : 'n/a'); } },
      xAxis: { type: 'value', axisLabel: axisMoney }, yAxis: { type: 'category', data: vals.map(x => x.v), inverse: true, axisLabel: { width: 160, overflow: 'truncate' } },
      series: [{ name: 'Monthly cost', type: 'bar', stack: 'x', barMaxWidth: 18, data: vals.map(x => ({ value: Math.round(x.c), itemStyle: { color: scoreColor(x.s) } })) }, { name: 'Annual savings', type: 'bar', barMaxWidth: 18, data: vals.map(x => Math.round(x.sv)), itemStyle: { color: '#13A10E', opacity: 0.45 } }] }, p => setTag(TAG_K, vals[p.dataIndex].v));
  };
  const tk = $('#tgKey'); if (tk) tk.onchange = () => { TAG_K = tk.value; CH = CH.filter(c => { if (c.getDom().id === 'tgVal') { c.dispose(); return false; } return true; }); drawVal(); };
  drawVal();
  DT('tgTable', { rows: untagged, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - Untagged resources', onRow: r => openRes(r), cols: resCols(['name', 'sub', 'rg', 'loc', 'env', 'st', 'c', 'sv', 'age']) });
}

// ---------------------------------------------------------------- page: idle & orphaned
function renderWaste(el) {
  const rs = fr(), idle = rs.filter(r => r.sgi === 0), stop = rs.filter(r => r.sgi === 4);
  const orph = (D.orph || []).filter(o => passSR(D.subs[o.s], o.g, [o.n, o.f, o.t, o.d].join(' '))), paid = orph.filter(o => /paid/i.test(o.ci));
  el.innerHTML = '<div class="kpis">' + kpi('Idle · unused · orphaned', fmt(idle.length), moneyK(sum(idle, r => r.tc)) + '/month · ' + moneyK(sum(idle, r => r.sv)) + '/yr savings', '#D13438', 'trash') +
    kpi('Orphaned findings', fmt(orph.length), fmt(paid.length) + ' still billed · ' + moneyK(sum(paid, o => o.c)) + '/month', '#F7630C', 'eye') +
    kpi('Stopped · deallocated', fmt(stop.length), moneyK(sum(stop, r => r.tc)) + '/month still billed', '#C19C00', 'pause') +
    kpi('Aged snapshots & backups', fmt(rs.filter(r => r.st === 'Aged').length), 'older than ' + TH.snap + ' days', '#8764B8', 'clock') + '</div>' +
    '<div class="grid g2">' + card('Monthly cost of idle, unused, stopped and orphaned resources', chartBox('waTree', 340), { info: 'Area = monthly cost by service and status. Click to open the service dashboard.' }) + card('Orphaned resources by finding', chartBox('waOrph', 340)) + '</div>' +
    card('Idle, unused and orphaned resources (' + idle.length + ')', '<div id="waIdle"></div>') + card('Orphaned resource findings (' + orph.length + ')', '<div id="waOrphT"></div>', { info: 'Unattached disks and NICs, unassociated public IPs, empty plans and pools, gateways without connections and similar findings from Resource Graph.' }) + card('Stopped and deallocated resources (' + stop.length + ')', '<div id="waStop"></div>');
  const wt = Array.from(groupBy(idle.concat(stop).filter(r => r.tc > 0), r => r.t).entries()).map(e => ({ name: T[e[0]].n, k: T[e[0]].k, value: Math.round(sum(e[1], r => r.tc)), children: Array.from(groupBy(e[1], r => r.st).entries()).map(f => ({ name: f[0], value: Math.round(sum(f[1], r => r.tc)), itemStyle: { color: STC(f[0]) } })) }));
  if (wt.length) ec('waTree', { tooltip: { formatter: p => '<b>' + esc(p.name) + '</b><br>' + money(p.value) + '/month' }, series: [{ type: 'treemap', data: wt, roam: false, nodeClick: false, breadcrumb: { show: false }, top: 0, left: 0, right: 0, bottom: 0, upperLabel: { show: true, height: 22, color: '#fff' }, itemStyle: { color: '#33415C' }, label: { formatter: p => p.name + '\n' + moneyK(p.value), color: '#fff', fontSize: 11 }, levels: [{ itemStyle: { borderColor: css('--card'), borderWidth: 3, gapWidth: 3 } }, { itemStyle: { borderColor: 'rgba(255,255,255,.3)', borderWidth: 1, gapWidth: 1 } }] }] }, p => { if (p.data && p.data.k) go('services/' + encodeURIComponent(p.data.k)); });
  else $('#waTree').innerHTML = emptyMsg('No idle or stopped resources with cost in the current filter.');
  const of = Array.from(groupBy(orph, o => o.f).entries()).map(e => ({ n: e[0], c: e[1].length, p: e[1].filter(o => /paid/i.test(o.ci)).length })).sort((a, b) => b.c - a.c);
  if (of.length) ec('waOrph', { grid: { left: 8, right: 20, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, xAxis: { type: 'value' }, yAxis: { type: 'category', data: of.map(x => x.n), inverse: true, axisLabel: { width: 190, overflow: 'truncate' } },
    series: [{ name: 'Billed', type: 'bar', stack: 'o', barMaxWidth: 18, data: of.map(x => x.p), itemStyle: { color: '#D13438' } }, { name: 'No direct cost', type: 'bar', stack: 'o', barMaxWidth: 18, data: of.map(x => x.c - x.p), itemStyle: { color: '#A0AEC0', borderRadius: [0, 4, 4, 0] } }] });
  else $('#waOrph').innerHTML = emptyMsg('No orphaned resources found.');
  DT('waIdle', { rows: idle, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - Idle resources', onRow: r => openRes(r), cols: resCols(['name', 'sub', 'rg', 'sku', 'st', 'u', 'c', 'sv', 'rec', 'us', 'age']) });
  DT('waOrphT', { rows: orph, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - Orphaned resources', cols: [{ k: 'f', n: 'Finding' }, { k: 'n', n: 'Resource', f: o => '<div class="rn"><b>' + esc(o.n) + '</b><small>' + esc(o.t) + '</small></div>', v: o => o.n }, { k: 'ci', n: 'Cost impact', f: o => pill(o.ci, /paid/i.test(o.ci) ? '#D13438' : '#8A8886'), v: o => o.ci }, { k: 'c', n: 'Monthly cost', num: true, f: o => money(o.c) }, { k: 'sku', n: 'SKU' }, { k: 'd', n: 'Detail', f: o => '<div class="clip" title="' + esc(o.d) + '">' + esc(o.d) + '</div>' }, { k: 'sub', n: 'Subscription', v: o => D.subs[o.s] }, { k: 'g', n: 'Resource group' }, { k: 'l', n: 'Location' }, { k: 'id', n: 'Resource ID', f: o => '<div class="clip muted" style="max-width:240px">' + esc(o.id) + '</div>' }] });
  DT('waStop', { rows: stop, pageSize: 25, sortK: 'c', sortD: -1, exportName: M.customer + ' - Stopped and deallocated', onRow: r => openRes(r), cols: resCols(['name', 'sub', 'rg', 'sku', 'st', 'c', 'sv', 'rec', 'us', 'age']) });
}

// ---------------------------------------------------------------- page: resource explorer
const EX = { type: '', cols: null };
const EX_DEFAULT = ['c:name', 'c:sub', 'c:rg', 'c:loc', 'c:env', 'c:sku', 'c:st', 'c:u', 'c:es', 'c:c', 'c:sv', 'c:rec'];
const PATN = ['Active hours %', 'Idle hours %', 'Business-hours avg', 'Off-hours avg', 'Weekday avg', 'Weekend avg', 'Peak hour', 'Peak day'];
function exDefs(t, rows) {
  const defs = [];
  ['name', 'type', 'sub', 'rg', 'loc', 'env', 'sku', 'st', 'u', 'es', 'c', 'sv', 'rec', 'pri', 'pv', 'us', 'age', 'cr', 'trend', 'id'].forEach(k => defs.push({ g: 'Core', id: 'c:' + k, d: RCOL[k] }));
  defs.push({ g: 'Core', id: 'c:k', d: { k: 'c:k', n: 'Kind', v: r => r.k } }, { g: 'Core', id: 'c:ms', d: { k: 'c:ms', n: 'Metrics status', v: r => r.ms } }, { g: 'Core', id: 'c:ed', d: { k: 'c:ed', n: 'Effective discount %', num: true, v: r => r.ed, f: r => pctS(r.ed, 1) } }, { g: 'Core', id: 'c:an', d: { k: 'c:an', n: 'Analysis note', v: r => r.an } }, { g: 'Core', id: 'c:ri', d: { k: 'c:ri', n: 'Rule IDs', v: r => r.recs.map(x => x.id).join(', ') } });
  PATN.forEach((n, j) => defs.push({ g: 'Usage pattern', id: 'pt:' + j, d: { k: 'pt:' + j, n: n, num: j < 7, v: r => r.pat ? r.pat[j] : null, f: r => r.pat ? (j < 6 ? fmtA(r.pat[j]) : esc(r.pat[j])) : '–' } }));
  if (t) {
    (t.fc || []).forEach(c => { const num = rows.some(r => isNum(fv(r, c))); defs.push({ g: 'Configuration', id: 'f:' + c, d: { k: 'f:' + c, n: c, num: num, v: r => fv(r, c), f: r => { const v = fv(r, c); return num && isNum(v) ? fmtA(v) : '<div class="clip" style="max-width:260px" title="' + esc(v) + '">' + esc(v) + '</div>'; } } }); });
    (t.mc || []).forEach(c => defs.push({ g: 'Metrics', id: 'm:' + c, d: { k: 'm:' + c, n: mLabel(t, c), num: true, v: r => fv(r, c), f: r => fmtA(fv(r, c)) } }));
    (t.pc || []).forEach(c => defs.push({ g: 'Properties', id: 'p:' + c, d: { k: 'p:' + c, n: c.replace(/^Prop\./, ''), v: r => fv(r, c), f: r => '<div class="clip" style="max-width:260px" title="' + esc(fv(r, c)) + '">' + esc(fv(r, c)) + '</div>' } }));
  }
  TAGK.filter(tk => rows.some(r => r.tgl[tk.k] != null)).slice(0, 200).forEach(tk => defs.push({ g: 'Tags', id: 't:' + tk.k, d: { k: 't:' + tk.k, n: 'Tag: ' + tk.d, v: r => r.tgl[tk.k] } }));
  return defs;
}
function renderExplorer(el, arg) {
  if (arg && tByKey[arg] && EX.type !== arg) { EX.type = arg; EX.cols = null; }
  if (!arg && !EX.type && F.type.size === 1) { const t0 = T.find(t => F.type.has(t.n)); if (t0) EX.type = t0.k; }
  const all = fr(), types = Array.from(groupBy(all, r => r.t).entries()).map(e => ({ t: T[e[0]], n: e[1].length })).sort((a, b) => a.t.n.localeCompare(b.t.n));
  if (EX.type && !types.some(x => x.t.k === EX.type)) EX.type = '';
  const t = EX.type ? tByKey[EX.type] : null, rows = t ? all.filter(r => r.T === t) : all, defs = exDefs(t, rows), byId = {}; defs.forEach(d => { byId[d.id] = d; });
  if (!EX.cols) { EX.cols = EX_DEFAULT.slice(); if (t) { (t.fc || []).slice(0, 5).forEach(c => EX.cols.push('f:' + c)); (t.mc || []).filter(c => /_(P95|Avg|Total)$/.test(c)).slice(0, 6).forEach(c => EX.cols.push('m:' + c)); } }
  EX.cols = EX.cols.filter(id => byId[id]);
  const groups = Array.from(groupBy(defs, d => d.g).entries());
  el.innerHTML = '<div class="card"><div class="card-b" style="padding-top:14px"><div class="ctl" style="margin-bottom:10px">Resource type <select class="sel" id="exType"><option value="">All types (' + fmt(all.length) + ')</option>' + types.map(x => '<option value="' + esc(x.t.k) + '"' + (x.t.k === EX.type ? ' selected' : '') + '>' + esc(x.t.n) + ' (' + fmt(x.n) + ')</option>').join('') + '</select>' +
    '<button class="btn sm" id="exColBtn">' + ICON.table + ' Columns (' + EX.cols.length + ' of ' + defs.length + ')</button><span class="muted sm">' + (t ? 'All configuration, metric, property and tag columns of ' + esc(t.n) + ' are available.' : 'Select a resource type to add its configuration, metric and property columns.') + '</span></div>' +
    '<div id="exCols" class="card" style="display:none;margin-bottom:12px;box-shadow:none"><div class="card-b"><div class="ctl"><div class="dt-s">' + ICON.search + '<input id="exColQ" placeholder="Find a column…"></div><button class="btn sm" data-ex="def">Default</button><button class="btn sm" data-ex="metrics">+ all metrics</button><button class="btn sm" data-ex="cfg">+ all configuration</button><button class="btn sm" data-ex="tags">+ all tags</button><button class="btn sm" data-ex="props">+ all properties</button><button class="btn sm" data-ex="min">Minimal</button></div>' +
    '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:4px 14px;max-height:340px;overflow:auto">' + groups.map(g => '<div style="grid-column:1/-1;margin-top:6px" class="pal-g">' + esc(g[0]) + ' (' + g[1].length + ')</div>' + g[1].map(d => '<label class="dd-o" data-n="' + esc((d.d.n + ' ' + d.id).toLowerCase()) + '"><input type="checkbox" data-id="' + esc(d.id) + '"' + (EX.cols.indexOf(d.id) >= 0 ? ' checked' : '') + '><span class="dd-t" title="' + esc(d.id) + '">' + esc(d.d.n) + '</span></label>').join('')).join('') + '</div></div></div>' +
    '<div id="exTable"></div></div></div>';
  $('#exType').onchange = e => { EX.type = e.target.value; EX.cols = null; if (location.hash.indexOf('#/explorer/') === 0) location.hash = '#/explorer'; else route(true); };
  $('#exColBtn').onclick = () => { const p = $('#exCols'); p.style.display = p.style.display === 'none' ? '' : 'none'; };
  $('#exColQ').addEventListener('input', e => { const q = e.target.value.toLowerCase(); $$('#exCols .dd-o').forEach(o => { o.style.display = o.dataset.n.indexOf(q) >= 0 ? '' : 'none'; }); });
  $('#exCols').addEventListener('change', e => { const id = e.target.dataset.id; if (!id) return; if (e.target.checked) { if (EX.cols.indexOf(id) < 0) EX.cols.push(id); } else EX.cols = EX.cols.filter(x => x !== id); draw(); $('#exColBtn').innerHTML = ICON.table + ' Columns (' + EX.cols.length + ' of ' + defs.length + ')'; });
  $$('[data-ex]').forEach(b => { b.onclick = () => {
    const a = b.dataset.ex, add = g => defs.filter(d => d.g === g).forEach(d => { if (EX.cols.indexOf(d.id) < 0) EX.cols.push(d.id); });
    if (a === 'def') EX.cols = null; else if (a === 'min') EX.cols = ['c:name', 'c:st', 'c:u', 'c:c', 'c:sv']; else if (a === 'metrics') add('Metrics'); else if (a === 'cfg') add('Configuration'); else if (a === 'tags') add('Tags'); else if (a === 'props') add('Properties');
    route(true); $('#exCols').style.display = '';
  }; });
  function draw() { DT('exTable', { rows, pageSize: 50, sortK: 'c', sortD: -1, exportName: M.customer + ' - Resource explorer' + (t ? ' - ' + t.n : ''), onRow: r => openRes(r), cols: EX.cols.map(id => byId[id].d) }); }
  draw();
}

// ---------------------------------------------------------------- page: methodology & rules
function renderMethod(el) {
  const recs = frecs();
  const th = [['Idle', '< ' + TH.idle + '% average and < ' + (2 * TH.idle) + '% P95'], ['Over-provisioned', 'P95 < ' + TH.low + '% (memory peak < ' + TH.memLow + '%)'], ['Under-provisioned', 'P95 > ' + TH.high + '% (memory P95 > ' + TH.memHigh + '%)'], ['Right-size target', 'P95 ≤ ' + TH.target + '% after the change'],
    ['Auto-discovered types', 'low < ' + TH.genLow + '%, hot > ' + TH.genHigh + '%'], ['Commitment candidates', 'running ≥ ' + TH.steady + '% of hours'], ['Aged snapshots', '> ' + TH.snap + ' days'], ['Business hours', M.bizStart + ':00–' + M.bizEnd + ':00, weekend ' + (M.weekend || []).join('/') + ' (' + M.tz + ')']];
  const src = M.sources || {};
  el.innerHTML = '<div class="grid g4">' + th.map(x => '<div class="envc"><h4 class="muted sm" style="text-transform:uppercase;letter-spacing:.05em">' + esc(x[0]) + '</h4><div style="font-weight:650">' + esc(x[1]) + '</div></div>').join('') + '</div>' +
    '<div class="grid g2">' + card('Methodology', '<div id="mtMethod"></div>') + card('Data sources & collection', '<div class="facts" style="grid-template-columns:1fr 1fr">' + Object.keys(src).map(k => '<div><span>' + esc(k) + '</span><b title="' + esc(src[k]) + '">' + esc(src[k]) + '</b></div>').join('') +
      '<div><span>Collected</span><b>' + esc(M.collectedAt) + '</b></div><div><span>Collected by</span><b>' + esc(M.collectedBy) + '</b></div><div><span>Metric requests</span><b>' + fmt(M.metricRequests) + (M.batch ? ' (batch API)' : '') + '</b></div><div><span>Inventory</span><b>' + fmt(M.inventory) + ' resources</b></div><div><span>Report generated</span><b>' + esc(M.generated) + '</b></div><div><span>Dashboard limits</span><b>' + fmt((M.limits || {}).props) + ' properties/type</b></div></div>' +
      '<h4 class="muted sm" style="margin:14px 0 6px;text-transform:uppercase">Scope</h4><div id="mtScope"></div>') + '</div>' +
    card('Recommendation rules', '<div id="mtRules"></div>', { info: 'Every suggested recommendation references one of these rules. Rule IDs are <service prefix>-<rule family>, e.g. VM-RS-SKU.' }) +
    '<div class="grid g2">' + card('Collection coverage by resource type', chartBox('mtCov', Math.max(280, 22 * Math.min(40, (D.cov || []).length) + 60))) + card('Inventory by ARM type', '<div id="mtInv"></div>', { info: 'All resources found in scope; types without a utilization analysis are free (e.g. NSGs, route tables) or covered through cost only.' }) + '</div>' +
    card('Metric dictionary', '<div id="mtDict"></div>', { info: 'Azure Monitor platform metrics behind every metric column (dedicated catalog and auto-discovered metrics).' });
  DT('mtMethod', { rows: D.method || [], pageSize: 25, search: false, cols: [{ k: 'Topic', n: 'Topic', f: x => '<b>' + esc(x.Topic) + '</b>' }, { k: 'Detail', n: 'Detail', f: x => '<div class="wrap" style="max-width:none">' + esc(x.Detail) + '</div>' }] });
  DT('mtScope', { rows: D.scope || [], pageSize: 10, search: false, cols: [{ k: 'SubscriptionName', n: 'Subscription' }, { k: 'SubscriptionId', n: 'ID' }, { k: 'State', n: 'State' }] });
  const rules = Object.keys(D.rules || {}).map(k => { const r = D.rules[k], l = recs.filter(x => x.rc === k); return { k, n: r.n, c: r.c, r: r.r, ap: (r.ap || []).join(', '), cnt: l.length, sa: dedupSav(l).total * 12, all: RC.filter(x => x.rc === k).length }; });
  DT('mtRules', { rows: rules, pageSize: 50, sortK: 'sa', sortD: -1, exportName: 'Recommendation rules', onRow: x => go('recs/' + encodeURIComponent(x.k)), cols: [{ k: 'k', n: 'Rule family', f: x => '<code>' + esc(x.k) + '</code>' }, { k: 'n', n: 'Name', f: x => '<b>' + esc(x.n) + '</b>' }, { k: 'c', n: 'Criteria', f: x => '<div class="wrap">' + esc(x.c) + '</div>' }, { k: 'r', n: 'Rationale', f: x => '<div class="wrap" style="min-width:200px">' + esc(x.r) + '</div>' }, { k: 'ap', n: 'Applies to', f: x => '<div class="wrap" style="min-width:200px">' + esc(x.ap) + '</div>' }, { k: 'cnt', n: 'Triggered (filtered)', num: true }, { k: 'all', n: 'Triggered (all)', num: true }, { k: 'sa', n: 'Savings / yr', num: true, f: x => x.sa ? '<b class="pos">' + money(x.sa) + '</b>' : '–' }] });
  const cov = (D.cov || []).slice().sort((a, b) => b.r - a.r).slice(0, 40);
  if (cov.length) ec('mtCov', { grid: { left: 8, right: 20, top: 30, bottom: 6, containLabel: true }, legend: { top: 0 }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, xAxis: { type: 'value' }, yAxis: { type: 'category', data: cov.map(c => c.n + (c.g ? ' *' : '')), inverse: true, axisLabel: { width: 180, overflow: 'truncate' } },
    series: [['Metrics OK', 'ok', '#13A10E'], ['No data', 'nd', '#A0AEC0'], ['Error', 'er', '#D13438'], ['Not collected', 'nc', '#E1E6EE']].map(s => ({ name: s[0], type: 'bar', stack: 'c', barMaxWidth: 14, itemStyle: { color: s[2] }, data: cov.map(c => s[1] === 'nc' ? Math.max(0, c.r - c.ok - c.nd - c.er) : c[s[1]]) })) });
  DT('mtInv', { rows: D.inv || [], pageSize: 15, sortK: 'n', sortD: -1, exportName: 'Inventory by type', cols: [{ k: 't', n: 'ARM type' }, { k: 'n', n: 'Resources', num: true }, { k: 'k', n: 'Analyzed as', f: x => x.k ? '<span class="pill" style="--pc:#13A10E">' + esc((tByKey[x.k] || {}).n || x.k) + '</span>' : '<span class="muted">inventory / cost only</span>', v: x => x.k }] });
  DT('mtDict', { rows: D.mcat || [], pageSize: 25, exportName: 'Metric dictionary', cols: [{ k: 't', n: 'Resource type' }, { k: 'a', n: 'Column prefix', f: x => '<code>' + esc(x.a) + '</code>' }, { k: 'm', n: 'Azure Monitor metric' }, { k: 'o', n: 'Statistics' }, { k: 'sub', n: 'Sub-resource' }, { k: 'f', n: 'Dimension filter' }] });
}

// ---------------------------------------------------------------- initialization
function setThemeIcon() { $('#btnTheme').innerHTML = document.documentElement.dataset.theme === 'dark' ? ICON.sun : ICON.moon; }
function init() {
  let saved = null; try { saved = localStorage.getItem('uco-theme'); } catch (e) { saved = null; }
  if (saved === 'dark' || (!saved && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches)) document.documentElement.dataset.theme = 'dark';
  $('#bCustomer').textContent = M.customer;
  $('#nav').innerHTML = PAGES.map(p => '<a class="nav-i" data-k="' + p.k + '" href="#/' + p.k + '" title="' + esc(p.n) + '">' + ICON[p.ic] + '<span>' + esc(p.n) + '</span>' + (p.badge ? '<span class="bdg" data-b="' + p.k + '"></span>' : '') + '</a>').join('');
  $('#sideFoot').innerHTML = esc(M.start) + ' → ' + esc(M.end) + '<br>' + fmt(M.analyzed) + ' resources · ' + fmt(T.length) + ' services<br>Generated ' + esc(M.generated) + '<br>Assessment v' + esc(M.version);
  $('#btnMenu').innerHTML = ICON.menu; $('#btnMenu').onclick = () => { document.body.classList.toggle('collapsed'); setTimeout(() => CH.forEach(c => c.resize()), 260); };
  setThemeIcon();
  $('#btnTheme').onclick = () => { const d = document.documentElement; d.dataset.theme = d.dataset.theme === 'dark' ? 'light' : 'dark'; try { localStorage.setItem('uco-theme', d.dataset.theme); } catch (e) { /* ignore */ } setThemeIcon(); route(true); };
  $('#btnPrint').innerHTML = ICON.print; $('#btnPrint').onclick = () => window.print();
  $('#icoSearch').innerHTML = ICON.search; $('#btnPalette').onclick = openPalette;
  $('#flbl').innerHTML = ICON.filter + ' Filters'; $('#icoFs').innerHTML = ICON.search;
  $('#fq').addEventListener('input', debounce(e => { F.q = e.target.value.trim(); applyFilters(); }, 260));
  $('#freset').onclick = () => { DIMS.forEach(d => F[d.k].clear()); F.tagKey = ''; F.tagVals.clear(); F.q = ''; $('#fq').value = ''; applyFilters(); };
  $('#fbtns').addEventListener('click', e => { const b = e.target.closest('.fbtn'); if (!b) return; const dd = $('#dd'); if (dd.classList.contains('open') && dd.dataset.dim === b.dataset.dim) { closeDD(); return; } openDD(b.dataset.dim); });
  $('#fchips').addEventListener('click', e => { const b = e.target.closest('button[data-cd]'); if (!b) return; const k = b.dataset.cd; if (k === 'q') { F.q = ''; $('#fq').value = ''; } else if (k === 'tag') F.tagVals.delete(b.dataset.cv); else F[k].delete(b.dataset.cv); applyFilters(); });
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); openPalette(); }
    else if (e.key === 'Escape') { closePalette(); closeDrawer(); closeDD(); }
    else if (e.key === '/' && !/^(input|select|textarea)$/i.test((document.activeElement || {}).tagName || '')) { e.preventDefault(); $('#fq').focus(); }
  });
  const savedPath = M.savedPath || (location.protocol === 'file:' ? decodeURIComponent(location.pathname).replace(/^\/([A-Za-z]:)/, '$1').replace(/\//g, '\\') : location.href);
  $('#footPath').textContent = savedPath; $('#footCopy').onclick = () => copyText(savedPath);
  $('#footMeta').innerHTML = 'Azure Utilization &amp; Cost Optimizer v' + esc(M.version) + ' · generated ' + esc(M.generated) + ' · data collected ' + esc(M.collectedAt) + ' · ' + fmt(M.metricRequests) + ' Azure Monitor requests · ' + esc(M.costType) + ' cost<br>Estimates are indicative – validate owners, licensing, reservations and compliance before changing resources.';
  renderFilterBar();
  route(false);
}
init();
})();
</script>
</body>
</html>
'@

#endregion

#region ============================== HTML DASHBOARD: RENDER ==============================

# An optional Templates/Dashboard.html next to this script overrides the embedded template (dashboard customization)
$TemplatePath = Join-Path $PSScriptRoot 'Templates' 'Dashboard.html'
$tplSource = if (Test-Path -LiteralPath $TemplatePath) { [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $TemplatePath).ProviderPath) } else { $script:EmbeddedDashboardTemplate }
if ([string]::IsNullOrWhiteSpace($tplSource)) { Write-Log 'Dashboard template not available - HTML dashboard skipped' Error; $HtmlPath = $null }
else {
    $libPath = Join-Path $PSScriptRoot 'Templates' 'lib' 'echarts.min.js'
    if (-not (Test-Path $libPath)) {
        try {
            New-Item -ItemType Directory -Path (Split-Path $libPath) -Force | Out-Null
            Invoke-WebRequest -Uri $EChartsUrl -OutFile $libPath -UseBasicParsing -TimeoutSec 90
            Write-Log "Charting library cached for offline use: $libPath" Info
        }
        catch { Write-Log "Charting library could not be downloaded ($($_.Exception.Message)); the dashboard will load it from $EChartsUrl when opened" Warning; Remove-Item $libPath -Force -ErrorAction SilentlyContinue }
    }
    $echartsJs = if (Test-Path $libPath) { [System.IO.File]::ReadAllText($libPath) } else { $null }
    if ($echartsJs -and ($echartsJs.Length -lt 200000 -or -not $echartsJs.Contains('echarts'))) { Write-Log 'Cached charting library looks invalid - using the CDN reference' Warning; $echartsJs = $null }
    $ecTag = if ($echartsJs) { "<script>`n" + $echartsJs.Replace('</script', '<\/script') + "`n</script>" } else { "<script src=`"$EChartsUrl`"></script>" }

    $tpl = $tplSource
    $tpl = $tpl.Replace('__TITLE__', [System.Net.WebUtility]::HtmlEncode("$CustomerName - Azure Utilization & Cost Optimizer"))
    $m1 = '<!--__ECHARTS__-->'; $m2 = '/*__DATA__*/'
    $i1 = $tpl.IndexOf($m1); $i2 = $tpl.IndexOf($m2)
    if ($i1 -lt 0 -or $i2 -lt 0 -or $i2 -lt $i1) { Write-Log 'Dashboard template markers not found - HTML dashboard skipped' Error; $HtmlPath = $null }
    else {
        $sb = [System.Text.StringBuilder]::new($tpl.Length + $ecTag.Length + $json.Length + 64)
        [void]$sb.Append($tpl, 0, $i1).Append($ecTag).Append($tpl, $i1 + $m1.Length, $i2 - $i1 - $m1.Length).Append($json).Append($tpl, $i2 + $m2.Length, $tpl.Length - $i2 - $m2.Length)
        [System.IO.File]::WriteAllText($HtmlPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
        Write-Log "HTML dashboard saved: $HtmlPath ($([Math]::Round((Get-Item -LiteralPath $HtmlPath).Length / 1MB, 1)) MB; $($jsRes.Count) resources, $($jsRecs.Count) recommendations)" Success
    }
}

#endregion

#region ============================== FINISH ==============================

$dur = (Get-Date) - $script:RunStart
Write-Log ('=' * 78) Section
Write-Log "PHASE 2 COMPLETE in $([Math]::Round($dur.TotalMinutes, 1)) min" Success
Write-Log "  Resources analyzed      : $($AllRows.Count) across $($TypeSummary.Count) resource types"
Write-Log "  Recommendations         : $($script:Recs.Count) (Delete $($CatSavings[0].Recommendations), Rightsize $($CatSavings[1].Recommendations), Configure $($CatSavings[2].Recommendations), Performance $PerfCount)"
Write-Log "  Estimated annual savings: $(Money $AnnualSavings) (de-duplicated per resource)"
Write-Log "  Efficiency score        : $(if ($null -ne $EfficiencyOverall) { "$EfficiencyOverall / 100 (cost-weighted)" } else { 'n/a' })"
Write-Log "  Master workbook         : $XlsxPath"
Write-Log "  Per-type workbooks      : $TypeOutPath"
if ($HtmlPath) { Write-Log "  HTML dashboard          : $HtmlPath" }
Write-Log ('=' * 78) Section
if (-not $NoOpen -and $HtmlPath) { try { Invoke-Item -LiteralPath $HtmlPath } catch { Write-Log "Open the dashboard manually: $HtmlPath" Warning } }

#endregion
