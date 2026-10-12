<#
.SYNOPSIS
    Phase 1 - Azure Utilization, Performance & Cost Optimization: DATA COLLECTION (v2).

.DESCRIPTION
    Read-only collector that measures "provisioned vs. actually utilized" for EVERY billed Azure resource
    (IaaS, PaaS, databases, containers, networking, analytics/AI and consumption-based services) across all
    subscriptions in scope.

    Data sources (all read-only):
      - Azure Resource Graph ........ inventory, SKUs, ALL resource properties (flattened) and ALL tags,
                                      orphaned resources, Advisor (Cost + Performance)
      - ARM resources API ........... resource created / last-changed time (resource age)
      - Azure Monitor Metrics ....... hourly utilization for the lookback window via the Metrics Batch API
                                      (50x fewer calls, fallback to per-resource ARM metrics): avg / P95 / max,
                                      daily series, hour-of-week usage patterns, business vs off-hours
      - Generic metric discovery .... any resource type without a dedicated catalog entry gets its metric
                                      definitions discovered automatically (utilization % + activity metrics)
      - Log Analytics (Usage table) . daily billable ingestion per workspace and per table vs. commitment tiers,
                                      Sentinel enablement and Sentinel pricing tier
      - Cost Management Query API ... cost per resource, daily cost per service, monthly cost history
      - Capacity / BillingBenefits .. reservation & savings plan utilization (best effort)
      - Azure Retail Prices API ..... pay-as-you-go, reservation and savings plan list prices

    Everything is exported as CSV into <OutputPath>\01_RawData so that Phase 2
    (Phase2-AnalyzeAndReport.ps1) can build the Excel master report, per-type workbooks and the HTML dashboard.

.PARAMETER OutputPath
    Assessment folder. Raw data is written to <OutputPath>\01_RawData.

.PARAMETER TenantId
    Optional tenant to scope the assessment. Defaults to the tenant of the current Az context.

.PARAMETER SubscriptionIds
    Optional list of subscription IDs. Defaults to all enabled subscriptions in the tenant.

.PARAMETER LookbackDays
    Utilization window in full days (7-93, Azure Monitor keeps 93 days of platform metrics). Default 30.

.PARAMETER ThrottleLimit
    Parallel metric requests (PowerShell 7 ForEach-Object -Parallel). Default 12.

.PARAMETER MetricsBatchSize
    Resources per Azure Monitor Metrics Batch API call (1-50). Default 20. Use -NoBatchMetrics to force per-resource calls.

.PARAMETER CostType
    AmortizedCost (default - reservations/savings plans spread to resources) or ActualCost.

.PARAMETER CostHistoryMonths
    Months of monthly cost history per service (validates the impact of previous optimization). Default 6.

.PARAMETER TimeZoneId
    Time zone used for daily buckets and business-hours analysis. Default 'Arab Standard Time' (UTC+3, Qatar).

.PARAMETER WeekendDays
    Weekend days for usage-pattern analysis. Default Friday, Saturday (Qatar).

.PARAMETER BusinessHoursStart / BusinessHoursEnd
    Business hours (local, 24h clock) for usage-pattern analysis. Default 07:00-18:00.

.PARAMETER IncludeResourceTypes
    Optional subset of catalog keys (e.g. VirtualMachines, SqlDatabases). Default = all.

.PARAMETER ConfigPath
    Optional JSON configuration (default: CostOptimization.Config.json next to this script) to extend the
    assessment: generic metric overrides per resource type, excluded types, tag column priorities.

.PARAMETER SkipMetrics / SkipCostData / SkipLogAnalytics / SkipPricing / SkipGenericTypes
    Skip individual data sources (e.g. when permissions are missing).

.NOTES
    Requires : PowerShell 7.0+, Az.Accounts (or Azure CLI sign-in). Reader + Monitoring Reader (metrics),
               Log Analytics Reader (workspace queries) and Cost Management Reader (cost data).
    Version  : 2.0
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$TenantId,

    [string[]]$SubscriptionIds,

    [ValidateRange(7, 93)]
    [int]$LookbackDays = 30,

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 12,

    [ValidateSet('AmortizedCost', 'ActualCost')]
    [string]$CostType = 'AmortizedCost',

    [string]$TimeZoneId = 'Arab Standard Time',

    [ValidateSet('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')]
    [string[]]$WeekendDays = @('Friday', 'Saturday'),

    [ValidateRange(0, 23)]
    [int]$BusinessHoursStart = 7,

    [ValidateRange(1, 24)]
    [int]$BusinessHoursEnd = 18,

    [string[]]$IncludeResourceTypes,

    [string]$ConfigPath,

    [ValidateRange(1, 50)]
    [int]$MetricsBatchSize = 20,

    [ValidateRange(0, 36)]
    [int]$CostHistoryMonths = 6,

    [switch]$NoBatchMetrics,
    [switch]$SkipMetrics,
    [switch]$SkipCostData,
    [switch]$SkipLogAnalytics,
    [switch]$SkipPricing,
    [switch]$SkipGenericTypes
)

#region ============================== BOOTSTRAP ==============================

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
$script:ScriptVersion = '2.0'
$script:RunStart = Get-Date

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw "Az.Accounts module is required. Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts -ErrorAction Stop -WarningAction SilentlyContinue

$RawPath = Join-Path $OutputPath '01_RawData'
$TypesPath = Join-Path $RawPath 'ResourceTypes'
$DailyPath = Join-Path $RawPath 'DailyUtilization'
$PatternPath = Join-Path $RawPath 'UsagePatterns'
foreach ($p in @($OutputPath, $RawPath, $TypesPath, $DailyPath, $PatternPath)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}
$script:LogFile = Join-Path $OutputPath 'Assessment_Log.txt'
$script:CollectionErrors = [System.Collections.Generic.List[object]]::new()
$script:Utf8Bom = [System.Text.UTF8Encoding]::new($true)

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Section')]
        [string]$Level = 'Info'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Info' { 'Gray' } 'Warning' { 'Yellow' } 'Error' { 'Red' } 'Success' { 'Green' } 'Section' { 'Cyan' }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
    try { "[$ts] [$Level] $Message" | Out-File -FilePath $script:LogFile -Append -Encoding utf8 } catch { Write-Verbose "Log file write failed: $($_.Exception.Message)" }
}

function Add-CollectionError {
    param([string]$Area, [string]$Target, [string]$Message)
    $script:CollectionErrors.Add([pscustomobject]@{
            TimeUtc = [DateTime]::UtcNow.ToString('s'); Area = $Area; Target = $Target
            Message = if ($Message.Length -gt 600) { $Message.Substring(0, 600) } else { $Message }
        })
}

function Write-CsvFile {
    param([object[]]$Rows, [string]$Path)
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    # Normalize: union of columns (first-seen order) so every row exports with the same header
    $cols = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $Rows) {
        $names = if ($r -is [System.Collections.IDictionary]) { $r.Keys } else { $r.PSObject.Properties.Name }
        foreach ($n in $names) { if ($seen.Add([string]$n)) { $cols.Add([string]$n) } }
    }
    $out = foreach ($r in $Rows) {
        $o = [ordered]@{}
        foreach ($c in $cols) {
            $v = if ($r -is [System.Collections.IDictionary]) { $r[$c] } else { $r.$c }
            if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { $v = [Math]::Round([double]$v, 4) }
            elseif ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-dd HH:mm:ss') }
            elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = (@($v) -join '; ') }
            $o[$c] = $v
        }
        [pscustomobject]$o
    }
    $out | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8BOM
}

function Get-NameFromId([string]$Id) {
    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    return ($Id.TrimEnd('/') -split '/')[-1]
}

function Get-TagString($Tags) {
    if ($null -eq $Tags) { return '' }
    $pairs = foreach ($p in $Tags.PSObject.Properties) { "$($p.Name)=$($p.Value)" }
    return ($pairs -join '; ')
}

function Get-TagValue($Tags, [string[]]$Names) {
    if ($null -eq $Tags) { return $null }
    foreach ($p in $Tags.PSObject.Properties) {
        foreach ($n in $Names) { if ($p.Name -ieq $n) { return [string]$p.Value } }
    }
    return $null
}

function Get-ManagedByService($r) {
    $rg = [string]$r.resourceGroup
    $tagNames = if ($r.tags) { @($r.tags.PSObject.Properties.Name) } else { @() }
    if ($rg -like 'databricks-rg-*' -or (Get-TagValue $r.tags @('Vendor')) -eq 'Databricks' -or ($tagNames -contains 'databricks-instance-name')) { return 'Databricks' }
    if ($rg -like 'MC_*' -or ($tagNames | Where-Object { $_ -like 'aks-managed-*' })) { return 'AKS' }
    if ((Get-TagValue $r.tags @('cm-resource-parent')) -match 'hostpools') { return 'AVD' }
    if ($rg -like 'synapseworkspace-managedrg-*') { return 'Synapse' }
    if ($rg -like 'ME_*' -or $rg -like 'managed-rg-*') { return 'Managed Environment' }
    if (-not [string]::IsNullOrWhiteSpace([string]$r.managedBy)) { return "Managed: $((([string]$r.managedBy) -split '/providers/')[-1] -replace '/.*$', '')" }
    return $null
}

function ConvertTo-Bool($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [bool]) { return $v }
    $s = [string]$v
    if ($s -match '^(?i)(true|enabled|yes)$') { return $true }
    if ($s -match '^(?i)(false|disabled|no)$') { return $false }
    return $null
}

function Get-Round($v, [int]$Digits = 2) {
    if ($null -eq $v) { return $null }
    try { return [Math]::Round([double]$v, $Digits) } catch { return $null }
}

# ---- Configuration (optional JSON; script parameters always win)
# Built-in defaults (identical to the optional CostOptimization.Config.json) keep the script self-contained
$script:DefaultConfigJson = @'
{
  "_readme": "Optional configuration for the Azure Utilization & Cost Optimization assessment. Edit values to tune thresholds or extend coverage without changing the scripts. Explicit script parameters always win over this file.",

  "Thresholds": {
    "CpuIdle": 5,
    "CpuLow": 40,
    "CpuHigh": 85,
    "MemLow": 60,
    "MemHigh": 90,
    "TargetUtilization": 70,
    "GenericLowUtilization": 30,
    "GenericHighUtilization": 85,
    "MinMonthlyCostToReportGenericType": 1,
    "MinMonthlyCostForGenericRecommendation": 1,
    "SnapshotMaxAgeDays": 90,
    "SteadyStateHoursPct": 95
  },

  "Collection": {
    "MaxPropertyColumnsPerType": 400,
    "MaxTagColumnsPerType": 75,
    "PriorityTagKeys": [ "Environment", "Env", "Owner", "CostCenter", "Cost Center", "Application", "App", "Project", "BusinessUnit", "Department", "Workload", "Team", "Service", "Criticality" ],
    "GenericMetrics": {
      "Enabled": true,
      "MaxMetricsPerType": 10,
      "SampleResourcesPerType": 3
    },
    "ExcludedResourceTypes": [
      "microsoft.network/networkinterfaces",
      "microsoft.network/networksecuritygroups",
      "microsoft.network/routetables",
      "microsoft.network/applicationsecuritygroups",
      "microsoft.network/networkwatchers",
      "microsoft.network/privatednszones/virtualnetworklinks",
      "microsoft.compute/availabilitysets",
      "microsoft.compute/proximityplacementgroups",
      "microsoft.compute/sshpublickeys",
      "microsoft.compute/virtualmachines/extensions",
      "microsoft.compute/virtualmachinescalesets/extensions",
      "microsoft.hybridcompute/machines/extensions",
      "microsoft.managedidentity/userassignedidentities",
      "microsoft.insights/workbooks",
      "microsoft.insights/datacollectionrules",
      "microsoft.insights/datacollectionendpoints",
      "microsoft.portal/dashboards",
      "microsoft.resources/templatespecs",
      "microsoft.resources/templatespecs/versions",
      "microsoft.web/connections",
      "microsoft.alertsmanagement/smartdetectoralertrules",
      "microsoft.operationsmanagement/solutions",
      "microsoft.sql/servers",
      "microsoft.sql/servers/firewallrules",
      "microsoft.dbforpostgresql/flexibleservers/databases",
      "microsoft.security/automations",
      "microsoft.visualstudio/account"
    ],
    "MetricOverrides": {
      "microsoft.signalrservice/signalr": [ "ServerLoad", "ConnectionCount", "MessageCount" ],
      "microsoft.signalrservice/webpubsub": [ "ServerLoad", "ConnectionCount", "OutboundTraffic" ],
      "microsoft.keyvault/vaults": [ "ServiceApiHit", "ServiceApiLatency", "SaturationShoebox" ],
      "microsoft.eventgrid/topics": [ "PublishSuccessCount", "DeliverySuccessCount", "DroppedEventCount" ],
      "microsoft.eventgrid/systemtopics": [ "PublishSuccessCount", "DeliverySuccessCount" ],
      "microsoft.automation/automationaccounts": [ "TotalJob", "TotalUpdateDeploymentRuns" ],
      "microsoft.network/dnszones": [ "QueryVolume", "RecordSetCount" ],
      "microsoft.network/trafficmanagerprofiles": [ "QpsByEndpoint", "ProbeAgentCurrentEndpointStateByProfileResourceId" ],
      "microsoft.appconfiguration/configurationstores": [ "HttpIncomingRequestCount", "ThrottledHttpRequestCount" ],
      "microsoft.relay/namespaces": [ "ListenerConnections-Success", "ActiveConnections" ],
      "microsoft.notificationhubs/namespaces/notificationhubs": [ "incoming", "outgoing.allpns.success" ],
      "microsoft.avs/privateclouds": [ "EffectiveCpuAverage", "UsageAverage", "DiskUsedPercentage" ],
      "microsoft.network/virtualnetworks": []
    },
    "CatalogExtraMetrics": {
      "_example_VirtualMachines": [ { "Name": "Inbound Flows", "Alias": "InboundFlows", "Out": "Avg,Max" } ]
    }
  },

  "Report": {
    "_readme": "HTML dashboard sizing. Properties per type and daily series are reduced automatically for very large estates to keep the file responsive.",
    "HtmlMaxPropertiesPerType": 100,
    "HtmlMaxDailyValues": 1500000,
    "EChartsUrl": "https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"
  },

  "Environment": {
    "ProductionTagPattern": "(?i)^(prod|production|prd|live|prod-.*)$",
    "NonProductionTagPattern": "(?i)(dev|test|uat|qa|stag|sandbox|sbx|poc|demo|non-?prod|pre-?prod|lab|train|trial)",
    "NonProductionNamePattern": "(?<![a-z])(dev|devtest|test|tst|uat|qa|sit|stg|stage|staging|sandbox|sbx|poc|demo|nonprod|non-prod|np|preprod|pre-prod|lab|training)(?![a-z])",
    "ProductionNamePattern": "(?<![a-z])(prod|prd|production|live)(?![a-z])"
  }
}
'@
if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'CostOptimization.Config.json' }
$script:Config = $null
if (Test-Path $ConfigPath) {
    try { $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable; Write-Log "Configuration loaded: $ConfigPath" }
    catch { Write-Log "Configuration file could not be parsed ($ConfigPath): $($_.Exception.Message) - using defaults" Warning }
}
elseif ($PSBoundParameters.ContainsKey('ConfigPath')) { Write-Log "Configuration file not found: $ConfigPath - using built-in defaults" Warning }
if ($null -eq $script:Config) { $script:Config = $script:DefaultConfigJson | ConvertFrom-Json -AsHashtable }
function Get-Cfg([string]$Path, $Default) {
    $node = $script:Config
    foreach ($p in ($Path -split '\.')) {
        if ($node -is [System.Collections.IDictionary] -and $node.Contains($p)) { $node = $node[$p] } else { return $Default }
    }
    if ($null -eq $node) { return $Default }
    return $node
}
$script:MaxPropColumns = [int](Get-Cfg 'Collection.MaxPropertyColumnsPerType' 400)
$script:MaxTagColumns = [int](Get-Cfg 'Collection.MaxTagColumnsPerType' 75)
$script:PriorityTagKeys = @(Get-Cfg 'Collection.PriorityTagKeys' @('Environment', 'Env', 'Owner', 'CostCenter', 'Application', 'Project'))
$script:ExcludedTypes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($t in @(Get-Cfg 'Collection.ExcludedResourceTypes' @())) { [void]$script:ExcludedTypes.Add([string]$t) }
$script:MetricOverrides = @{}
$mo = Get-Cfg 'Collection.MetricOverrides' @{}
foreach ($k in @($mo.Keys)) { if (-not ([string]$k).StartsWith('_')) { $script:MetricOverrides[([string]$k).ToLower()] = @($mo[$k]) } }
$script:CatalogExtraMetrics = Get-Cfg 'Collection.CatalogExtraMetrics' @{}
$script:GenericEnabled = [bool](Get-Cfg 'Collection.GenericMetrics.Enabled' $true) -and -not $SkipGenericTypes
$script:GenericMaxMetrics = [int](Get-Cfg 'Collection.GenericMetrics.MaxMetricsPerType' 10)
$script:GenericSamples = [int](Get-Cfg 'Collection.GenericMetrics.SampleResourcesPerType' 3)
$script:GenericMinCost = [double](Get-Cfg 'Thresholds.MinMonthlyCostToReportGenericType' 1)

# ---- Full property flattening (every ARM property becomes a filterable column: Prop.<path>)
function Add-FlatProperties([System.Collections.IDictionary]$Target, $Obj, [string]$Prefix, [int]$Depth = 0) {
    if ($null -eq $Obj) { return }
    if ($Obj -is [string] -or $Obj -is [ValueType]) { Set-FlatValue $Target $Prefix $Obj; return }
    if ($Obj -is [System.Collections.IDictionary]) {
        foreach ($k in @($Obj.Keys)) { Add-FlatChild $Target $Obj[$k] ([string]$k) $Prefix $Depth }
        return
    }
    if ($Obj -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $Obj.PSObject.Properties) { Add-FlatChild $Target $p.Value $p.Name $Prefix $Depth }
        return
    }
    if ($Obj -is [System.Collections.IEnumerable]) {
        $items = @($Obj)
        if ($items.Count -eq 0) { return }
        $scalarCount = @($items | Where-Object { $null -eq $_ -or $_ -is [string] -or $_ -is [ValueType] }).Count
        if ($scalarCount -eq $items.Count) { Set-FlatValue $Target $Prefix (($items | ForEach-Object { [string]$_ }) -join '; ') }
        else {
            $Target["$Prefix.count"] = $items.Count
            $json = try { ConvertTo-Json -InputObject $items -Depth 4 -Compress -WarningAction SilentlyContinue } catch { '' }
            Set-FlatValue $Target $Prefix $json
        }
        return
    }
    Set-FlatValue $Target $Prefix ([string]$Obj)
}
function Add-FlatChild([System.Collections.IDictionary]$Target, $Val, [string]$Name, [string]$Prefix, [int]$Depth) {
    if ($null -eq $Val) { return }
    if ($Name -match '(?i)(password|secret|connectionstring|accesskey|sharedkey|adminusername|sastoken|primarykey|secondarykey|^keys?$)') { return }
    $path = if ($Prefix) { "$Prefix.$Name" } else { $Name }
    if ($Depth -ge 5 -and -not ($Val -is [string] -or $Val -is [ValueType])) {
        $json = try { ConvertTo-Json -InputObject $Val -Depth 3 -Compress -WarningAction SilentlyContinue } catch { '' }
        Set-FlatValue $Target $path $json
        return
    }
    Add-FlatProperties $Target $Val $path ($Depth + 1)
}
function Set-FlatValue([System.Collections.IDictionary]$Target, [string]$Path, $Value) {
    if ($null -eq $Value) { return }
    if ($Value -is [DateTime]) { $Value = $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    elseif ($Value -is [bool]) { $Value = if ($Value) { 'True' } else { 'False' } }
    elseif ($Value -is [string]) { if ($Value.Length -eq 0) { return }; if ($Value.Length -gt 1000) { $Value = $Value.Substring(0, 1000) + '...' } }
    $Target[$Path] = $Value
}

function Add-ResourceDetails([System.Collections.IDictionary]$Row, $r) {
    $Row.Kind = [string]$r.kind
    $Row.SkuName = [string]$r.sku.name
    $Row.SkuTier = [string]$r.sku.tier
    $Row.SkuCapacity = $r.sku.capacity
    $Row.SkuFamily = [string]$r.sku.family
    $times = $script:ResourceTimes[([string]$r.id).ToLower()]
    $Row.CreatedTime = if ($times) { $times.Created } else { $null }
    $Row.LastChangedTime = if ($times) { $times.Changed } else { $null }
    $Row.AgeDays = if ($times -and $times.Created) { Get-AgeDays $times.Created } else { $null }
    $Row.ManagedBy = [string]$r.managedBy
    $Row.IdentityType = [string]$r.identity.type
    $Row.ResourceGroupTags = $script:RgTags["$($r.subscriptionId)|$($r.resourceGroup)".ToLower()]
    if ($r.tags) {
        foreach ($p in $r.tags.PSObject.Properties) {
            $k = ([string]$p.Name).Trim()
            if ($k) { $Row["Tag.$k"] = [string]$p.Value }
        }
    }
    $flat = [ordered]@{}
    Add-FlatProperties $flat $r.properties ''
    foreach ($k in $flat.Keys) { $Row["Prop.$k"] = $flat[$k] }
    if ($r.plan) { $planFlat = [ordered]@{}; Add-FlatProperties $planFlat $r.plan ''; foreach ($k in $planFlat.Keys) { $Row["Plan.$k"] = $planFlat[$k] } }
}

# Resource created / changed times (ARM resources API with $expand) - resource age analysis
$script:ResourceTimes = @{}

#endregion

#region ============================== TIME WINDOW ==============================

try { $script:Tz = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId) }
catch { Write-Log "Time zone '$TimeZoneId' not found - using UTC" Warning; $script:Tz = [TimeZoneInfo]::Utc; $TimeZoneId = 'UTC' }

$nowLocal = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $script:Tz)
$EndLocal = [DateTime]::SpecifyKind($nowLocal.Date, [DateTimeKind]::Unspecified)       # today 00:00 local (exclusive)
$StartLocal = $EndLocal.AddDays(-$LookbackDays)
$StartUtc = [TimeZoneInfo]::ConvertTimeToUtc($StartLocal, $script:Tz)
$EndUtc = [TimeZoneInfo]::ConvertTimeToUtc($EndLocal, $script:Tz)
$TotalHours = [int][Math]::Round(($EndUtc - $StartUtc).TotalHours)
$Timespan = "$($StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'))/$($EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
$TzOffsetHours = $script:Tz.GetUtcOffset([DateTime]::UtcNow).TotalHours

# Pre-computed local calendar for every hour bucket (shared read-only with parallel workers)
$HourLocalHour = [int[]]::new($TotalHours)
$HourLocalDow = [int[]]::new($TotalHours)
$HourDayIndex = [int[]]::new($TotalHours)
$DayLabels = [string[]]::new($LookbackDays)
for ($d = 0; $d -lt $LookbackDays; $d++) { $DayLabels[$d] = $StartLocal.AddDays($d).ToString('yyyy-MM-dd') }
for ($i = 0; $i -lt $TotalHours; $i++) {
    $lt = [TimeZoneInfo]::ConvertTimeFromUtc($StartUtc.AddHours($i), $script:Tz)
    $HourLocalHour[$i] = $lt.Hour
    $HourLocalDow[$i] = [int]$lt.DayOfWeek
    $HourDayIndex[$i] = [Math]::Min($LookbackDays - 1, [Math]::Max(0, [int][Math]::Floor(($lt - $StartLocal).TotalDays)))
}
$WeekendDowInts = @($WeekendDays | ForEach-Object { [int][System.DayOfWeek]$_ })

#endregion

#region ============================== AUTH & REST ==============================

# Authentication: Az PowerShell context first, Azure CLI ("az login") as fallback.
$script:AzContext = Get-AzContext -ErrorAction SilentlyContinue
$script:AzCli = Get-Command az -ErrorAction SilentlyContinue
$script:CliAccount = $null
if ($script:AzCli) {
    try { $script:CliAccount = (& az account show -o json 2>$null | Out-String | ConvertFrom-Json) } catch { $script:CliAccount = $null }
}
if (-not $script:AzContext -and -not $script:CliAccount) { throw "Not signed in to Azure. Run Connect-AzAccount (or az login) first." }
$envInfo = if ($script:AzContext) { $script:AzContext.Environment } else { $null }
$script:ArmBase = ([string]$envInfo.ResourceManagerUrl).TrimEnd('/')
if ([string]::IsNullOrWhiteSpace($script:ArmBase)) { $script:ArmBase = 'https://management.azure.com' }
$script:ArmResource = "$script:ArmBase/"
$script:LaResource = [string]$envInfo.AzureOperationalInsightsEndpointResourceId
if ([string]::IsNullOrWhiteSpace($script:LaResource)) { $script:LaResource = 'https://api.loganalytics.io' }
$script:LaEndpoint = ([string]$envInfo.AzureOperationalInsightsEndpoint).TrimEnd('/')
if ([string]::IsNullOrWhiteSpace($script:LaEndpoint)) { $script:LaEndpoint = 'https://api.loganalytics.io/v1' }
if (-not $TenantId) { $TenantId = if ($script:AzContext) { [string]$script:AzContext.Tenant.Id } else { [string]$script:CliAccount.tenantId } }
$script:MetricsResource = 'https://metrics.monitor.azure.com'
$script:BatchMetrics = ($script:ArmBase -eq 'https://management.azure.com') -and -not $NoBatchMetrics
$script:MetricRequests = 0
$script:TokenCache = @{}
$script:AuthSource = $null

function Get-BearerToken {
    param([string]$Resource = $script:ArmResource, [switch]$Force)
    $c = $script:TokenCache[$Resource]
    if (-not $Force -and $c -and $c.ExpiresOn -gt [DateTimeOffset]::UtcNow.AddMinutes(10)) { return $c.Token }
    $tok = $null; $exp = $null; $err = $null
    if ($script:AzContext -and $script:AuthSource -ne 'AzureCLI') {
        try {
            $t = Get-AzAccessToken -ResourceUrl $Resource -TenantId $TenantId -ErrorAction Stop -WarningAction SilentlyContinue
            $tok = $t.Token
            if ($tok -is [System.Security.SecureString]) { $tok = [System.Net.NetworkCredential]::new('', $tok).Password }
            $exp = if ($t.ExpiresOn) { [DateTimeOffset]$t.ExpiresOn } else { [DateTimeOffset]::UtcNow.AddMinutes(45) }
            if (-not $script:AuthSource) { $script:AuthSource = 'AzPowerShell' }
        }
        catch { $err = $_.Exception.Message; $tok = $null }
    }
    if (-not $tok -and $script:AzCli) {
        $raw = & az account get-access-token --resource $Resource --tenant $TenantId -o json 2>$null | Out-String
        if ($LASTEXITCODE -eq 0 -and $raw) {
            $j = $raw | ConvertFrom-Json
            $tok = $j.accessToken
            $exp = if ($j.expires_on) { [DateTimeOffset]::FromUnixTimeSeconds([long]$j.expires_on) } else { [DateTimeOffset][DateTime]::Parse($j.expiresOn) }
            if ($script:AuthSource -ne 'AzureCLI') { Write-Log "Using Azure CLI credentials for $Resource$(if ($err) { ' (Az PowerShell token unavailable)' })" Warning }
            $script:AuthSource = 'AzureCLI'
        }
    }
    if (-not $tok) { throw "Unable to acquire an access token for $Resource. Run Connect-AzAccount -TenantId $TenantId (or az login). $err" }
    $script:TokenCache[$Resource] = @{ Token = $tok; ExpiresOn = $exp }
    return $tok
}

function Get-RetryAfterSeconds($Headers, [int]$Attempt) {
    $wait = [Math]::Min(60, [Math]::Pow(2, $Attempt) + (Get-Random -Minimum 0 -Maximum 3))
    if ($Headers) {
        foreach ($k in @($Headers.Keys)) {
            if ($k -match 'retry-after') {
                $val = 0
                if ([int]::TryParse((@($Headers[$k])[0]), [ref]$val)) { $wait = [Math]::Max($wait, [Math]::Min($val, 300)) }
            }
        }
    }
    return [int]$wait
}

function Invoke-AzApi {
    <# Robust REST call (token refresh, 429/5xx retry with Retry-After). Throws on non-retryable errors. #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [string]$Resource = $script:ArmResource,
        [int]$MaxRetries = 6
    )
    $attempt = 0
    while ($true) {
        $attempt++
        $sc = 0; $rh = $null; $resp = $null; $netErr = $null
        $params = @{
            Uri = $Uri; Method = $Method; Headers = @{ Authorization = "Bearer $(Get-BearerToken -Resource $Resource)" }
            SkipHttpErrorCheck = $true; StatusCodeVariable = 'sc'; ResponseHeadersVariable = 'rh'; TimeoutSec = 300
        }
        if ($null -ne $Body) {
            $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 30 -Compress }
            $params.ContentType = 'application/json'
        }
        try { $resp = Invoke-RestMethod @params } catch { $netErr = $_.Exception.Message; $sc = 0 }
        if ($sc -ge 200 -and $sc -lt 300) { return $resp }
        if ($sc -eq 401 -and $attempt -le 2) { $null = Get-BearerToken -Resource $Resource -Force; continue }
        if (($sc -eq 0 -or $sc -eq 429 -or $sc -ge 500) -and $attempt -le $MaxRetries) {
            Start-Sleep -Seconds (Get-RetryAfterSeconds $rh $attempt)
            continue
        }
        $msg = if ($resp -and $resp.error) { "$($resp.error.code): $($resp.error.message)" }
        elseif ($resp -and $resp.code) { "$($resp.code): $($resp.message)" }
        elseif ($netErr) { $netErr } else { "HTTP $sc" }
        throw [System.Exception]::new("HTTP $sc :: $msg")
    }
}

function Get-AllPages {
    <# GET with nextLink pagination (ARM list pattern). #>
    param([string]$Uri, [string]$Resource = $script:ArmResource, [int]$MaxPages = 200)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri; $page = 0
    while ($next -and $page -lt $MaxPages) {
        $page++
        $resp = Invoke-AzApi -Uri $next -Resource $Resource
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.nextLink
    }
    return $items.ToArray()
}

function Invoke-ARGQuery {
    <# Azure Resource Graph via REST, chunked by subscription, paginated with skipToken. #>
    param([Parameter(Mandatory)][string]$Query, [string[]]$Subscriptions = $script:SubIds)
    $all = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Subscriptions.Count; $i += 200) {
        $chunk = @($Subscriptions[$i..([Math]::Min($i + 199, $Subscriptions.Count - 1))])
        $skip = $null
        do {
            $opts = @{ resultFormat = 'objectArray'; '$top' = 1000 }
            if ($skip) { $opts['$skipToken'] = $skip }
            $body = @{ subscriptions = $chunk; query = $Query; options = $opts }
            $resp = Invoke-AzApi -Uri "$script:ArmBase/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01" -Method POST -Body $body
            if ($resp.data) { foreach ($row in $resp.data) { $all.Add($row) } }
            $skip = $resp.'$skipToken'
        } while ($skip)
    }
    return $all.ToArray()
}

#endregion

#region ============================== PARALLEL WORKERS ==============================

# Metric worker: Azure Monitor Metrics Batch API (up to 50 resources per call) with per-resource ARM fallback -> statistics.
# Self-contained (runs inside ForEach-Object -Parallel runspaces), PS7 only.
$script:MetricWorkerText = @'
param($Item, $Ctx)
$ErrorActionPreference = 'Stop'
function Invoke-Http([string]$Url, [string]$Method, [string]$Body, [string]$Token) {
    $attempt = 0
    while ($true) {
        $attempt++
        $sc = 0; $rh = $null; $resp = $null; $netErr = $null
        $p = @{ Uri = $Url; Method = $Method; Headers = @{ Authorization = ('Bearer ' + $Token) }; SkipHttpErrorCheck = $true; StatusCodeVariable = 'sc'; ResponseHeadersVariable = 'rh'; TimeoutSec = 120 }
        if ($Body) { $p.Body = $Body; $p.ContentType = 'application/json' }
        try { $resp = Invoke-RestMethod @p } catch { $netErr = $_.Exception.Message; $sc = 0 }
        if ($sc -ge 200 -and $sc -lt 300) { return @{ Ok = $true; Data = $resp } }
        if ((($sc -eq 429 -or $sc -ge 500) -and $attempt -le 5) -or ($sc -eq 0 -and $attempt -le 1)) {
            $wait = [Math]::Min(60, [Math]::Pow(2, $attempt))
            if ($rh) { foreach ($k in @($rh.Keys)) { if ($k -match 'retry-after') { $v = 0; if ([int]::TryParse((@($rh[$k])[0]), [ref]$v)) { $wait = [Math]::Max($wait, [Math]::Min($v, 120)) } } } }
            Start-Sleep -Seconds $wait
            continue
        }
        $msg = if ($resp -and $resp.error) { "$($resp.error.code): $($resp.error.message)" } elseif ($resp -and $resp.code) { "$($resp.code): $($resp.message)" } elseif ($netErr) { $netErr } else { "HTTP $sc" }
        return @{ Ok = $false; Status = $sc; Error = $msg }
    }
}
function Get-Pct([double[]]$sorted, [double]$p) {
    if ($sorted.Length -eq 0) { return $null }
    $idx = [Math]::Ceiling($p * $sorted.Length) - 1
    if ($idx -lt 0) { $idx = 0 }; if ($idx -ge $sorted.Length) { $idx = $sorted.Length - 1 }
    return $sorted[$idx]
}
function New-Acc([string]$Id) { @{ Id = $Id.ToLower(); Series = @{}; Errors = [System.Collections.Generic.List[string]]::new(); Calls = 0; CallsOk = 0 } }

# Hourly series -> statistics, daily aggregates and hour-of-week usage pattern
function Add-MetricSeries([hashtable]$Series, $MetricValues, [hashtable]$Specs) {
    $NH = $Ctx.TotalHours; $ND = $Ctx.Days
    foreach ($m in @($MetricValues)) {
        if ($null -eq $m) { continue }
        $spec = $Specs[[string]$m.name.value]
        if (-not $spec) { continue }
        foreach ($ts in @($m.timeseries)) {
            if ($null -eq $ts) { continue }
            $dimVals = @($ts.metadatavalues | Where-Object { $_ } | ForEach-Object { [string]$_.value })
            $key = if ($dimVals.Count -gt 0) { "$($spec.Alias)|$($dimVals -join '/')" } else { $spec.Alias }
            $scale = [double]$spec.Scale; if ($scale -eq 0) { $scale = 1 }
            $A = [double[]]::new($NH); $X = [double[]]::new($NH); $N = [double[]]::new($NH); $T = [double[]]::new($NH)
            for ($i = 0; $i -lt $NH; $i++) { $A[$i] = [double]::NaN; $X[$i] = [double]::NaN; $N[$i] = [double]::NaN; $T[$i] = [double]::NaN }
            $any = $false
            foreach ($pt in @($ts.data)) {
                if ($null -eq $pt) { continue }
                $tsv = $pt.timeStamp
                $dt = if ($tsv -is [DateTime]) { $tsv.ToUniversalTime() } else { [DateTime]::Parse([string]$tsv, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
                $idx = [int][Math]::Floor(($dt - $Ctx.StartUtc).TotalHours)
                if ($idx -lt 0 -or $idx -ge $NH) { continue }
                if ($null -ne $pt.average) { $A[$idx] = [double]$pt.average / $scale; $any = $true }
                if ($null -ne $pt.maximum) { $X[$idx] = [double]$pt.maximum / $scale; $any = $true }
                if ($null -ne $pt.minimum) { $N[$idx] = [double]$pt.minimum / $scale }
                if ($null -ne $pt.total) { $T[$idx] = [double]$pt.total / $scale; $any = $true }
            }
            if (-not $any) { continue }
            $B = if ($spec.Basis -eq 'Max') { $X } elseif ($spec.Basis -eq 'Total') { $T } else { $A }
            $bl = [System.Collections.Generic.List[double]]::new()
            $mx = [double]::NaN; $mn = [double]::NaN; $tot = 0.0; $totCnt = 0; $maxT = [double]::NaN
            for ($i = 0; $i -lt $NH; $i++) {
                if (-not [double]::IsNaN($B[$i])) { $bl.Add($B[$i]) }
                if (-not [double]::IsNaN($X[$i])) { if ([double]::IsNaN($mx) -or $X[$i] -gt $mx) { $mx = $X[$i] } }
                if (-not [double]::IsNaN($N[$i])) { if ([double]::IsNaN($mn) -or $N[$i] -lt $mn) { $mn = $N[$i] } }
                if (-not [double]::IsNaN($T[$i])) { $tot += $T[$i]; $totCnt++; if ([double]::IsNaN($maxT) -or $T[$i] -gt $maxT) { $maxT = $T[$i] } }
            }
            $sorted = $bl.ToArray(); [Array]::Sort($sorted)
            $avg = if ($sorted.Length) { ($sorted | Measure-Object -Average).Average } else { $null }
            $st = @{
                Avg = $avg; P95 = (Get-Pct $sorted 0.95); P05 = (Get-Pct $sorted 0.05); Med = (Get-Pct $sorted 0.5); P99 = (Get-Pct $sorted 0.99)
                MinAvg = $(if ($sorted.Length) { $sorted[0] } else { $null })
                Max = if ([double]::IsNaN($mx)) { $null } else { $mx }
                Min = if ([double]::IsNaN($mn)) { $null } else { $mn }
                Total = if ($totCnt) { $tot } else { $null }
                MaxT = if ([double]::IsNaN($maxT)) { $null } else { $maxT }
                Points = $sorted.Length; Unit = [string]$m.unit
            }
            $dA = [double[]]::new($ND); $dC = [int[]]::new($ND); $dX = [double[]]::new($ND); $dN = [double[]]::new($ND); $dT = [double[]]::new($ND); $dTC = [int[]]::new($ND)
            for ($d = 0; $d -lt $ND; $d++) { $dX[$d] = [double]::NaN; $dN[$d] = [double]::NaN }
            for ($i = 0; $i -lt $NH; $i++) {
                $d = $Ctx.DayIdx[$i]
                if (-not [double]::IsNaN($B[$i])) { $dA[$d] += $B[$i]; $dC[$d]++ }
                if (-not [double]::IsNaN($X[$i])) { if ([double]::IsNaN($dX[$d]) -or $X[$i] -gt $dX[$d]) { $dX[$d] = $X[$i] } }
                if (-not [double]::IsNaN($N[$i])) { if ([double]::IsNaN($dN[$d]) -or $N[$i] -lt $dN[$d]) { $dN[$d] = $N[$i] } }
                if (-not [double]::IsNaN($T[$i])) { $dT[$d] += $T[$i]; $dTC[$d]++ }
            }
            $daily = [System.Collections.Generic.List[object]]::new()
            for ($d = 0; $d -lt $ND; $d++) {
                if ($dC[$d] -eq 0 -and $dTC[$d] -eq 0) { continue }
                $daily.Add(@($d, $(if ($dC[$d]) { $dA[$d] / $dC[$d] } else { $null }), $(if ([double]::IsNaN($dX[$d])) { $null } else { $dX[$d] }), $(if ([double]::IsNaN($dN[$d])) { $null } else { $dN[$d] }), $(if ($dTC[$d]) { $dT[$d] } else { $null }), $dC[$d]))
            }
            $st.Daily = $daily
            if ($spec.Pattern) {
                $P = if ($spec.PatternAgg -eq 'Total') { $T } else { $B }
                $howS = [double[]]::new(168); $howC = [int[]]::new(168)
                $bizS = 0.0; $bizC = 0; $offS = 0.0; $offC = 0; $wkS = 0.0; $wkC = 0; $weS = 0.0; $weC = 0; $idle = 0; $cnt = 0
                $hS = [double[]]::new(24); $hC = [int[]]::new(24); $dowS = [double[]]::new(7); $dowC = [int[]]::new(7)
                for ($i = 0; $i -lt $NH; $i++) {
                    $v = $P[$i]; if ([double]::IsNaN($v)) { continue }
                    $hr = $Ctx.LocalHour[$i]; $dw = $Ctx.LocalDow[$i]; $cnt++
                    $howS[$dw * 24 + $hr] += $v; $howC[$dw * 24 + $hr]++
                    $hS[$hr] += $v; $hC[$hr]++; $dowS[$dw] += $v; $dowC[$dw]++
                    $isWe = $Ctx.WeekendDows -contains $dw
                    if ($isWe) { $weS += $v; $weC++ } else { $wkS += $v; $wkC++ }
                    if (-not $isWe -and $hr -ge $Ctx.BizStart -and $hr -lt $Ctx.BizEnd) { $bizS += $v; $bizC++ } else { $offS += $v; $offC++ }
                    if ($spec.Idle -ge 0 -and $v -le $spec.Idle) { $idle++ }
                }
                $peakH = $null; $peakV = [double]::MinValue
                for ($h2 = 0; $h2 -lt 24; $h2++) { if ($hC[$h2] -gt 0 -and ($hS[$h2] / $hC[$h2]) -gt $peakV) { $peakV = $hS[$h2] / $hC[$h2]; $peakH = $h2 } }
                $peakD = $null; $peakDV = [double]::MinValue
                for ($d2 = 0; $d2 -lt 7; $d2++) { if ($dowC[$d2] -gt 0 -and ($dowS[$d2] / $dowC[$d2]) -gt $peakDV) { $peakDV = $dowS[$d2] / $dowC[$d2]; $peakD = $d2 } }
                $st.Pattern = @{
                    BizAvg = if ($bizC) { $bizS / $bizC } else { $null }; OffAvg = if ($offC) { $offS / $offC } else { $null }
                    WeekdayAvg = if ($wkC) { $wkS / $wkC } else { $null }; WeekendAvg = if ($weC) { $weS / $weC } else { $null }
                    PeakHour = $peakH; PeakDow = $peakD
                    IdlePct = if ($cnt -and $spec.Idle -ge 0) { 100.0 * $idle / $cnt } else { $null }
                    ActivePct = 100.0 * $cnt / $NH
                    HowSum = $howS; HowCnt = $howC
                }
            }
            $Series[$key] = $st
        }
    }
}

# Per-resource ARM metrics (fallback / global resources); retries metric-by-metric when a combined call is rejected
function Invoke-SingleResource([string]$Id, $Calls) {
    $acc = New-Acc $Id
    foreach ($call in @($Calls)) {
        $urlBase = "$($Ctx.ArmBase)$Id$($call.SubPath)/providers/Microsoft.Insights/metrics?api-version=2023-10-01&timespan=$($Ctx.Timespan)&interval=PT1H&aggregation=Average,Maximum,Minimum,Total"
        if ($call.Filter) { $urlBase += '&$filter=' + [uri]::EscapeDataString($call.Filter) + '&top=250' }
        $joined = ($call.Names | ForEach-Object { [uri]::EscapeDataString($_) }) -join ','
        $acc.Calls++
        $r = Invoke-Http "$urlBase&metricnames=$joined" 'GET' $null $Ctx.ArmToken
        if ($r.Ok) { $acc.CallsOk++; Add-MetricSeries $acc.Series $r.Data.value $call.Specs }
        elseif ($r.Status -eq 400 -and @($call.Names).Count -gt 1) {
            foreach ($oneName in $call.Names) {
                $acc.Calls++
                $r2 = Invoke-Http "$urlBase&metricnames=$([uri]::EscapeDataString($oneName))" 'GET' $null $Ctx.ArmToken
                if ($r2.Ok) { $acc.CallsOk++; Add-MetricSeries $acc.Series $r2.Data.value $call.Specs } else { $acc.Errors.Add("$oneName -> $($r2.Error)") }
            }
        }
        else { $acc.Errors.Add("$((@($call.Names) -join ',')) -> $($r.Error)") }
    }
    return $acc
}

$results = [System.Collections.Generic.List[object]]::new()
$requests = 0
if ($Item.Mode -eq 'Batch') {
    $names = (@($Item.Names) | ForEach-Object { [uri]::EscapeDataString($_) }) -join ','
    $url = "https://$($Item.Region).metrics.monitor.azure.com/subscriptions/$($Item.SubscriptionId)/metrics:getBatch?starttime=$($Ctx.StartIso)&endtime=$($Ctx.EndIso)&interval=PT1H&metricnamespace=$([uri]::EscapeDataString($Item.Namespace))&metricnames=$names&aggregation=average,maximum,minimum,total&api-version=2023-10-01"
    if ($Item.Filter) { $url += '&filter=' + [uri]::EscapeDataString($Item.Filter) + '&top=250' }
    $body = @{ resourceids = @($Item.ResourceIds) } | ConvertTo-Json -Compress
    $requests++
    $r = if ($Ctx.MetricsToken) { Invoke-Http $url 'POST' $body $Ctx.MetricsToken } else { @{ Ok = $false; Status = 0; Error = 'No metrics token' } }
    if ($r.Ok) {
        $accs = @{}
        foreach ($bid in @($Item.OriginalIds)) { $a = New-Acc $bid; $a.Calls = 1; $a.CallsOk = 1; $accs[$a.Id] = $a }
        $sp = ([string]$Item.SubPath).ToLower()
        foreach ($v in @($r.Data.values)) {
            if ($null -eq $v) { continue }
            $rid = ([string]$v.resourceid).ToLower()
            if ($sp -and $rid.EndsWith($sp)) { $rid = $rid.Substring(0, $rid.Length - $sp.Length) }
            $a = $accs[$rid]
            if (-not $a) { continue }
            if ($v.PSObject.Properties['error'] -and $v.error) { $a.Errors.Add("batch -> $($v.error.code): $($v.error.message)"); $a.CallsOk = 0; continue }
            Add-MetricSeries $a.Series $v.value $Item.Specs
        }
        foreach ($a in $accs.Values) { $results.Add($a) }
    }
    elseif ($r.Status -eq 0) {
        # Timeout / network failure: record the error instead of multiplying slow calls per resource
        foreach ($bid in @($Item.OriginalIds)) { $a = New-Acc $bid; $a.Calls = 1; $a.Errors.Add("batch timeout -> $($r.Error)"); $results.Add($a) }
    }
    else {
        # Batch rejected (unsupported metric for some resources, permissions, region...) -> per-resource ARM metrics
        $call = @{ SubPath = $Item.SubPath; Filter = $Item.Filter; Names = $Item.Names; Specs = $Item.Specs }
        foreach ($bid in @($Item.OriginalIds)) {
            $acc = Invoke-SingleResource $bid @($call)
            $requests += $acc.Calls
            $results.Add($acc)
        }
    }
}
else {
    $acc = Invoke-SingleResource $Item.Id $Item.Calls
    $requests += $acc.Calls
    $results.Add($acc)
}
[pscustomobject]@{ Requests = $requests; Results = $results.ToArray() }
'@

# Generic REST worker (GET/POST) for parallel ARM / Log Analytics calls
$script:RestWorkerText = @'
param($Item, $Ctx)
$attempt = 0
while ($true) {
    $attempt++
    $sc = 0; $rh = $null; $resp = $null; $netErr = $null
    $p = @{ Uri = $Item.Url; Method = $Item.Method; Headers = @{ Authorization = "Bearer $($Ctx.Tokens[$Item.TokenKey])" }; SkipHttpErrorCheck = $true; StatusCodeVariable = 'sc'; ResponseHeadersVariable = 'rh'; TimeoutSec = 120 }
    if ($Item.Body) { $p.Body = $Item.Body; $p.ContentType = 'application/json' }
    try { $resp = Invoke-RestMethod @p } catch { $netErr = $_.Exception.Message; $sc = 0 }
    if ($sc -ge 200 -and $sc -lt 300) { return [pscustomobject]@{ Key = $Item.Key; Ok = $true; Status = $sc; Data = $resp; Error = $null } }
    if ((($sc -eq 429 -or $sc -ge 500) -and $attempt -le 5) -or ($sc -eq 0 -and $attempt -le 1)) {
        $wait = [Math]::Min(60, [Math]::Pow(2, $attempt))
        if ($rh) { foreach ($k in @($rh.Keys)) { if ($k -match 'retry-after') { $v = 0; if ([int]::TryParse((@($rh[$k])[0]), [ref]$v)) { $wait = [Math]::Max($wait, [Math]::Min($v, 120)) } } } }
        Start-Sleep -Seconds $wait; continue
    }
    $msg = if ($resp -and $resp.error) { "$($resp.error.code): $($resp.error.message)" } elseif ($resp -and $resp.code) { "$($resp.code): $($resp.message)" } elseif ($netErr) { $netErr } else { "HTTP $sc" }
    return [pscustomobject]@{ Key = $Item.Key; Ok = $false; Status = $sc; Data = $null; Error = $msg }
}
'@

function Invoke-ParallelRest {
    <# Items: @{ Key; Url; Method='GET'; Body=$null(json string); TokenKey='arm'|'la' } #>
    param([object[]]$Items, [string]$Label = 'requests')
    $results = @{}
    if (-not $Items -or $Items.Count -eq 0) { return $results }
    $workerText = $script:RestWorkerText
    $batch = 400
    for ($b = 0; $b -lt $Items.Count; $b += $batch) {
        $slice = @($Items[$b..([Math]::Min($b + $batch - 1, $Items.Count - 1))])
        $ctx = @{ Tokens = @{ arm = (Get-BearerToken -Resource $script:ArmResource) } }
        if ($slice | Where-Object { $_.TokenKey -eq 'la' }) { $ctx.Tokens.la = Get-BearerToken -Resource $script:LaResource }
        $out = $slice | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $it = $_
            try { $sb = [scriptblock]::Create($using:workerText); & $sb $it $using:ctx }
            catch { [pscustomobject]@{ Key = $it.Key; Ok = $false; Status = 0; Data = $null; Error = "Worker failure: $($_.Exception.Message)" } }
        }
        foreach ($o in $out) { if ($o) { $results[$o.Key] = $o } }
        Write-Log "    $Label : $([Math]::Min($b + $batch, $Items.Count))/$($Items.Count) done" Info
    }
    return $results
}

#endregion

#region ============================== CATALOG HELPERS ==============================

function M {
    <# Metric spec. Out = stats exported as columns (Avg,P95,Max,Min,P05,Med,Total,MaxT).
       Basis = series used for Avg/P95 (Avg = hourly average, Max = hourly peak, Total = hourly total). #>
    param([string]$Name, [string]$Alias, [double]$Scale = 1, [string]$Out = 'Avg,P95,Max', [switch]$Pattern,
        [double]$Idle = -1, [string]$SubPath = '', [string]$Filter = '', [string]$PatternAgg = 'Avg', [string]$Basis = 'Avg')
    @{ Name = $Name; Alias = $Alias; Scale = $Scale; Out = $Out; Pattern = [bool]$Pattern; Idle = $Idle
        SubPath = $SubPath; Filter = $Filter; PatternAgg = $PatternAgg; Basis = $Basis }
}

function Get-Count($x) { return @($x | Where-Object { $null -ne $_ }).Count }

function Get-Ratio($a, $b, [int]$Digits = 1) {
    if ($null -eq $a -or $null -eq $b) { return $null }
    try { $bd = [double]$b; if ($bd -eq 0) { return $null }; return [Math]::Round(100.0 * [double]$a / $bd, $Digits) } catch { return $null }
}

function Get-UsedPct($Available, $Total) {
    if ($null -eq $Available -or -not $Total) { return $null }
    return [Math]::Round([Math]::Max(0.0, [Math]::Min(100.0, (1 - [double]$Available / [double]$Total) * 100)), 1)
}

function Get-SumNonNull {
    $vals = @($args | Where-Object { $null -ne $_ })
    if ($vals.Count -eq 0) { return $null }
    return [Math]::Round([double](($vals | Measure-Object -Sum).Sum), 3)
}

function ConvertTo-GiB([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return 0 }
    if ($s -match '^\s*([\d\.]+)\s*(Gi|G)\s*$') { return [double]$Matches[1] }
    if ($s -match '^\s*([\d\.]+)\s*(Mi|M)\s*$') { return [double]$Matches[1] / 1024 }
    if ($s -match '^\s*([\d\.]+)\s*$') { return [double]$Matches[1] }
    return 0
}

function Get-AgeDays([string]$When) {
    if ([string]::IsNullOrWhiteSpace($When)) { return $null }
    try { return [int]([DateTime]::UtcNow - ([DateTime]::Parse($When, [Globalization.CultureInfo]::InvariantCulture)).ToUniversalTime()).TotalDays } catch { return $null }
}

$script:VmSkuCache = @{}
$script:VmSkuRows = [System.Collections.Generic.List[object]]::new()
function Get-VmSkuInfo {
    <# vCPU / memory / max data disks for a VM size in a region (Microsoft.Compute/skus, cached per region). #>
    param([string]$Size, [string]$Location, [string]$SubscriptionId)
    $empty = @{ vCPUs = $null; MemoryGB = $null; MaxDataDisks = $null; Family = $null; PremiumIO = $null }
    if ([string]::IsNullOrWhiteSpace($Size) -or [string]::IsNullOrWhiteSpace($Location)) { return $empty }
    $loc = $Location.ToLower()
    if (-not $script:VmSkuCache.ContainsKey($loc)) {
        $map = @{}
        try {
            $f = [uri]::EscapeDataString("location eq '$loc'")
            $resp = Invoke-AzApi -Uri "$script:ArmBase/subscriptions/$SubscriptionId/providers/Microsoft.Compute/skus?api-version=2021-07-01&`$filter=$f"
            foreach ($s in @($resp.value)) {
                if ($s.resourceType -ne 'virtualMachines') { continue }
                $caps = @{}; foreach ($c in @($s.capabilities)) { $caps[[string]$c.name] = [string]$c.value }
                $restricted = [bool](@($s.restrictions | Where-Object { $_.type -eq 'Location' }).Count)
                $info = @{
                    vCPUs = if ($caps.vCPUs) { [int]$caps.vCPUs } else { $null }
                    MemoryGB = if ($caps.MemoryGB) { [double]$caps.MemoryGB } else { $null }
                    MaxDataDisks = if ($caps.MaxDataDiskCount) { [int]$caps.MaxDataDiskCount } else { $null }
                    PremiumIO = $caps.PremiumIO; Family = [string]$s.family
                }
                $map[[string]$s.name] = $info
                $script:VmSkuRows.Add([pscustomobject]@{
                        Location = $loc; VmSize = [string]$s.name; Family = [string]$s.family; vCPUs = $info.vCPUs; MemoryGB = $info.MemoryGB
                        MaxDataDisks = $info.MaxDataDisks; PremiumIO = $caps.PremiumIO; ACUs = $caps.ACUs; CpuArchitecture = $caps.CpuArchitectureType
                        HyperVGenerations = $caps.HyperVGenerations; MaxNICs = $caps.MaxNetworkInterfaces; RestrictedForSubscription = $restricted
                    })
            }
            Write-Log "    VM size catalog loaded for $loc ($($map.Count) sizes)" Info
        }
        catch { Add-CollectionError 'VmSkuCatalog' $loc $_.Exception.Message; Write-Log "    Could not load VM size catalog for $loc : $($_.Exception.Message)" Warning }
        $script:VmSkuCache[$loc] = $map
    }
    $hit = $script:VmSkuCache[$loc][$Size]
    if ($hit) { return $hit } else { return $empty }
}

function Get-FlexServerProps($r) {
    $sku = [string]$r.sku.name
    $vc = $null; if ($sku -match '^Standard_[A-Za-z]+?(\d+)') { $vc = [int]$Matches[1] }
    [ordered]@{
        ServerSku = $sku; Tier = [string]$r.sku.tier; vCores = $vc
        StorageGB = $r.properties.storage.storageSizeGB; StorageIops = $r.properties.storage.iops
        StorageAutoGrow = [string]$r.properties.storage.autoGrow; StorageTier = [string]$r.properties.storage.tier
        HAMode = [string]$r.properties.highAvailability.mode; State = [string]$r.properties.state; Version = [string]$r.properties.version
        BackupRetentionDays = $r.properties.backup.backupRetentionDays
        ReplicationRole = [string]$r.properties.replicationRole
    }
}

#endregion

#region ============================== RESOURCE CATALOG ==============================
# Each entry: Key, Display, Category, Types (ARG types), Where (extra KQL), Props {param($r)}, Expand {param($r,$base)},
#             Metrics {param($r,$row)} -> metric specs, Derive {param($row,$series)}, PreCollect {param($resources)}

$script:Catalog = [System.Collections.Generic.List[hashtable]]::new()

# ---------------------------------------------------------------- COMPUTE
$script:Catalog.Add(@{
        Key = 'VirtualMachines'; Display = 'Virtual Machines'; Category = 'Compute'; Types = @('microsoft.compute/virtualmachines')
        Props = { param($r)
            $size = [string]$r.properties.hardwareProfile.vmSize
            $sku = Get-VmSkuInfo $size $r.location $r.subscriptionId
            $img = $r.properties.storageProfile.imageReference
            $dd = @($r.properties.storageProfile.dataDisks | Where-Object { $_ })
            [ordered]@{
                VmSize = $size; vCPUs = $sku.vCPUs; MemoryGB = $sku.MemoryGB; VmFamily = $sku.Family
                PowerState = ([string]$r.properties.extended.instanceView.powerState.code) -replace '^PowerState/', ''
                OsType = [string]$r.properties.storageProfile.osDisk.osType
                OsName = [string]$r.properties.extended.instanceView.osName
                LicenseType = [string]$r.properties.licenseType; Priority = [string]$r.properties.priority
                ImagePublisher = [string]$img.publisher; ImageOffer = [string]$img.offer; ImageSku = [string]$img.sku
                OsDiskType = [string]$r.properties.storageProfile.osDisk.managedDisk.storageAccountType
                OsDiskSizeGB = $r.properties.storageProfile.osDisk.diskSizeGB
                DataDiskCount = $dd.Count
                DataDiskTotalGB = [double](($dd | Measure-Object -Property diskSizeGB -Sum).Sum)
                DataDiskTypes = (@($dd | ForEach-Object { $_.managedDisk.storageAccountType } | Where-Object { $_ } | Select-Object -Unique) -join ',')
                MaxDataDisks = $sku.MaxDataDisks
                ScaleSet = Get-NameFromId ([string]$r.properties.virtualMachineScaleSet.id)
                ManagedByService = Get-ManagedByService $r; TimeCreated = [string]$r.properties.timeCreated
            } }
        Metrics = { param($r, $row)
            $m = [System.Collections.Generic.List[object]]::new()
            $m.Add((M 'Percentage CPU' 'CPU' -Out 'Avg,P95,Max,Med,P99' -Pattern -Idle 5))
            $m.Add((M 'Available Memory Percentage' 'AvailMemPct' -Out 'Avg,MinAvg,P05'))
            $m.Add((M 'Available Memory Bytes' 'AvailMemGB' -Scale 1GB -Out 'Avg,MinAvg,P05,Min'))
            $m.Add((M 'Network In Total' 'NetInGB' -Scale 1GB -Out 'Total,MaxT' -Basis 'Total'))
            $m.Add((M 'Network Out Total' 'NetOutGB' -Scale 1GB -Out 'Total,MaxT' -Basis 'Total'))
            $m.Add((M 'Disk Read Operations/Sec' 'DiskReadIOPS' -Out 'Avg,P95,Max'))
            $m.Add((M 'Disk Write Operations/Sec' 'DiskWriteIOPS' -Out 'Avg,P95,Max'))
            $m.Add((M 'Disk Read Bytes' 'DiskReadGB' -Scale 1GB -Out 'Total'))
            $m.Add((M 'Disk Write Bytes' 'DiskWriteGB' -Scale 1GB -Out 'Total'))
            $m.Add((M 'VM Uncached IOPS Consumed Percentage' 'VmUncachedIopsPct' -Out 'Avg,P95,Max'))
            $m.Add((M 'VM Uncached Bandwidth Consumed Percentage' 'VmUncachedBwPct' -Out 'Avg,P95,Max'))
            $m.Add((M 'VM Cached IOPS Consumed Percentage' 'VmCachedIopsPct' -Out 'Avg,P95,Max'))
            $m.Add((M 'Data Disk IOPS Consumed Percentage' 'DataDiskIopsPct' -Out 'P95,Max' -Basis 'Max'))
            $m.Add((M 'OS Disk IOPS Consumed Percentage' 'OsDiskIopsPct' -Out 'P95,Max' -Basis 'Max'))
            $m.Add((M 'Inbound Flows' 'InboundFlows' -Out 'Avg,Max'))
            $m.Add((M 'Outbound Flows' 'OutboundFlows' -Out 'Avg,Max'))
            if ([string]$row.VmSize -match '^(?i)Standard_B') { $m.Add((M 'CPU Credits Remaining' 'CpuCreditsRemaining' -Out 'Avg,MinAvg')); $m.Add((M 'CPU Credits Consumed' 'CpuCreditsConsumed' -Out 'Avg,Max')) }
            $m.ToArray() }
        Derive = { param($row, $series)
            if ($null -ne $row.AvailMemPct_Avg) {
                $row.MemUsedPct_Avg = Get-Round (100 - [double]$row.AvailMemPct_Avg) 1
                $row.MemUsedPct_P95 = if ($null -ne $row.AvailMemPct_P05) { Get-Round (100 - [double]$row.AvailMemPct_P05) 1 } else { $null }
                $row.MemUsedPct_Max = if ($null -ne $row.AvailMemPct_MinAvg) { Get-Round (100 - [double]$row.AvailMemPct_MinAvg) 1 } else { $null }
            }
            else {
                $row.MemUsedPct_Avg = Get-UsedPct $row.AvailMemGB_Avg $row.MemoryGB
                $row.MemUsedPct_P95 = Get-UsedPct $row.AvailMemGB_P05 $row.MemoryGB
                $row.MemUsedPct_Max = Get-UsedPct $row.AvailMemGB_MinAvg $row.MemoryGB
            }
            $row.CpuUsedvCPUs_P95 = if ($null -ne $row.CPU_P95 -and $row.vCPUs) { Get-Round ([double]$row.CPU_P95 * [double]$row.vCPUs / 100) 2 } else { $null }
            $row.MemUsedGB_Max = if ($null -ne $row.MemUsedPct_Max -and $row.MemoryGB) { Get-Round ([double]$row.MemUsedPct_Max * [double]$row.MemoryGB / 100) 1 } else { $null }
            $row.NetInMbps_PeakHour = if ($null -ne $row.NetInGB_MaxT) { Get-Round ([double]$row.NetInGB_MaxT * 8192 / 3600) 2 } else { $null }
            $row.NetOutMbps_PeakHour = if ($null -ne $row.NetOutGB_MaxT) { Get-Round ([double]$row.NetOutGB_MaxT * 8192 / 3600) 2 } else { $null }
            $row.TotalIOPS_P95 = Get-SumNonNull $row.DiskReadIOPS_P95 $row.DiskWriteIOPS_P95 }
    })

$script:Catalog.Add(@{
        Key = 'VMScaleSets'; Display = 'VM Scale Sets'; Category = 'Compute'; Types = @('microsoft.compute/virtualmachinescalesets')
        Props = { param($r)
            $size = [string]$r.sku.name; $sku = Get-VmSkuInfo $size $r.location $r.subscriptionId
            $cap = [int]($r.sku.capacity ?? 0)
            $as = $script:Autoscale[([string]$r.id).ToLower()]
            [ordered]@{
                VmSize = $size; Instances = $cap; vCPUsPerInstance = $sku.vCPUs; MemoryGBPerInstance = $sku.MemoryGB
                TotalvCPUs = if ($sku.vCPUs) { $sku.vCPUs * $cap } else { $null }; VmFamily = $sku.Family
                OrchestrationMode = [string]$r.properties.orchestrationMode
                OsType = [string]$r.properties.virtualMachineProfile.storageProfile.osDisk.osType
                Priority = [string]$r.properties.virtualMachineProfile.priority
                LicenseType = [string]$r.properties.virtualMachineProfile.licenseType
                UpgradePolicy = [string]$r.properties.upgradePolicy.mode
                AutoscaleEnabled = $as.Enabled; AutoscaleMin = $as.Min; AutoscaleMax = $as.Max
                ManagedByService = Get-ManagedByService $r; TimeCreated = [string]$r.properties.timeCreated
            } }
        Metrics = { param($r, $row)
            if ($row.ManagedByService -eq 'AKS' -or $row.Instances -eq 0) { return @() }
            @((M 'Percentage CPU' 'CPU' -Out 'Avg,P95,Max,Med,P99' -Pattern -Idle 5),
                (M 'Available Memory Percentage' 'AvailMemPct' -Out 'Avg,MinAvg,P05'),
                (M 'Available Memory Bytes' 'AvailMemGB' -Scale 1GB -Out 'Avg,MinAvg,P05,Min'),
                (M 'Network In Total' 'NetInGB' -Scale 1GB -Out 'Total,MaxT' -Basis 'Total'),
                (M 'Network Out Total' 'NetOutGB' -Scale 1GB -Out 'Total,MaxT' -Basis 'Total'),
                (M 'VM Uncached IOPS Consumed Percentage' 'VmUncachedIopsPct' -Out 'Avg,P95,Max')) }
        Derive = { param($row, $series)
            if ($null -ne $row.AvailMemPct_Avg) {
                $row.MemUsedPct_Avg = Get-Round (100 - [double]$row.AvailMemPct_Avg) 1
                $row.MemUsedPct_P95 = if ($null -ne $row.AvailMemPct_P05) { Get-Round (100 - [double]$row.AvailMemPct_P05) 1 } else { $null }
                $row.MemUsedPct_Max = if ($null -ne $row.AvailMemPct_MinAvg) { Get-Round (100 - [double]$row.AvailMemPct_MinAvg) 1 } else { $null }
            }
            else {
                $row.MemUsedPct_Avg = Get-UsedPct $row.AvailMemGB_Avg $row.MemoryGBPerInstance
                $row.MemUsedPct_P95 = Get-UsedPct $row.AvailMemGB_P05 $row.MemoryGBPerInstance
                $row.MemUsedPct_Max = Get-UsedPct $row.AvailMemGB_MinAvg $row.MemoryGBPerInstance
            } }
    })

$script:Catalog.Add(@{
        Key = 'AKSNodePools'; Display = 'AKS Clusters (Node Pools)'; Category = 'Containers'; Types = @('microsoft.containerservice/managedclusters')
        Props = { param($r) [ordered]@{
                KubernetesVersion = [string]$r.properties.kubernetesVersion; ClusterTier = [string]$r.sku.tier
                ClusterPowerState = [string]$r.properties.powerState.code; NodeResourceGroup = [string]$r.properties.nodeResourceGroup
                NetworkPlugin = [string]$r.properties.networkProfile.networkPlugin
                CostAnalysisEnabled = $r.properties.metricsProfile.costAnalysis.enabled
            } }
        Expand = { param($r, $base)
            foreach ($p in @($r.properties.agentPoolProfiles | Where-Object { $_ })) {
                $sku = Get-VmSkuInfo ([string]$p.vmSize) $r.location $r.subscriptionId
                $row = [ordered]@{}; foreach ($k in $base.Keys) { $row[$k] = $base[$k] }
                $cnt = [int]($p.count ?? 0)
                $row.NodePool = [string]$p.name; $row.PoolMode = [string]$p.mode; $row.VmSize = [string]$p.vmSize
                $row.vCPUsPerNode = $sku.vCPUs; $row.MemoryGBPerNode = $sku.MemoryGB; $row.NodeCount = $cnt
                $row.AutoscaleEnabled = [bool]$p.enableAutoScaling; $row.MinCount = $p.minCount; $row.MaxCount = $p.maxCount
                $row.TotalvCPUs = if ($sku.vCPUs) { $sku.vCPUs * $cnt } else { $null }
                $row.TotalMemoryGB = if ($sku.MemoryGB) { $sku.MemoryGB * $cnt } else { $null }
                $row.PoolPowerState = [string]$p.powerState.code; $row.Priority = [string]$p.scaleSetPriority
                $row.OsType = [string]$p.osType; $row.OsSku = [string]$p.osSKU; $row.PoolVersion = [string]$p.orchestratorVersion
                $row._Dim = [string]$p.name
                $row
            } }
        Metrics = { param($r, $row)
            if ($row.ClusterPowerState -eq 'Stopped') { return @() }
            @((M 'node_cpu_usage_percentage' 'NodeCPU' -Out 'Avg,P95,Max,P99' -Filter "nodepool eq '*'" -Pattern -Idle 5),
                (M 'node_memory_working_set_percentage' 'NodeMemory' -Filter "nodepool eq '*'"),
                (M 'node_memory_rss_percentage' 'NodeMemoryRss' -Out 'Avg,P95,Max' -Filter "nodepool eq '*'"),
                (M 'node_disk_usage_percentage' 'NodeDisk' -Out 'P95,Max' -Filter "nodepool eq '*'")) }
    })

$script:Catalog.Add(@{
        Key = 'ContainerApps'; Display = 'Container Apps'; Category = 'Containers'; Types = @('microsoft.app/containerapps')
        Props = { param($r)
            $cs = @($r.properties.template.containers | Where-Object { $_ })
            [ordered]@{
                Environment = Get-NameFromId ([string]($r.properties.managedEnvironmentId ?? $r.properties.environmentId))
                WorkloadProfile = [string]$r.properties.workloadProfileName
                MinReplicas = $r.properties.template.scale.minReplicas; MaxReplicas = $r.properties.template.scale.maxReplicas
                CpuCoresPerReplica = [double](($cs | ForEach-Object { [double]($_.resources.cpu ?? 0) } | Measure-Object -Sum).Sum)
                MemoryGiBPerReplica = [double](($cs | ForEach-Object { ConvertTo-GiB ([string]$_.resources.memory) } | Measure-Object -Sum).Sum)
                ContainerCount = $cs.Count; IngressEnabled = [bool]$r.properties.configuration.ingress
                RunningStatus = [string]$r.properties.runningStatus
            } }
        Metrics = { param($r, $row) @(
                (M 'UsageNanoCores' 'CpuCores' -Scale 1e9 -Pattern -Idle 0.01),
                (M 'WorkingSetBytes' 'MemoryGB' -Scale 1GB),
                (M 'CpuPercentage' 'CpuPct' -Out 'Avg,P95,Max'),
                (M 'MemoryPercentage' 'MemPct' -Out 'Avg,P95,Max'),
                (M 'Requests' 'Requests' -Out 'Total,MaxT' -Basis 'Total'),
                (M 'RestartCount' 'Restarts' -Out 'Total'),
                (M 'Replicas' 'Replicas' -Out 'Avg,Max')) }
        Derive = { param($row, $series)
            $row.CpuUtilPct_P95 = if ($null -ne $row.CpuPct_P95) { $row.CpuPct_P95 } else { Get-Ratio $row.CpuCores_P95 $row.CpuCoresPerReplica }
            $row.MemUtilPct_P95 = if ($null -ne $row.MemPct_P95) { $row.MemPct_P95 } else { Get-Ratio $row.MemoryGB_P95 $row.MemoryGiBPerReplica } }
    })

$script:Catalog.Add(@{
        Key = 'ContainerAppEnvironments'; Display = 'Container App Environments'; Category = 'Containers'; Types = @('microsoft.app/managedenvironments')
        Props = { param($r)
            $wps = @($r.properties.workloadProfiles | Where-Object { $_ })
            $ded = @($wps | Where-Object { $_.workloadProfileType -and $_.workloadProfileType -ne 'Consumption' })
            [ordered]@{
                WorkloadProfiles = (@($wps | ForEach-Object { "$($_.name):$($_.workloadProfileType)[$($_.minimumCount)-$($_.maximumCount)]" }) -join '; ')
                DedicatedProfiles = $ded.Count
                DedicatedMinNodes = [int](($ded | ForEach-Object { [int]($_.minimumCount ?? 0) } | Measure-Object -Sum).Sum)
                InfrastructureRG = [string]$r.properties.infrastructureResourceGroup
            } }
    })

$script:Catalog.Add(@{
        Key = 'ContainerInstances'; Display = 'Container Instances'; Category = 'Containers'; Types = @('microsoft.containerinstance/containergroups')
        Props = { param($r)
            $cs = @($r.properties.containers | Where-Object { $_ })
            [ordered]@{
                OsType = [string]$r.properties.osType; State = [string]$r.properties.instanceView.state
                RestartPolicy = [string]$r.properties.restartPolicy; ContainerCount = $cs.Count
                CpuCores = [double](($cs | ForEach-Object { [double]($_.properties.resources.requests.cpu ?? 0) } | Measure-Object -Sum).Sum)
                MemoryGB = [double](($cs | ForEach-Object { [double]($_.properties.resources.requests.memoryInGB ?? 0) } | Measure-Object -Sum).Sum)
                Sku = [string]$r.properties.sku
            } }
        Metrics = { param($r, $row) @((M 'CpuUsage' 'CpuMillicores' -Pattern -Idle 5), (M 'MemoryUsage' 'MemoryMB' -Scale 1MB)) }
        Derive = { param($row, $series)
            $row.CpuUtilPct_P95 = Get-Ratio $row.CpuMillicores_P95 (1000 * [double]($row.CpuCores ?? 0))
            $row.MemUtilPct_P95 = Get-Ratio $row.MemoryMB_P95 (1024 * [double]($row.MemoryGB ?? 0)) }
    })

$script:Catalog.Add(@{
        Key = 'ContainerRegistries'; Display = 'Container Registries'; Category = 'Containers'; Types = @('microsoft.containerregistry/registries')
        Props = { param($r)
            $inc = switch ([string]$r.sku.name) { 'Basic' { 10 } 'Standard' { 100 } 'Premium' { 500 } default { $null } }
            [ordered]@{
                RegistrySku = [string]$r.sku.name; IncludedStorageGB = $inc
                GeoReplications = [int]($script:AcrReplications[([string]$r.id).ToLower()] ?? 0)
                RetentionPolicy = [string]$r.properties.policies.retentionPolicy.status
            } }
        Metrics = { param($r, $row) @(
                (M 'StorageUsed' 'StorageUsedGB' -Scale 1GB -Out 'Avg,Max'),
                (M 'TotalPullCount' 'Pulls' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'TotalPushCount' 'Pushes' -Out 'Total')) }
    })

# ---------------------------------------------------------------- APP SERVICE
$script:Catalog.Add(@{
        Key = 'AppServicePlans'; Display = 'App Service Plans'; Category = 'App Service'; Types = @('microsoft.web/serverfarms')
        Props = { param($r)
            $as = $script:Autoscale[([string]$r.id).ToLower()]
            [ordered]@{
                PlanSku = [string]$r.sku.name; PlanTier = [string]$r.sku.tier; Instances = $r.sku.capacity
                OsType = if ($r.properties.reserved -eq $true -or [string]$r.kind -match 'linux') { 'Linux' } else { 'Windows' }
                NumberOfSites = $r.properties.numberOfSites; MaxElasticWorkers = $r.properties.maximumElasticWorkerCount
                ElasticScaleEnabled = $r.properties.elasticScaleEnabled; PerSiteScaling = $r.properties.perSiteScaling
                Status = [string]$r.properties.status
                AppServiceEnvironment = Get-NameFromId ([string]$r.properties.hostingEnvironmentProfile.id)
                AutoscaleEnabled = $as.Enabled; AutoscaleMin = $as.Min; AutoscaleMax = $as.Max
            } }
        Metrics = { param($r, $row)
            if ($row.PlanTier -in @('Free', 'Shared', 'Dynamic', 'FlexConsumption') -or $row.PlanSku -in @('F1', 'D1', 'Y1', 'FC1')) { return @() }
            @((M 'CpuPercentage' 'CPU' -Pattern -Idle 5),
                (M 'MemoryPercentage' 'Memory'),
                (M 'HttpQueueLength' 'HttpQueue' -Out 'Avg,Max'),
                (M 'BytesReceived' 'DataInGB' -Scale 1GB -Out 'Total'),
                (M 'BytesSent' 'DataOutGB' -Scale 1GB -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'AppServices'; Display = 'Web & Function Apps'; Category = 'App Service'; Types = @('microsoft.web/sites')
        Props = { param($r)
            $kind = [string]$r.kind
            $appType = if ($kind -match 'workflowapp') { 'Logic App (Standard)' } elseif ($kind -match 'functionapp') { 'Function App' } else { 'Web App' }
            [ordered]@{
                AppType = $appType; State = [string]$r.properties.state
                AppServicePlan = Get-NameFromId ([string]$r.properties.serverFarmId); PlanSkuFamily = [string]$r.properties.sku
                DefaultHostName = [string]$r.properties.defaultHostName; HttpsOnly = $r.properties.httpsOnly
                LastModifiedUtc = [string]$r.properties.lastModifiedTimeUtc; UsageState = [string]$r.properties.usageState
                AppServicePlanId = [string]$r.properties.serverFarmId
            } }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Stopped') { return @() }
            $m = @((M 'Requests' 'Requests' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'Http5xx' 'Http5xx' -Out 'Total'),
                (M 'Http4xx' 'Http4xx' -Out 'Total'),
                (M 'HttpResponseTime' 'ResponseTimeSec' -Out 'Avg,P95,Max'),
                (M 'MemoryWorkingSet' 'MemoryMB' -Scale 1MB -Out 'Avg,Max'))
            if ($row.AppType -eq 'Function App') { $m += @((M 'FunctionExecutionCount' 'Executions' -Out 'Total'), (M 'FunctionExecutionUnits' 'ExecutionUnits' -Out 'Total')) }
            else { $m += (M 'CpuTime' 'CpuTimeSec' -Out 'Total') }
            $m }
    })

# ---------------------------------------------------------------- DATABASES
$script:Catalog.Add(@{
        Key = 'SqlDatabases'; Display = 'Azure SQL Databases'; Category = 'Databases'; Types = @('microsoft.sql/servers/databases')
        Where = "| where name !~ 'master'"
        Props = { param($r)
            $tier = [string]$r.sku.tier; $skuName = [string]$r.sku.name
            $model = if ($tier -in @('Basic', 'Standard', 'Premium')) { 'DTU' } elseif ($tier -eq 'DataWarehouse') { 'DWU' } else { 'vCore' }
            [ordered]@{
                ServerName = ($r.id -split '/')[8]; Edition = $tier; ServiceObjective = $skuName; Capacity = $r.sku.capacity
                Family = [string]$r.sku.family; PurchaseModel = $model
                ComputeTier = if ($skuName -match '_S_') { 'Serverless' } else { 'Provisioned' }
                ElasticPool = Get-NameFromId ([string]$r.properties.elasticPoolId)
                MaxSizeGB = if ($r.properties.maxSizeBytes) { [Math]::Round([double]$r.properties.maxSizeBytes / 1GB, 1) } else { $null }
                LicenseType = [string]$r.properties.licenseType
                ReplicaCount = $r.properties.highAvailabilityReplicaCount
                AutoPauseDelayMin = $r.properties.autoPauseDelay; MinCapacity = $r.properties.minCapacity
                Status = [string]$r.properties.status
                SecondaryType = [string]$r.properties.secondaryType; CreationDate = [string]$r.properties.creationDate
                ElasticPoolId = [string]$r.properties.elasticPoolId
            } }
        Metrics = { param($r, $row)
            if ($row.PurchaseModel -eq 'DWU') { return @((M 'dwu_consumption_percent' 'DWU' -Pattern -Idle 2), (M 'cpu_percent' 'CPU')) }
            $m = [System.Collections.Generic.List[object]]::new()
            if ($row.ElasticPool) { $m.Add((M 'cpu_percent' 'CPU' -Pattern -Idle 2)) }
            elseif ($row.PurchaseModel -eq 'DTU') {
                $m.Add((M 'dtu_consumption_percent' 'DTU' -Pattern -Idle 2)); $m.Add((M 'cpu_percent' 'CPU'))
                $m.Add((M 'physical_data_read_percent' 'DataIO' -Out 'Avg,Max')); $m.Add((M 'log_write_percent' 'LogWrite' -Out 'Avg,Max'))
            }
            else {
                $m.Add((M 'cpu_percent' 'CPU' -Pattern -Idle 2)); $m.Add((M 'sql_instance_memory_percent' 'Memory'))
                $m.Add((M 'workers_percent' 'Workers' -Out 'Avg,Max')); $m.Add((M 'physical_data_read_percent' 'DataIO' -Out 'Avg,Max'))
                $m.Add((M 'log_write_percent' 'LogWrite' -Out 'Avg,Max'))
                if ($row.ComputeTier -eq 'Serverless') { $m.Add((M 'app_cpu_billed' 'BilledvCoreSec' -Out 'Total')) }
            }
            $m.Add((M 'connection_successful' 'Connections' -Out 'Total' -Basis 'Total'))
            $m.Add((M 'connection_failed' 'ConnectionsFailed' -Out 'Total'))
            $m.Add((M 'sessions_percent' 'Sessions' -Out 'P95,Max'))
            $m.Add((M 'deadlock' 'Deadlocks' -Out 'Total'))
            if ($row.Edition -ne 'Hyperscale') { $m.Add((M 'storage' 'StorageUsedGB' -Scale 1GB -Out 'Max' -Basis 'Max')) }
            $m.Add((M 'allocated_data_storage' 'AllocatedGB' -Scale 1GB -Out 'Max' -Basis 'Max'))
            $m.ToArray() }
        Derive = { param($row, $series) $row.StorageUsedPct = Get-Ratio $row.StorageUsedGB_Max $row.MaxSizeGB }
    })

$script:Catalog.Add(@{
        Key = 'SqlElasticPools'; Display = 'SQL Elastic Pools'; Category = 'Databases'; Types = @('microsoft.sql/servers/elasticpools')
        Props = { param($r)
            $tier = [string]$r.sku.tier
            [ordered]@{
                ServerName = ($r.id -split '/')[8]; Edition = $tier; PoolSku = [string]$r.sku.name; Capacity = $r.sku.capacity
                Family = [string]$r.sku.family; PurchaseModel = if ($tier -in @('Basic', 'Standard', 'Premium')) { 'DTU' } else { 'vCore' }
                MaxSizeGB = if ($r.properties.maxSizeBytes) { [Math]::Round([double]$r.properties.maxSizeBytes / 1GB, 1) } else { $null }
                PerDbMin = $r.properties.perDatabaseSettings.minCapacity; PerDbMax = $r.properties.perDatabaseSettings.maxCapacity
                LicenseType = [string]$r.properties.licenseType
                ReplicaCount = $r.properties.highAvailabilityReplicaCount
                DatabaseCount = [int]($script:PoolDbCount[([string]$r.id).ToLower()] ?? 0)
            } }
        Metrics = { param($r, $row)
            $m = if ($row.PurchaseModel -eq 'DTU') { @((M 'dtu_consumption_percent' 'DTU' -Pattern -Idle 2), (M 'cpu_percent' 'CPU')) }
            else { @((M 'cpu_percent' 'CPU' -Pattern -Idle 2), (M 'workers_percent' 'Workers' -Out 'Avg,Max'), (M 'sql_instance_memory_percent' 'Memory')) }
            $m + @((M 'storage_used' 'StorageUsedGB' -Scale 1GB -Out 'Max' -Basis 'Max'),
                (M 'allocated_data_storage' 'AllocatedGB' -Scale 1GB -Out 'Max' -Basis 'Max'),
                (M 'storage_percent' 'StoragePct' -Out 'Max' -Basis 'Max'),
                (M 'sessions_percent' 'Sessions' -Out 'P95,Max'),
                (M 'log_write_percent' 'LogWrite' -Out 'Avg,Max'),
                (M 'physical_data_read_percent' 'DataIO' -Out 'Avg,Max')) }
    })

$script:Catalog.Add(@{
        Key = 'SqlManagedInstances'; Display = 'SQL Managed Instances'; Category = 'Databases'; Types = @('microsoft.sql/managedinstances')
        Props = { param($r) [ordered]@{
                MiSku = [string]$r.sku.name; Tier = [string]$r.sku.tier; vCores = $r.properties.vCores
                StorageSizeGB = $r.properties.storageSizeInGB; LicenseType = [string]$r.properties.licenseType
                HardwareFamily = [string]$r.sku.family; State = [string]$r.properties.state
                InstancePool = Get-NameFromId ([string]$r.properties.instancePoolId); HybridSecondaryUsage = [string]$r.properties.hybridSecondaryUsage
            } }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Stopped') { return @() }
            @((M 'avg_cpu_percent' 'CPU' -Pattern -Idle 5),
                (M 'storage_space_used_mb' 'StorageUsedGB' -Scale 1024 -Out 'Max' -Basis 'Max'),
                (M 'reserved_storage_mb' 'ReservedStorageGB' -Scale 1024 -Out 'Max' -Basis 'Max'),
                (M 'io_requests' 'IORequests' -Out 'Total')) }
        Derive = { param($row, $series) $row.StorageUsedPct = Get-Ratio $row.StorageUsedGB_Max ($row.ReservedStorageGB_Max ?? $row.StorageSizeGB) }
    })

$script:Catalog.Add(@{
        Key = 'SqlVirtualMachines'; Display = 'SQL Server on Azure VMs'; Category = 'Databases'; Types = @('microsoft.sqlvirtualmachine/sqlvirtualmachines')
        Props = { param($r) [ordered]@{
                SqlImageOffer = [string]$r.properties.sqlImageOffer; SqlEdition = [string]$r.properties.sqlImageSku
                SqlLicenseType = [string]$r.properties.sqlServerLicenseType; SqlManagement = [string]$r.properties.sqlManagement
                VirtualMachine = Get-NameFromId ([string]$r.properties.virtualMachineResourceId)
                VirtualMachineId = [string]$r.properties.virtualMachineResourceId
            } }
    })

$script:Catalog.Add(@{
        Key = 'MySQLFlexibleServers'; Display = 'MySQL Flexible Servers'; Category = 'Databases'; Types = @('microsoft.dbformysql/flexibleservers')
        Props = { param($r) Get-FlexServerProps $r }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Stopped') { return @() }
            @((M 'cpu_percent' 'CPU' -Pattern -Idle 5), (M 'memory_percent' 'Memory'),
                (M 'storage_percent' 'StoragePct' -Out 'Max' -Basis 'Max'), (M 'storage_used' 'StorageUsedGB' -Scale 1GB -Out 'Max' -Basis 'Max'),
                (M 'active_connections' 'Connections' -Out 'Avg,Max'), (M 'io_consumption_percent' 'IOPct' -Out 'Avg,Max')) }
    })

$script:Catalog.Add(@{
        Key = 'PostgreSQLFlexibleServers'; Display = 'PostgreSQL Flexible Servers'; Category = 'Databases'; Types = @('microsoft.dbforpostgresql/flexibleservers')
        Props = { param($r) Get-FlexServerProps $r }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Stopped') { return @() }
            @((M 'cpu_percent' 'CPU' -Pattern -Idle 5), (M 'memory_percent' 'Memory'),
                (M 'storage_percent' 'StoragePct' -Out 'Max' -Basis 'Max'), (M 'storage_used' 'StorageUsedGB' -Scale 1GB -Out 'Max' -Basis 'Max'),
                (M 'active_connections' 'Connections' -Out 'Avg,Max'), (M 'disk_iops_consumed_percentage' 'IOPct' -Out 'Avg,Max')) }
    })

$script:Catalog.Add(@{
        Key = 'LegacyDatabaseServers'; Display = 'MySQL/PostgreSQL/MariaDB Single Servers (retired)'; Category = 'Databases'
        Types = @('microsoft.dbformysql/servers', 'microsoft.dbforpostgresql/servers', 'microsoft.dbformariadb/servers')
        Props = { param($r) [ordered]@{
                ServerSku = [string]$r.sku.name; Tier = [string]$r.sku.tier; vCores = $r.sku.capacity
                StorageGB = if ($r.properties.storageProfile.storageMB) { [Math]::Round([double]$r.properties.storageProfile.storageMB / 1024, 0) } else { $null }
                State = [string]$r.properties.userVisibleState; Version = [string]$r.properties.version
            } }
        Metrics = { param($r, $row) @((M 'cpu_percent' 'CPU' -Pattern -Idle 5), (M 'memory_percent' 'Memory'),
                (M 'storage_percent' 'StoragePct' -Out 'Max' -Basis 'Max'), (M 'active_connections' 'Connections' -Out 'Avg,Max')) }
    })

$script:Catalog.Add(@{
        Key = 'CosmosDB'; Display = 'Cosmos DB Accounts'; Category = 'Databases'; Types = @('microsoft.documentdb/databaseaccounts')
        Props = { param($r)
            $caps = @($r.properties.capabilities | Where-Object { $_ } | ForEach-Object { [string]$_.name })
            $api = if ([string]$r.kind -eq 'MongoDB') { 'MongoDB' } elseif ($caps -contains 'EnableCassandra') { 'Cassandra' } elseif ($caps -contains 'EnableTable') { 'Table' } elseif ($caps -contains 'EnableGremlin') { 'Gremlin' } else { 'NoSQL' }
            $locs = @($r.properties.locations | Where-Object { $_ })
            [ordered]@{
                ApiType = $api; CapacityMode = if ($caps -contains 'EnableServerless') { 'Serverless' } else { 'Provisioned' }
                Regions = (@($locs | ForEach-Object { $_.locationName }) -join ', '); RegionCount = $locs.Count
                MultiRegionWrites = $r.properties.enableMultipleWriteLocations
                ConsistencyLevel = [string]$r.properties.consistencyPolicy.defaultConsistencyLevel
                BackupPolicy = [string]$r.properties.backupPolicy.type; ContinuousBackupTier = [string]$r.properties.backupPolicy.continuousModeProperties.tier
                FreeTier = $r.properties.enableFreeTier; AnalyticalStorage = $r.properties.enableAnalyticalStorage
                TotalThroughputLimit = $r.properties.capacity.totalThroughputLimit
            } }
        Metrics = { param($r, $row)
            $m = @((M 'TotalRequestUnits' 'RU' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'NormalizedRUConsumption' 'NormalizedRUPct' -Out 'Avg,P95,Max' -Basis 'Max'),
                (M 'ThrottledRequestPercentage' 'ThrottledPct' -Out 'Avg,P95,Max'),
                (M 'ServerSideLatencyDirect' 'LatencyMs' -Out 'Avg,P95'),
                (M 'TotalRequests' 'Requests' -Out 'Total'),
                (M 'DataUsage' 'DataGB' -Scale 1GB -Out 'Max' -Basis 'Max'),
                (M 'IndexUsage' 'IndexGB' -Scale 1GB -Out 'Max' -Basis 'Max'))
            if ($row.CapacityMode -ne 'Serverless') {
                $f = "DatabaseName eq '*' and CollectionName eq '*'"
                $m += @((M 'ProvisionedThroughput' 'ProvisionedRU' -Out 'Max' -Basis 'Max' -Filter $f),
                    (M 'AutoscaleMaxThroughput' 'AutoscaleMaxRU' -Out 'Max' -Basis 'Max' -Filter $f))
            }
            $m }
        Derive = { param($row, $series)
            $prov = 0.0; $auto = 0; $n = 0
            foreach ($k in @($series.Keys)) {
                if ($k -like 'ProvisionedRU|*' -and $null -ne $series[$k].Max) { $prov += [double]$series[$k].Max; $n++ }
                if ($k -like 'AutoscaleMaxRU|*' -and [double]($series[$k].Max ?? 0) -gt 0) { $auto++ }
            }
            $row.ProvisionedRUps = if ($n) { $prov } else { $null }
            $row.ThroughputScopes = $n; $row.AutoscaleScopes = $auto
            $row.AvgConsumedRUps = if ($null -ne $row.RU_Total) { [Math]::Round([double]$row.RU_Total / ($TotalHours * 3600), 1) } else { $null }
            $row.PeakHourRUps = if ($null -ne $row.RU_MaxT) { [Math]::Round([double]$row.RU_MaxT / 3600, 1) } else { $null }
            $row.RUUtilPct_Avg = Get-Ratio $row.AvgConsumedRUps $row.ProvisionedRUps }
    })

$script:Catalog.Add(@{
        Key = 'RedisCaches'; Display = 'Azure Cache for Redis'; Category = 'Databases'; Types = @('microsoft.cache/redis')
        Props = { param($r)
            $fam = [string]$r.properties.sku.family; $cap = $r.properties.sku.capacity; $size = "$fam$cap"
            $memMap = @{ C0 = 0.25; C1 = 1; C2 = 2.5; C3 = 6; C4 = 13; C5 = 26; C6 = 53; P1 = 6; P2 = 13; P3 = 26; P4 = 53; P5 = 120 }
            [ordered]@{
                CacheSku = [string]$r.properties.sku.name; CacheSize = $size; MemoryGBPerShard = $memMap[$size]
                ShardCount = $r.properties.shardCount; ReplicasPerMaster = $r.properties.replicasPerMaster
                RedisVersion = [string]$r.properties.redisVersion
            } }
        Metrics = { param($r, $row) @(
                (M 'serverLoad' 'ServerLoad' -Pattern -Idle 5),
                (M 'percentProcessorTime' 'CPU'),
                (M 'usedmemorypercentage' 'MemoryUsedPct'),
                (M 'connectedclients' 'ConnectedClients' -Out 'Avg,Max'),
                (M 'totalcommandsprocessed' 'Commands' -Out 'Total'),
                (M 'cachehits' 'CacheHits' -Out 'Total'),
                (M 'cachemisses' 'CacheMisses' -Out 'Total'),
                (M 'evictedkeys' 'EvictedKeys' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'ManagedRedis'; Display = 'Azure Managed Redis / Redis Enterprise'; Category = 'Databases'; Types = @('microsoft.cache/redisenterprise')
        Props = { param($r) [ordered]@{ CacheSku = [string]$r.sku.name; Capacity = $r.sku.capacity; HighAvailability = [string]$r.properties.highAvailability } }
    })

# ---------------------------------------------------------------- STORAGE
$script:Catalog.Add(@{
        Key = 'ManagedDisks'; Display = 'Managed Disks'; Category = 'Storage'; Types = @('microsoft.compute/disks')
        Props = { param($r)
            $state = [string]$r.properties.diskState
            $lou = [string]$r.properties.LastOwnershipUpdateTime
            [ordered]@{
                DiskSizeGB = $r.properties.diskSizeGB; DiskTier = [string]$r.properties.tier; DiskSku = [string]$r.sku.name
                DiskState = $state; AttachedTo = Get-NameFromId ([string]$r.managedBy)
                IsOsDisk = -not [string]::IsNullOrEmpty([string]$r.properties.osType); OsType = [string]$r.properties.osType
                ProvisionedIOPS = $r.properties.diskIOPSReadWrite; ProvisionedMBps = $r.properties.diskMBpsReadWrite
                BurstingEnabled = $r.properties.burstingEnabled; MaxShares = $r.properties.maxShares
                TimeCreated = [string]$r.properties.timeCreated
                LastOwnershipUpdate = $lou; DaysUnattached = if ($state -eq 'Unattached') { Get-AgeDays $lou } else { $null }
                ManagedByService = Get-ManagedByService $r; AttachedToId = [string]$r.managedBy
            } }
        Metrics = { param($r, $row)
            if ($row.DiskState -ne 'Attached') { return @() }
            @((M 'Composite Disk Read Operations/sec' 'ReadIOPS'),
                (M 'Composite Disk Write Operations/sec' 'WriteIOPS'),
                (M 'Composite Disk Read Bytes/sec' 'ReadMBps' -Scale 1MB),
                (M 'Composite Disk Write Bytes/sec' 'WriteMBps' -Scale 1MB)) }
        Derive = { param($row, $series)
            $row.TotalIOPS_Avg = Get-SumNonNull $row.ReadIOPS_Avg $row.WriteIOPS_Avg
            $row.TotalIOPS_P95 = Get-SumNonNull $row.ReadIOPS_P95 $row.WriteIOPS_P95
            $row.TotalIOPS_Max = Get-SumNonNull $row.ReadIOPS_Max $row.WriteIOPS_Max
            $row.TotalMBps_P95 = Get-SumNonNull $row.ReadMBps_P95 $row.WriteMBps_P95
            $row.TotalMBps_Max = Get-SumNonNull $row.ReadMBps_Max $row.WriteMBps_Max
            $row.IOPSUtilPct_P95 = Get-Ratio $row.TotalIOPS_P95 $row.ProvisionedIOPS
            $row.IOPSUtilPct_Max = Get-Ratio $row.TotalIOPS_Max $row.ProvisionedIOPS
            $row.MBpsUtilPct_P95 = Get-Ratio $row.TotalMBps_P95 $row.ProvisionedMBps }
    })

$script:Catalog.Add(@{
        Key = 'Snapshots'; Display = 'Disk Snapshots'; Category = 'Storage'; Types = @('microsoft.compute/snapshots')
        Props = { param($r)
            $src = [string]$r.properties.creationData.sourceResourceId
            [ordered]@{
                SizeGB = $r.properties.diskSizeGB; SnapshotSku = [string]$r.sku.name; Incremental = $r.properties.incremental
                TimeCreated = [string]$r.properties.timeCreated; AgeDays = Get-AgeDays ([string]$r.properties.timeCreated)
                SourceDisk = Get-NameFromId $src
                SourceExists = if ($src -match '/disks/') { $script:DiskIdSet.Contains($src.ToLower()) } else { $null }
                SourceResourceId = $src
            } }
    })

$script:Catalog.Add(@{
        Key = 'StorageAccounts'; Display = 'Storage Accounts'; Category = 'Storage'; Types = @('microsoft.storage/storageaccounts')
        Props = { param($r)
            $sku = [string]$r.sku.name
            [ordered]@{
                StorageKind = [string]$r.kind; StorageSku = $sku; Performance = [string]$r.sku.tier; Redundancy = ($sku -split '_')[-1]
                AccessTier = [string]$r.properties.accessTier; HnsEnabled = $r.properties.isHnsEnabled
                LargeFileShares = [string]$r.properties.largeFileSharesState; NfsV3 = $r.properties.isNfsV3Enabled
                SftpEnabled = $r.properties.isSftpEnabled; PublicNetworkAccess = [string]$r.properties.publicNetworkAccess
                CreationTime = [string]$r.properties.creationTime; ManagedByService = Get-ManagedByService $r
            } }
        Metrics = { param($r, $row)
            $m = @((M 'UsedCapacity' 'UsedCapacityGB' -Scale 1GB -Out 'Avg,Max'),
                (M 'Transactions' 'Transactions' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'Ingress' 'IngressGB' -Scale 1GB -Out 'Total'),
                (M 'Egress' 'EgressGB' -Scale 1GB -Out 'Total'),
                (M 'SuccessE2ELatency' 'E2ELatencyMs' -Out 'Avg,P95'),
                (M 'Availability' 'AvailabilityPct' -Out 'Avg,MinAvg'))
            if ($row.StorageKind -ne 'FileStorage') {
                $m += @((M 'BlobCapacity' 'BlobCapacityGB' -Scale 1GB -Out 'Avg' -SubPath '/blobServices/default' -Filter "Tier eq '*'"),
                    (M 'BlobCount' 'BlobCount' -Out 'Avg' -SubPath '/blobServices/default' -Filter "Tier eq '*'"))
            }
            if ($row.StorageKind -notin @('BlobStorage', 'BlockBlobStorage')) {
                $m += @((M 'FileCapacity' 'FileCapacityGB' -Scale 1GB -Out 'Avg,Max' -SubPath '/fileServices/default'),
                    (M 'FileShareCount' 'FileShareCount' -Out 'Max' -SubPath '/fileServices/default'),
                    (M 'FileShareCapacityQuota' 'FileQuotaGB' -Scale 1GB -Out 'Max' -SubPath '/fileServices/default'),
                    (M 'Transactions' 'FileTransactions' -Out 'Total' -SubPath '/fileServices/default'))
            }
            $m }
        Derive = { param($row, $series)
            $tiers = [ordered]@{}; $tot = 0.0; $cnt = 0.0
            foreach ($k in @($series.Keys)) {
                if ($k -like 'BlobCapacityGB|*') { $t = ($k -split '\|', 2)[1]; $v = [double]($series[$k].Avg ?? 0); $tiers[$t] = $v; $tot += $v }
                if ($k -like 'BlobCount|*') { $cnt += [double]($series[$k].Avg ?? 0) }
            }
            $row.BlobTotalGB = if ($tiers.Count) { [Math]::Round($tot, 2) } else { $null }
            $row.BlobHotGB = Get-Round $tiers['Hot'] 2; $row.BlobCoolGB = Get-Round $tiers['Cool'] 2
            $row.BlobColdGB = Get-Round $tiers['Cold'] 2; $row.BlobArchiveGB = Get-Round $tiers['Archive'] 2
            $row.BlobTierBreakdown = (@($tiers.GetEnumerator() | ForEach-Object { "$($_.Key)=$([Math]::Round($_.Value, 1)) GB" }) -join '; ')
            $row.BlobCountTotal = if ($tiers.Count) { [long]$cnt } else { $null }
            $row.FileQuotaUsedPct = Get-Ratio $row.FileCapacityGB_Max $row.FileQuotaGB_Max }
    })

$script:Catalog.Add(@{
        Key = 'NetAppCapacityPools'; Display = 'Azure NetApp Files Capacity Pools'; Category = 'Storage'; Types = @('microsoft.netapp/netappaccounts/capacitypools')
        Props = { param($r) [ordered]@{
                ServiceLevel = [string]$r.properties.serviceLevel; PoolSizeTiB = if ($r.properties.size) { [Math]::Round([double]$r.properties.size / 1TB, 2) } else { $null }
                QosType = [string]$r.properties.qosType; CoolAccess = $r.properties.coolAccess
            } }
        Metrics = { param($r, $row) @(
                (M 'VolumePoolAllocatedUsed' 'AllocatedUsedTiB' -Scale 1TB -Out 'Avg,Max' -Basis 'Max'),
                (M 'VolumePoolTotalLogicalSize' 'LogicalSizeTiB' -Scale 1TB -Out 'Avg,Max' -Basis 'Max')) }
        Derive = { param($row, $series) $row.PoolUsedPct = Get-Ratio $row.LogicalSizeTiB_Max $row.PoolSizeTiB }
    })

$script:Catalog.Add(@{
        Key = 'RecoveryServicesVaults'; Display = 'Recovery Services Vaults'; Category = 'Storage'; Types = @('microsoft.recoveryservices/vaults')
        Props = { param($r)
            $b = $script:BackupItemsByVault[([string]$r.id).ToLower()]
            [ordered]@{
                VaultSku = [string]$r.sku.name
                StorageRedundancy = [string]($r.properties.redundancySettings.standardTierStorageRedundancy ?? '')
                ProtectedItems = [int]($b.Count ?? 0); ProtectedByWorkload = $b.ByWorkload
                ItemsOrphanedSource = [int]($b.Orphaned ?? 0); ItemsProtectionStopped = [int]($b.Stopped ?? 0)
            } }
    })

# ---------------------------------------------------------------- MONITORING & SECURITY
$script:Catalog.Add(@{
        Key = 'LogAnalyticsWorkspaces'; Display = 'Log Analytics Workspaces & Sentinel'; Category = 'Monitoring & Security'; Types = @('microsoft.operationalinsights/workspaces')
        Props = { param($r)
            $sent = $script:SentinelByWs[([string]$r.id).ToLower()]
            $cap = $r.properties.workspaceCapping.dailyQuotaGb
            [ordered]@{
                PricingTier = [string]$r.properties.sku.name; CommitmentGBPerDay = $r.properties.sku.capacityReservationLevel
                RetentionDays = $r.properties.retentionInDays
                DailyCapGB = if ($null -ne $cap -and [double]$cap -ge 0) { $cap } else { $null }
                SentinelEnabled = [bool]$sent; SentinelPricing = $sent.Pricing; SentinelCommitmentGBPerDay = $sent.Level
                DedicatedCluster = Get-NameFromId ([string]$r.properties.features.clusterResourceId)
                WorkspaceId = [string]$r.properties.customerId; CreatedDate = [string]$r.properties.createdDate
                LastSkuUpdate = [string]$r.properties.sku.lastSkuUpdate
            } }
        Metrics = { param($r, $row) @((M 'Query Count' 'QueryCount' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0), (M 'Query Failure Count' 'QueryFailures' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'LogAnalyticsClusters'; Display = 'Log Analytics Dedicated Clusters'; Category = 'Monitoring & Security'; Types = @('microsoft.operationalinsights/clusters')
        Props = { param($r) [ordered]@{ CommitmentGBPerDay = $r.sku.capacity; ClusterSku = [string]$r.sku.name; BillingType = [string]$r.properties.billingType } }
    })

$script:Catalog.Add(@{
        Key = 'ApplicationInsights'; Display = 'Application Insights'; Category = 'Monitoring & Security'; Types = @('microsoft.insights/components')
        Props = { param($r) [ordered]@{
                ApplicationType = [string]$r.properties.Application_Type; IngestionMode = [string]$r.properties.IngestionMode
                WorkspaceName = Get-NameFromId ([string]$r.properties.WorkspaceResourceId)
                IsClassic = [string]::IsNullOrEmpty([string]$r.properties.WorkspaceResourceId)
                RetentionDays = $r.properties.RetentionInDays; SamplingPct = $r.properties.SamplingPercentage
            } }
        Metrics = { param($r, $row) @((M 'requests/count' 'Requests' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'requests/failed' 'FailedRequests' -Out 'Total'), (M 'requests/duration' 'RequestDurationMs' -Out 'Avg,P95'),
                (M 'exceptions/count' 'Exceptions' -Out 'Total'), (M 'dependencies/failed' 'FailedDependencies' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'ManagedHSMs'; Display = 'Key Vault Managed HSMs'; Category = 'Monitoring & Security'; Types = @('microsoft.keyvault/managedhsms')
        Props = { param($r) [ordered]@{ HsmSku = [string]$r.sku.name; SecurityDomainState = [string]$r.properties.securityDomainProperties.activationStatus } }
        Metrics = { param($r, $row) @((M 'ServiceApiHit' 'ApiHits' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0)) }
    })

# ---------------------------------------------------------------- MESSAGING & INTEGRATION
$script:Catalog.Add(@{
        Key = 'EventHubNamespaces'; Display = 'Event Hubs Namespaces'; Category = 'Messaging & Integration'; Types = @('microsoft.eventhub/namespaces')
        Props = { param($r) [ordered]@{
                Tier = [string]$r.sku.tier; Capacity = $r.sku.capacity; AutoInflate = $r.properties.isAutoInflateEnabled
                MaxThroughputUnits = $r.properties.maximumThroughputUnits; KafkaEnabled = $r.properties.kafkaEnabled
                Status = [string]$r.properties.status
                DedicatedCluster = Get-NameFromId ([string]$r.properties.clusterArmId)
            } }
        Metrics = { param($r, $row)
            $m = @((M 'IncomingMessages' 'InMsgs' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'OutgoingMessages' 'OutMsgs' -Out 'Total'),
                (M 'ServerErrors' 'ServerErrors' -Out 'Total'),
                (M 'QuotaExceededErrors' 'QuotaExceeded' -Out 'Total'),
                (M 'IncomingBytes' 'InMB' -Scale 1MB -Out 'Total,MaxT' -Basis 'Total'),
                (M 'OutgoingBytes' 'OutMB' -Scale 1MB -Out 'Total,MaxT' -Basis 'Total'),
                (M 'ThrottledRequests' 'Throttled' -Out 'Total'),
                (M 'ActiveConnections' 'Connections' -Out 'Avg,Max'))
            if ($row.Tier -eq 'Premium') { $m += @((M 'NamespaceCpuUsage' 'CPU'), (M 'NamespaceMemoryUsage' 'Memory')) }
            $m }
        Derive = { param($row, $series)
            $row.PeakHourIngressMBps = if ($null -ne $row.InMB_MaxT) { [Math]::Round([double]$row.InMB_MaxT / 3600, 3) } else { $null }
            $row.PeakHourEgressMBps = if ($null -ne $row.OutMB_MaxT) { [Math]::Round([double]$row.OutMB_MaxT / 3600, 3) } else { $null }
            $cap = [double]($row.Capacity ?? 0)
            $row.TUUtilPct_PeakHour = if ($row.Tier -in @('Basic', 'Standard') -and $cap -gt 0) {
                [Math]::Round([Math]::Max([double]((Get-Ratio $row.PeakHourIngressMBps $cap) ?? 0), [double]((Get-Ratio $row.PeakHourEgressMBps (2 * $cap)) ?? 0)), 1)
            } else { $null } }
    })

$script:Catalog.Add(@{
        Key = 'ServiceBusNamespaces'; Display = 'Service Bus Namespaces'; Category = 'Messaging & Integration'; Types = @('microsoft.servicebus/namespaces')
        Props = { param($r) [ordered]@{ Tier = [string]$r.sku.tier; Capacity = $r.sku.capacity; Status = [string]$r.properties.status } }
        Metrics = { param($r, $row)
            $m = @((M 'IncomingMessages' 'InMsgs' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'OutgoingMessages' 'OutMsgs' -Out 'Total'),
                (M 'ActiveConnections' 'Connections' -Out 'Avg,Max'),
                (M 'ActiveMessages' 'ActiveMessages' -Out 'Avg,Max'),
                (M 'DeadletteredMessages' 'DeadLettered' -Out 'Max' -Basis 'Max'),
                (M 'ServerErrors' 'ServerErrors' -Out 'Total'),
                (M 'ThrottledRequests' 'Throttled' -Out 'Total'))
            if ($row.Tier -eq 'Premium') { $m += @((M 'NamespaceCpuUsage' 'CPU'), (M 'NamespaceMemoryUsage' 'Memory')) }
            $m }
    })

$script:Catalog.Add(@{
        Key = 'LogicApps'; Display = 'Logic Apps (Consumption)'; Category = 'Messaging & Integration'; Types = @('microsoft.logic/workflows')
        Props = { param($r) [ordered]@{
                State = [string]$r.properties.state; WorkflowSku = [string]$r.properties.sku.name
                IntegrationAccount = Get-NameFromId ([string]$r.properties.integrationAccount.id); ChangedTime = [string]$r.properties.changedTime
            } }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Disabled') { return @() }
            @((M 'RunsStarted' 'Runs' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'RunsFailed' 'RunsFailed' -Out 'Total'),
                (M 'TotalBillableExecutions' 'BillableExecutions' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'APIManagement'; Display = 'API Management'; Category = 'Messaging & Integration'; Types = @('microsoft.apimanagement/service')
        Props = { param($r) [ordered]@{
                ApimSku = [string]$r.sku.name; Units = $r.sku.capacity; PlatformVersion = [string]$r.properties.platformVersion
                VNetType = [string]$r.properties.virtualNetworkType; AdditionalLocations = Get-Count $r.properties.additionalLocations
            } }
        Metrics = { param($r, $row)
            if ($row.ApimSku -eq 'Consumption') { return @((M 'Requests' 'Requests' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0)) }
            @((M 'Capacity' 'Capacity' -Pattern -Idle 5), (M 'Requests' 'Requests' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'IoTHubs'; Display = 'IoT Hubs'; Category = 'Messaging & Integration'; Types = @('microsoft.devices/iothubs')
        Props = { param($r)
            $per = switch ([string]$r.sku.name) { 'S1' { 400000 } 'S2' { 6000000 } 'S3' { 300000000 } 'B1' { 400000 } 'B2' { 6000000 } 'B3' { 300000000 } 'F1' { 8000 } default { $null } }
            [ordered]@{ IotSku = [string]$r.sku.name; Units = $r.sku.capacity; DailyMessageQuota = if ($per) { $per * [int]($r.sku.capacity ?? 1) } else { $null } } }
        Metrics = { param($r, $row) @(
                (M 'dailyMessageQuotaUsed' 'DailyMsgsUsed' -Out 'Max' -Basis 'Max'),
                (M 'd2c.telemetry.ingress.success' 'TelemetryMsgs' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'devices.connectedDevices.allProtocol' 'ConnectedDevices' -Out 'Avg,Max')) }
        Derive = { param($row, $series) $row.QuotaUtilPct_Max = Get-Ratio $row.DailyMsgsUsed_Max $row.DailyMessageQuota }
    })

$script:Catalog.Add(@{
        Key = 'StreamAnalyticsJobs'; Display = 'Stream Analytics Jobs'; Category = 'Messaging & Integration'; Types = @('microsoft.streamanalytics/streamingjobs')
        Props = { param($r) [ordered]@{
                JobState = [string]$r.properties.jobState; JobSku = [string]$r.properties.sku.name
                StreamingUnits = $r.properties.transformation.properties.streamingUnits; LastOutputEventTime = [string]$r.properties.lastOutputEventTime
            } }
        Metrics = { param($r, $row) @((M 'ResourceUtilization' 'SUUtilPct' -Pattern -Idle 5), (M 'InputEvents' 'InputEvents' -Out 'Total'), (M 'OutputEvents' 'OutputEvents' -Out 'Total')) }
    })

# ---------------------------------------------------------------- NETWORKING
$script:Catalog.Add(@{
        Key = 'ApplicationGateways'; Display = 'Application Gateways'; Category = 'Networking'; Types = @('microsoft.network/applicationgateways')
        Props = { param($r)
            $pools = @($r.properties.backendAddressPools | Where-Object { $_ })
            $targets = 0; foreach ($p in $pools) { $targets += (Get-Count $p.properties.backendAddresses) + (Get-Count $p.properties.backendIPConfigurations) }
            [ordered]@{
                GatewaySku = [string]$r.properties.sku.name; GatewayTier = [string]$r.properties.sku.tier; FixedCapacity = $r.properties.sku.capacity
                AutoscaleMin = $r.properties.autoscaleConfiguration.minCapacity; AutoscaleMax = $r.properties.autoscaleConfiguration.maxCapacity
                BackendPools = $pools.Count; BackendTargets = $targets; Listeners = Get-Count $r.properties.httpListeners
                OperationalState = [string]$r.properties.operationalState
            } }
        Metrics = { param($r, $row)
            if ($row.OperationalState -eq 'Stopped') { return @() }
            if ($row.GatewayTier -match 'v2') {
                @((M 'CapacityUnits' 'CapacityUnits' -Pattern -Idle 1), (M 'ComputeUnits' 'ComputeUnits'),
                    (M 'Throughput' 'ThroughputMbps' -Scale 125000), (M 'TotalRequests' 'Requests' -Out 'Total'),
                    (M 'CurrentConnections' 'Connections' -Out 'Avg,Max'))
            }
            else { @((M 'Throughput' 'ThroughputMbps' -Scale 125000 -Pattern -Idle 0.01), (M 'TotalRequests' 'Requests' -Out 'Total'), (M 'CurrentConnections' 'Connections' -Out 'Avg,Max')) } }
    })

$script:Catalog.Add(@{
        Key = 'LoadBalancers'; Display = 'Load Balancers'; Category = 'Networking'; Types = @('microsoft.network/loadbalancers')
        Props = { param($r)
            $pools = @($r.properties.backendAddressPools | Where-Object { $_ })
            $targets = 0; foreach ($p in $pools) { $targets += (Get-Count $p.properties.loadBalancerBackendAddresses) + (Get-Count $p.properties.backendIPConfigurations) }
            $fe = @($r.properties.frontendIPConfigurations | Where-Object { $_ })
            [ordered]@{
                LbSku = [string]$r.sku.name; Exposure = if (@($fe | Where-Object { $_.properties.publicIPAddress }).Count) { 'Public' } else { 'Internal' }
                FrontendIPs = $fe.Count; BackendPools = $pools.Count; BackendTargets = $targets
                LbRules = Get-Count $r.properties.loadBalancingRules; NatRules = Get-Count $r.properties.inboundNatRules
                OutboundRules = Get-Count $r.properties.outboundRules; ManagedByService = Get-ManagedByService $r
            } }
        Metrics = { param($r, $row)
            if ($row.LbSku -ne 'Standard') { return @() }
            @((M 'ByteCount' 'DataGB' -Scale 1GB -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'PacketCount' 'Packets' -Out 'Total'), (M 'VipAvailability' 'DataPathAvailability' -Out 'Avg,Min')) }
    })

$script:Catalog.Add(@{
        Key = 'PublicIPAddresses'; Display = 'Public IP Addresses'; Category = 'Networking'; Types = @('microsoft.network/publicipaddresses')
        Props = { param($r)
            $cfg = [string]$r.properties.ipConfiguration.id; $nat = [string]$r.properties.natGateway.id
            $attachedType = if ($cfg) { if ($cfg -match '/providers/Microsoft\.[^/]+/([^/]+)/') { $Matches[1] } else { 'Unknown' } } elseif ($nat) { 'natGateways' } else { $null }
            [ordered]@{
                PipSku = [string]$r.sku.name; PipTier = [string]$r.sku.tier; AllocationMethod = [string]$r.properties.publicIPAllocationMethod
                IpAddress = [string]$r.properties.ipAddress; IsAttached = [bool]($cfg -or $nat); AttachedToType = $attachedType
                AttachedTo = if ($cfg) { ($cfg -split '/')[8] } elseif ($nat) { Get-NameFromId $nat } else { $null }
                DdosProtectionMode = [string]$r.properties.ddosSettings.protectionMode
                ManagedByService = Get-ManagedByService $r
            } }
    })

$script:Catalog.Add(@{
        Key = 'NatGateways'; Display = 'NAT Gateways'; Category = 'Networking'; Types = @('microsoft.network/natgateways')
        Props = { param($r) [ordered]@{
                NatSku = [string]$r.sku.name; PublicIPs = (Get-Count $r.properties.publicIpAddresses) + (Get-Count $r.properties.publicIpPrefixes)
                Subnets = Get-Count $r.properties.subnets; IdleTimeoutMin = $r.properties.idleTimeoutInMinutes
            } }
        Metrics = { param($r, $row)
            if ($row.Subnets -eq 0) { return @() }
            @((M 'ByteCount' 'DataGB' -Scale 1GB -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0), (M 'TotalConnectionCount' 'Connections' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'AzureFirewalls'; Display = 'Azure Firewalls'; Category = 'Networking'; Types = @('microsoft.network/azurefirewalls')
        Props = { param($r) [ordered]@{
                FirewallTier = [string]$r.properties.sku.tier; FirewallSkuName = [string]$r.properties.sku.name
                IpConfigurations = Get-Count $r.properties.ipConfigurations; ThreatIntelMode = [string]$r.properties.threatIntelMode
                FirewallPolicy = Get-NameFromId ([string]$r.properties.firewallPolicy.id); VirtualHub = Get-NameFromId ([string]$r.properties.virtualHub.id)
            } }
        Metrics = { param($r, $row)
            if ($row.IpConfigurations -eq 0 -and -not $row.VirtualHub) { return @() }
            @((M 'Throughput' 'ThroughputMbps' -Scale 1e6 -Pattern -Idle 0.1), (M 'DataProcessed' 'DataProcessedGB' -Scale 1GB -Out 'Total'),
                (M 'ApplicationRuleHit' 'AppRuleHits' -Out 'Total'), (M 'NetworkRuleHit' 'NetRuleHits' -Out 'Total'),
                (M 'SNATPortUtilization' 'SnatPortPct' -Out 'P95,Max'), (M 'FirewallHealth' 'HealthPct' -Out 'Avg,MinAvg')) }
    })

$script:Catalog.Add(@{
        Key = 'VirtualNetworkGateways'; Display = 'VPN & ExpressRoute Gateways'; Category = 'Networking'; Types = @('microsoft.network/virtualnetworkgateways')
        Props = { param($r)
            $c = $script:GwConnections[([string]$r.id).ToLower()]
            [ordered]@{
                GatewayType = [string]$r.properties.gatewayType; VpnType = [string]$r.properties.vpnType; GatewaySku = [string]$r.properties.sku.name
                Generation = [string]$r.properties.vpnGatewayGeneration; ActiveActive = $r.properties.activeActive; BgpEnabled = $r.properties.enableBgp
                Connections = [int]($c ?? 0)
                P2SConfigured = (Get-Count $r.properties.vpnClientConfiguration.vpnClientAddressPool.addressPrefixes) -gt 0
            } }
        Metrics = { param($r, $row)
            if ($row.GatewayType -eq 'ExpressRoute') { @((M 'ExpressRouteGatewayBitsPerSecond' 'BitsMbps' -Scale 1e6 -Pattern -Idle 0.01), (M 'ExpressRouteGatewayCpuUtilization' 'CPU')) }
            else { @((M 'AverageBandwidth' 'AvgBandwidthMbps' -Scale 125000 -Pattern -Idle 0.01), (M 'P2SConnectionCount' 'P2SConnections' -Out 'Avg,Max'), (M 'TunnelAverageBandwidth' 'TunnelMbps' -Scale 125000)) } }
    })

$script:Catalog.Add(@{
        Key = 'ExpressRouteCircuits'; Display = 'ExpressRoute Circuits'; Category = 'Networking'; Types = @('microsoft.network/expressroutecircuits')
        Props = { param($r)
            $bw = $r.properties.serviceProviderProperties.bandwidthInMbps
            if (-not $bw -and $r.properties.bandwidthInGbps) { $bw = [double]$r.properties.bandwidthInGbps * 1000 }
            [ordered]@{
                BandwidthMbps = $bw; CircuitTier = [string]$r.sku.tier; CircuitFamily = [string]$r.sku.family
                Provider = [string]$r.properties.serviceProviderProperties.serviceProviderName
                PeeringLocation = [string]$r.properties.serviceProviderProperties.peeringLocation
                ProviderState = [string]$r.properties.serviceProviderProvisioningState; Peerings = Get-Count $r.properties.peerings
            } }
        Metrics = { param($r, $row) @((M 'BitsInPerSecond' 'InMbps' -Scale 1e6 -Pattern -Idle 0.01), (M 'BitsOutPerSecond' 'OutMbps' -Scale 1e6)) }
        Derive = { param($row, $series)
            $row.PeakUtilPct = Get-Ratio ([Math]::Max([double]($row.InMbps_Max ?? 0), [double]($row.OutMbps_Max ?? 0))) $row.BandwidthMbps
            $row.P95UtilPct = Get-Ratio ([Math]::Max([double]($row.InMbps_P95 ?? 0), [double]($row.OutMbps_P95 ?? 0))) $row.BandwidthMbps }
    })

$script:Catalog.Add(@{
        Key = 'BastionHosts'; Display = 'Azure Bastion'; Category = 'Networking'; Types = @('microsoft.network/bastionhosts')
        Props = { param($r) [ordered]@{ BastionSku = [string]$r.sku.name; ScaleUnits = $r.properties.scaleUnits; Tunneling = $r.properties.enableTunneling } }
        Metrics = { param($r, $row) if ($row.BastionSku -eq 'Developer') { return @() }; @((M 'sessions' 'Sessions' -Out 'Avg,Max' -Pattern -Idle 0)) }
    })

$script:Catalog.Add(@{
        Key = 'PrivateEndpoints'; Display = 'Private Endpoints'; Category = 'Networking'; Types = @('microsoft.network/privateendpoints')
        Props = { param($r)
            $conns = @(@($r.properties.privateLinkServiceConnections) + @($r.properties.manualPrivateLinkServiceConnections) | Where-Object { $_ })
            $c = $conns | Select-Object -First 1
            $target = [string]$c.properties.privateLinkServiceId
            [ordered]@{
                ConnectionState = [string]$c.properties.privateLinkServiceConnectionState.status
                TargetResource = Get-NameFromId $target; TargetType = if ($target -match '/providers/([^/]+/[^/]+)/') { $Matches[1] } else { $null }
                GroupIds = (@($c.properties.groupIds) -join ','); Subnet = Get-NameFromId ([string]$r.properties.subnet.id)
                ManagedByService = Get-ManagedByService $r
            } }
        Metrics = { param($r, $row) @((M 'PEBytesIn' 'BytesInMB' -Scale 1MB -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0), (M 'PEBytesOut' 'BytesOutMB' -Scale 1MB -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'FrontDoorCdn'; Display = 'Front Door & CDN Profiles'; Category = 'Networking'; Types = @('microsoft.cdn/profiles', 'microsoft.network/frontdoors')
        Props = { param($r) [ordered]@{ ProfileSku = if ($r.sku.name) { [string]$r.sku.name } else { 'Classic Front Door' }; ProfileState = [string]($r.properties.resourceState ?? $r.properties.enabledState) } }
        Metrics = { param($r, $row)
            $sizeMetric = if ($r.type -eq 'microsoft.network/frontdoors') { 'BillableResponseSize' } else { 'ResponseSize' }
            @((M 'RequestCount' 'Requests' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0), (M $sizeMetric 'ResponseGB' -Scale 1GB -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'DdosProtectionPlans'; Display = 'DDoS Protection Plans'; Category = 'Networking'; Types = @('microsoft.network/ddosprotectionplans')
        Props = { param($r) [ordered]@{
                ProtectedVNets = Get-Count $r.properties.virtualNetworks
                VNetNames = (@($r.properties.virtualNetworks | Where-Object { $_ } | ForEach-Object { Get-NameFromId $_.id }) -join ', ')
            } }
    })

$script:Catalog.Add(@{
        Key = 'VirtualHubs'; Display = 'Virtual WAN Hubs'; Category = 'Networking'; Types = @('microsoft.network/virtualhubs')
        Props = { param($r) [ordered]@{
                HubSku = [string]$r.properties.sku; RoutingInfraUnits = $r.properties.virtualRouterAutoScaleConfiguration.minCapacity
                VpnGateway = Get-NameFromId ([string]$r.properties.vpnGateway.id); ErGateway = Get-NameFromId ([string]$r.properties.expressRouteGateway.id)
                Firewall = Get-NameFromId ([string]$r.properties.azureFirewall.id)
            } }
    })

# ---------------------------------------------------------------- ANALYTICS & AI
$script:Catalog.Add(@{
        Key = 'DataFactories'; Display = 'Data Factories'; Category = 'Analytics & AI'; Types = @('microsoft.datafactory/factories')
        PreCollect = { param($resources)
            $items = foreach ($r in $resources) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)/integrationRuntimes?api-version=2018-06-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'Data Factory integration runtimes'
            foreach ($k in $res.Keys) {
                if (-not $res[$k].Ok) { continue }
                $irs = @($res[$k].Data.value | Where-Object { $_ })
                $managed = @($irs | Where-Object { $_.properties.type -eq 'Managed' })
                $script:AdfIRs[$k] = @{
                    Summary = (@($irs | ForEach-Object {
                                $p = $_.properties; $cp = $p.typeProperties.computeProperties
                                if ($p.type -eq 'Managed') { "$($_.name) [Managed $($cp.nodeSize) x$($cp.numberOfNodes) $($p.state)$(if ($p.typeProperties.ssisProperties) { ' SSIS' })]" } else { "$($_.name) [$($p.type) $($p.state)]" }
                            }) -join '; ')
                    SsisRunning = @($managed | Where-Object { $_.properties.typeProperties.ssisProperties -and $_.properties.state -eq 'Started' }).Count
                }
            } }
        Props = { param($r)
            $ir = $script:AdfIRs[([string]$r.id).ToLower()]
            [ordered]@{ IntegrationRuntimes = $ir.Summary; RunningSsisIRs = $ir.SsisRunning; CreateTime = [string]$r.properties.createTime } }
        Metrics = { param($r, $row) @(
                (M 'PipelineSucceededRuns' 'PipelineRunsOK' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'PipelineFailedRuns' 'PipelineRunsFailed' -Out 'Total'),
                (M 'ActivitySucceededRuns' 'ActivityRuns' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'SynapseSqlPools'; Display = 'Synapse Dedicated SQL Pools'; Category = 'Analytics & AI'; Types = @('microsoft.synapse/workspaces/sqlpools')
        Props = { param($r) [ordered]@{ Workspace = ($r.id -split '/')[8]; PoolSku = [string]$r.sku.name; Status = [string]$r.properties.status; CreationDate = [string]$r.properties.creationDate } }
        Metrics = { param($r, $row) @((M 'DWUUsedPercent' 'DWUUsedPct' -Pattern -Idle 2), (M 'CPUPercent' 'CPU'), (M 'MemoryUsedPercent' 'Memory'), (M 'ActiveQueries' 'ActiveQueries' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'SynapseSparkPools'; Display = 'Synapse Spark Pools'; Category = 'Analytics & AI'; Types = @('microsoft.synapse/workspaces/bigdatapools')
        Props = { param($r) [ordered]@{
                Workspace = ($r.id -split '/')[8]; NodeSize = [string]$r.properties.nodeSize; NodeCount = $r.properties.nodeCount
                AutoScale = $r.properties.autoScale.enabled; AutoScaleMin = $r.properties.autoScale.minNodeCount; AutoScaleMax = $r.properties.autoScale.maxNodeCount
                AutoPause = $r.properties.autoPause.enabled; AutoPauseDelayMin = $r.properties.autoPause.delayInMinutes; SparkVersion = [string]$r.properties.sparkVersion
            } }
    })

$script:Catalog.Add(@{
        Key = 'DataExplorerClusters'; Display = 'Data Explorer Clusters'; Category = 'Analytics & AI'; Types = @('microsoft.kusto/clusters')
        Props = { param($r) [ordered]@{
                ClusterSku = [string]$r.sku.name; Instances = $r.sku.capacity; State = [string]$r.properties.state
                OptimizedAutoscale = $r.properties.optimizedAutoscale.isEnabled; AutoscaleMin = $r.properties.optimizedAutoscale.minimum
                AutoscaleMax = $r.properties.optimizedAutoscale.maximum; AutoStop = $r.properties.enableAutoStop
            } }
        Metrics = { param($r, $row)
            if ($row.State -eq 'Stopped') { return @() }
            @((M 'CPU' 'CPU' -Pattern -Idle 5), (M 'CacheUtilization' 'CacheUtilPct'), (M 'IngestionUtilization' 'IngestionUtilPct' -Out 'Avg,Max')) }
    })

$script:Catalog.Add(@{
        Key = 'Databricks'; Display = 'Databricks Workspaces'; Category = 'Analytics & AI'; Types = @('microsoft.databricks/workspaces')
        Props = { param($r) [ordered]@{ WorkspaceSku = [string]$r.sku.name; ManagedResourceGroup = Get-NameFromId ([string]$r.properties.managedResourceGroupId) } }
    })

$script:Catalog.Add(@{
        Key = 'HDInsightClusters'; Display = 'HDInsight Clusters'; Category = 'Analytics & AI'; Types = @('microsoft.hdinsight/clusters')
        Props = { param($r)
            $roles = @($r.properties.computeProfile.roles | Where-Object { $_ })
            [ordered]@{
                ClusterKind = [string]$r.properties.clusterDefinition.kind; ClusterState = [string]$r.properties.clusterState; Tier = [string]$r.properties.tier
                Roles = (@($roles | ForEach-Object { "$($_.name):$($_.hardwareProfile.vmSize) x$($_.targetInstanceCount)" }) -join '; ')
                TotalNodes = [int](($roles | ForEach-Object { [int]($_.targetInstanceCount ?? 0) } | Measure-Object -Sum).Sum)
            } }
        Metrics = { param($r, $row) @((M 'NumActiveWorkers' 'ActiveWorkers' -Out 'Avg,Max' -Pattern -Idle 0), (M 'CategorizedGatewayRequests' 'GatewayRequests' -Out 'Total')) }
    })

$script:Catalog.Add(@{
        Key = 'AnalysisServices'; Display = 'Analysis Services'; Category = 'Analytics & AI'; Types = @('microsoft.analysisservices/servers')
        Props = { param($r) [ordered]@{ ServerSku = [string]$r.sku.name; Tier = [string]$r.sku.tier; Replicas = $r.sku.capacity; State = [string]$r.properties.state } }
        Metrics = { param($r, $row) if ($row.State -eq 'Paused') { return @() }; @((M 'qpu_metric' 'QPU' -Pattern -Idle 1), (M 'memory_metric' 'MemoryGB' -Scale 1GB)) }
    })

$script:Catalog.Add(@{
        Key = 'PowerBIFabricCapacities'; Display = 'Power BI Embedded & Fabric Capacities'; Category = 'Analytics & AI'
        Types = @('microsoft.powerbidedicated/capacities', 'microsoft.fabric/capacities')
        Props = { param($r) [ordered]@{ CapacitySku = [string]$r.sku.name; Tier = [string]$r.sku.tier; State = [string]$r.properties.state } }
        Metrics = { param($r, $row)
            if ($r.type -ne 'microsoft.powerbidedicated/capacities' -or $row.State -eq 'Paused') { return @() }
            @((M 'cpu_metric' 'CPU' -Pattern -Idle 2), (M 'overload_metric' 'Overload' -Out 'Max')) }
    })

$script:Catalog.Add(@{
        Key = 'CognitiveServices'; Display = 'Azure AI Services & OpenAI'; Category = 'Analytics & AI'; Types = @('microsoft.cognitiveservices/accounts')
        PreCollect = { param($resources)
            $ai = @($resources | Where-Object { $_.kind -in @('OpenAI', 'AIServices') })
            $items = foreach ($r in $ai) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)/deployments?api-version=2024-10-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'AI model deployments'
            foreach ($r in $ai) {
                $k = ([string]$r.id).ToLower(); $x = $res[$k]
                if (-not $x -or -not $x.Ok) { continue }
                $deps = @($x.Data.value | Where-Object { $_ })
                $script:AiDeployments[$k] = $deps
                foreach ($d in $deps) {
                    $skuName = [string]$d.sku.name
                    $script:AiDeploymentRows.Add([ordered]@{
                            SubscriptionName = $script:SubNames[[string]$r.subscriptionId]; SubscriptionId = [string]$r.subscriptionId
                            ResourceGroup = [string]$r.resourceGroup; ResourceName = [string]$r.name; ResourceType = 'AI Model Deployments'
                            Location = [string]$r.location; Kind = [string]$r.kind
                            Deployment = [string]$d.name; Model = [string]$d.properties.model.name; ModelVersion = [string]$d.properties.model.version
                            DeploymentSku = $skuName; Capacity = $d.sku.capacity; IsProvisioned = ($skuName -match 'Provisioned')
                            PTUs = if ($skuName -match 'Provisioned') { $d.sku.capacity } else { $null }
                            _Dim = [string]$d.name; _Id = [string]$r.id; _Tags = ''
                        })
                }
            } }
        Props = { param($r)
            $deps = @($script:AiDeployments[([string]$r.id).ToLower()] | Where-Object { $_ })
            $ptu = @($deps | Where-Object { [string]$_.sku.name -match 'Provisioned' })
            [ordered]@{
                ServiceKind = [string]$r.kind; AccountSku = [string]$r.sku.name; CustomDomain = [string]$r.properties.customSubDomainName
                Deployments = $deps.Count; ProvisionedDeployments = $ptu.Count
                TotalPTUs = [int](($ptu | ForEach-Object { [int]($_.sku.capacity ?? 0) } | Measure-Object -Sum).Sum)
            } }
        Metrics = { param($r, $row)
            $m = @((M 'TotalCalls' 'Calls' -Out 'Total' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0), (M 'TotalErrors' 'Errors' -Out 'Total'), (M 'Latency' 'LatencyMs' -Out 'Avg,P95'))
            if ($row.ServiceKind -in @('OpenAI', 'AIServices')) {
                $m += @((M 'ProcessedPromptTokens' 'PromptTokens' -Out 'Total'), (M 'GeneratedTokens' 'GeneratedTokens' -Out 'Total'), (M 'TotalTokens' 'TotalTokens' -Out 'Total'))
                if ($row.ProvisionedDeployments -gt 0) {
                    $m += (M 'AzureOpenAIProvisionedManagedUtilizationV2' 'PTUUtilPct' -Out 'Avg,P95,Max' -Filter "ModelDeploymentName eq '*'")
                }
            }
            $m }
        Derive = { param($row, $series)
            $vals = @(foreach ($k in @($series.Keys)) { if ($k -like 'PTUUtilPct|*') { $series[$k].P95 } })
            $row.PTUUtilPct_P95_Max = if ($vals.Count) { Get-Round (($vals | Measure-Object -Maximum).Maximum) 1 } else { $null } }
    })

$script:Catalog.Add(@{
        Key = 'SearchServices'; Display = 'Azure AI Search'; Category = 'Analytics & AI'; Types = @('microsoft.search/searchservices')
        Props = { param($r) [ordered]@{
                SearchSku = [string]$r.sku.name; Replicas = $r.properties.replicaCount; Partitions = $r.properties.partitionCount
                SearchUnits = [int]($r.properties.replicaCount ?? 1) * [int]($r.properties.partitionCount ?? 1)
                HostingMode = [string]$r.properties.hostingMode; SemanticSearch = [string]$r.properties.semanticSearch
            } }
        Metrics = { param($r, $row)
            if ($row.SearchSku -eq 'free') { return @() }
            @((M 'SearchQueriesPerSecond' 'QPS' -Out 'Avg,P95,Max' -Pattern -Idle 0.01), (M 'ThrottledSearchQueriesPercentage' 'ThrottledPct' -Out 'Avg,Max'), (M 'SearchLatency' 'LatencySec' -Out 'Avg,P95')) }
    })

$script:Catalog.Add(@{
        Key = 'MLWorkspaces'; Display = 'Machine Learning & AI Foundry Workspaces'; Category = 'Analytics & AI'; Types = @('microsoft.machinelearningservices/workspaces')
        PreCollect = { param($resources)
            $items = foreach ($r in $resources) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)/computes?api-version=2024-04-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'ML computes'
            foreach ($r in $resources) {
                $x = $res[([string]$r.id).ToLower()]
                if (-not $x -or -not $x.Ok) { continue }
                foreach ($c in @($x.Data.value | Where-Object { $_ })) {
                    $p = $c.properties; $pp = $p.properties
                    $script:MlComputeRows.Add([ordered]@{
                            SubscriptionName = $script:SubNames[[string]$r.subscriptionId]; SubscriptionId = [string]$r.subscriptionId
                            ResourceGroup = [string]$r.resourceGroup; ResourceName = [string]$r.name; ResourceType = 'ML Compute'; Location = [string]$r.location
                            ComputeName = [string]$c.name; ComputeType = [string]$p.computeType; VmSize = [string]$pp.vmSize
                            State = [string]($pp.state ?? $p.provisioningState)
                            MinNodes = $pp.scaleSettings.minNodeCount; MaxNodes = $pp.scaleSettings.maxNodeCount; CurrentNodes = $pp.currentNodeCount
                            IdleShutdown = [string]($pp.idleTimeBeforeShutdown ?? $pp.scaleSettings.nodeIdleTimeBeforeScaleDown)
                            LastOperation = [string]$pp.lastOperation.operationName; LastOperationTime = [string]$pp.lastOperation.operationTime
                            _Id = [string]$r.id; _Tags = ''
                        })
                }
            } }
        Props = { param($r) [ordered]@{ WorkspaceKind = [string]$r.kind; WorkspaceSku = [string]$r.sku.name } }
    })

$script:Catalog.Add(@{
        Key = 'BatchAccounts'; Display = 'Batch Accounts (Pools)'; Category = 'Compute'; Types = @('microsoft.batch/batchaccounts')
        PreCollect = { param($resources)
            $items = foreach ($r in $resources) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)/pools?api-version=2024-07-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'Batch pools'
            foreach ($k in $res.Keys) { if ($res[$k].Ok) { $script:BatchPools[$k] = @($res[$k].Data.value | Where-Object { $_ }) } }
        }
        Props = { param($r)
            $pools = @($script:BatchPools[([string]$r.id).ToLower()] | Where-Object { $_ })
            [ordered]@{
                PoolAllocationMode = [string]$r.properties.poolAllocationMode; Pools = $pools.Count
                PoolSummary = (@($pools | ForEach-Object { $p = $_.properties; "$($_.name) [$($p.vmSize) dedicated $($p.currentDedicatedNodes) / spot $($p.currentLowPriorityNodes)$(if ($p.scaleSettings.autoScale) { ' autoscale' } else { ' fixed' })]" }) -join '; ')
                DedicatedNodes = [int](($pools | ForEach-Object { [int]($_.properties.currentDedicatedNodes ?? 0) } | Measure-Object -Sum).Sum)
                LowPriorityNodes = [int](($pools | ForEach-Object { [int]($_.properties.currentLowPriorityNodes ?? 0) } | Measure-Object -Sum).Sum)
                FixedScalePoolsWithNodes = @($pools | Where-Object { -not $_.properties.scaleSettings.autoScale -and [int]($_.properties.currentDedicatedNodes ?? 0) -gt 0 }).Count
                DedicatedCoreQuota = $r.properties.dedicatedCoreQuota
            } }
        Metrics = { param($r, $row) @(
                (M 'TotalNodeCount' 'TotalNodes' -Out 'Avg,Max' -Pattern -Idle 0),
                (M 'IdleNodeCount' 'IdleNodes' -Out 'Avg,Max'),
                (M 'RunningNodeCount' 'RunningNodes' -Out 'Avg,Max'),
                (M 'CoreCount' 'Cores' -Out 'Avg,Max'),
                (M 'TaskCompleteEvent' 'TasksCompleted' -Out 'Total')) }
        Derive = { param($row, $series) $row.IdleNodePct_Avg = Get-Ratio $row.IdleNodes_Avg $row.TotalNodes_Avg }
    })

$script:Catalog.Add(@{
        Key = 'DedicatedHosts'; Display = 'Dedicated Hosts'; Category = 'Compute'; Types = @('microsoft.compute/hostgroups/hosts')
        PreCollect = { param($resources)
            $items = foreach ($r in $resources) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)?`$expand=instanceView&api-version=2024-03-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'Dedicated host instance views'
            foreach ($k in $res.Keys) { if ($res[$k].Ok) { $script:HostViews[$k] = $res[$k].Data } }
        }
        Props = { param($r)
            $v = $script:HostViews[([string]$r.id).ToLower()]
            $vms = @($r.properties.virtualMachines | Where-Object { $_ })
            $alloc = @($v.properties.instanceView.availableCapacity.allocatableVMs | Where-Object { $_ })
            [ordered]@{
                HostSku = [string]$r.sku.name; HostGroup = ($r.id -split '/')[8]; VMsOnHost = $vms.Count
                VMs = (@($vms | ForEach-Object { Get-NameFromId ([string]$_.id) }) -join ', ')
                RemainingCapacity = (@($alloc | Where-Object { $_.count -gt 0 } | ForEach-Object { "$($_.vmSize) x$($_.count)" }) -join '; ')
                LicenseType = [string]$r.properties.licenseType; AutoReplaceOnFailure = $r.properties.autoReplaceOnFailure
                ProvisioningTime = [string]$r.properties.provisioningTime
            } }
    })

$script:Catalog.Add(@{
        Key = 'CapacityReservations'; Display = 'On-demand Capacity Reservations'; Category = 'Compute'; Types = @('microsoft.compute/capacityreservationgroups/capacityreservations')
        PreCollect = { param($resources)
            $items = foreach ($r in $resources) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)?`$expand=instanceView&api-version=2024-03-01"; Method = 'GET'; TokenKey = 'arm' } }
            $res = Invoke-ParallelRest -Items @($items) -Label 'Capacity reservation instance views'
            foreach ($k in $res.Keys) { if ($res[$k].Ok) { $script:CapResViews[$k] = $res[$k].Data } }
        }
        Props = { param($r)
            $v = $script:CapResViews[([string]$r.id).ToLower()]
            $allocated = @($v.properties.instanceView.utilizationInfo.virtualMachinesAllocated | Where-Object { $_ }).Count
            $reserved = [int]($r.sku.capacity ?? 0)
            [ordered]@{
                VmSize = [string]$r.sku.name; ReservedQuantity = $reserved; VMsAllocated = $allocated
                ReservationUtilizationPct = if ($reserved -gt 0) { [Math]::Round(100.0 * $allocated / $reserved, 1) } else { $null }
                UnusedReservedVMs = [Math]::Max(0, $reserved - $allocated)
                ReservationGroup = ($r.id -split '/')[8]; TimeCreated = [string]$r.properties.timeCreated
            } }
    })

$script:Catalog.Add(@{
        Key = 'ElasticSans'; Display = 'Elastic SAN'; Category = 'Storage'; Types = @('microsoft.elasticsan/elasticsans')
        Props = { param($r)
            $base = [double]($r.properties.baseSizeTiB ?? 0); $ext = [double]($r.properties.extendedCapacitySizeTiB ?? 0)
            [ordered]@{
                SanSku = [string]$r.sku.name; BaseSizeTiB = $base; ExtendedCapacityTiB = $ext; ProvisionedTiB = $base + $ext
                VolumeGroups = $r.properties.volumeGroupCount; AllocatedVolumeGiB = $r.properties.totalVolumeSizeGiB
                AllocatedPct = if (($base + $ext) -gt 0 -and $null -ne $r.properties.totalVolumeSizeGiB) { [Math]::Round(100.0 * [double]$r.properties.totalVolumeSizeGiB / 1024 / ($base + $ext), 1) } else { $null }
                ProvisionedIops = $r.properties.totalIops; ProvisionedMBps = $r.properties.totalMBps
            } }
        Metrics = { param($r, $row) @(
                (M 'ElasticSanUsedCapacity' 'UsedCapacityTiB' -Scale 1TB -Out 'Max' -Basis 'Max'),
                (M 'Transactions' 'Transactions' -Out 'Total,MaxT' -Basis 'Total' -Pattern -PatternAgg 'Total' -Idle 0),
                (M 'Ingress' 'IngressGB' -Scale 1GB -Out 'Total'),
                (M 'Egress' 'EgressGB' -Scale 1GB -Out 'Total'),
                (M 'SuccessServerLatency' 'ServerLatencyMs' -Out 'Avg,P95')) }
        Derive = { param($row, $series) $row.UsedPct = Get-Ratio $row.UsedCapacityTiB_Max $row.ProvisionedTiB }
    })

$script:Catalog.Add(@{
        Key = 'SignalRWebPubSub'; Display = 'SignalR & Web PubSub'; Category = 'Messaging & Integration'; Types = @('microsoft.signalrservice/signalr', 'microsoft.signalrservice/webpubsub')
        Props = { param($r) [ordered]@{
                ServiceKind = if ([string]$r.type -match 'webpubsub') { 'Web PubSub' } else { 'SignalR' }; Tier = [string]$r.sku.tier; ServiceSku = [string]$r.sku.name
                Units = [int]($r.sku.capacity ?? 1); ConnectionCapacity = 1000 * [int]($r.sku.capacity ?? 1)
                ServiceMode = [string](@($r.properties.features | Where-Object { $_.flag -eq 'ServiceMode' } | Select-Object -First 1).value)
            } }
        Metrics = { param($r, $row)
            if ($row.Tier -eq 'Free') { return @() }
            @((M 'ServerLoad' 'ServerLoad' -Out 'Avg,P95,Max' -Pattern -Idle 2),
                (M 'ConnectionCount' 'Connections' -Out 'Avg,Max' -Basis 'Max'),
                (M 'MessageCount' 'Messages' -Out 'Total')) }
        Derive = { param($row, $series) $row.ConnectionUtilPct_Max = Get-Ratio $row.Connections_Max $row.ConnectionCapacity }
    })

$script:Catalog.Add(@{
        Key = 'AzureVirtualDesktop'; Display = 'Azure Virtual Desktop Host Pools'; Category = 'Compute'; Types = @('microsoft.desktopvirtualization/hostpools')
        Props = { param($r) [ordered]@{
                HostPoolType = [string]$r.properties.hostPoolType; LoadBalancer = [string]$r.properties.loadBalancerType
                MaxSessionLimit = $r.properties.maxSessionLimit; StartVMOnConnect = $r.properties.startVMOnConnect
            } }
        Metrics = $null
    })

#endregion

#region ============================== COLLECTION ENGINE ==============================

function New-BaseRow($r, $cat) {
    [ordered]@{
        SubscriptionName = $script:SubNames[[string]$r.subscriptionId]
        SubscriptionId   = [string]$r.subscriptionId
        ResourceGroup    = [string]$r.resourceGroup
        ResourceName     = [string]$r.name
        ResourceType     = $cat.Display
        ArmType          = [string]$r.type
        Location         = [string]$r.location
        EnvironmentTag   = Get-TagValue $r.tags @('environment', 'env', 'stage', 'environmenttype', 'environment-type')
    }
}

function New-MetricItem($r, [object[]]$Specs) {
    $groups = [ordered]@{}
    foreach ($s in $Specs) {
        $gk = "$($s.SubPath)|$($s.Filter)"
        if (-not $groups.Contains($gk)) { $groups[$gk] = [System.Collections.Generic.List[object]]::new() }
        $groups[$gk].Add($s)
    }
    $calls = foreach ($gk in $groups.Keys) {
        $list = $groups[$gk]
        for ($i = 0; $i -lt $list.Count; $i += 20) {
            $chunk = @($list[$i..([Math]::Min($i + 19, $list.Count - 1))])
            $specMap = @{}; foreach ($s in $chunk) { $specMap[$s.Name] = $s }
            @{ SubPath = $chunk[0].SubPath; Filter = $chunk[0].Filter; Names = @($chunk | ForEach-Object { $_.Name }); Specs = $specMap }
        }
    }
    @{ Id = [string]$r.id; Type = [string]$r.type; Location = [string]$r.location; SubscriptionId = [string]$r.subscriptionId; Calls = @($calls) }
}

function Get-MetricsTokenSafe {
    if (-not $script:BatchMetrics) { return $null }
    try { return Get-BearerToken -Resource $script:MetricsResource }
    catch { Write-Log "Metrics Batch API token unavailable - using per-resource metrics ($($_.Exception.Message))" Warning; $script:BatchMetrics = $false; return $null }
}

function Invoke-MetricCollection([object[]]$Items, [string]$Label) {
    $results = @{}
    if (-not $Items -or $Items.Count -eq 0) { return $results }
    $workerText = $script:MetricWorkerText
    $metricsToken = Get-MetricsTokenSafe
    # Build work: Metrics Batch API groups (same subscription / region / namespace / metric set), per-resource ARM otherwise
    $work = [System.Collections.Generic.List[object]]::new()
    $batchSize = [Math]::Max(1, [int][Math]::Floor($MetricsBatchSize * 30.0 / [Math]::Max(30, $LookbackDays)))
    if ($metricsToken) {
        $groups = [ordered]@{}
        foreach ($it in $Items) {
            $loc = (([string]$it.Location).ToLower() -replace '\s', '')
            if ([string]::IsNullOrWhiteSpace($loc) -or $loc -eq 'global') { $work.Add(@{ Mode = 'Single'; Id = $it.Id; Calls = $it.Calls }); continue }
            foreach ($call in $it.Calls) {
                $ns = ([string]$it.Type).ToLower()
                if ($call.SubPath) { $ns = "$ns/" + (($call.SubPath.Trim('/') -split '/')[0]).ToLower() }
                $sig = (@($call.Names) | ForEach-Object { $s = $call.Specs[$_]; "$($s.Name)~$($s.Alias)~$($s.Out)~$($s.Scale)~$($s.Basis)~$($s.Pattern)~$($s.PatternAgg)~$($s.Idle)" }) -join '|'
                $key = "$($it.SubscriptionId)|$loc|$ns|$($call.SubPath)|$($call.Filter)|$sig".ToLower()
                if (-not $groups.Contains($key)) { $groups[$key] = @{ Sub = $it.SubscriptionId; Region = $loc; Ns = $ns; Call = $call; Ids = [System.Collections.Generic.List[string]]::new() } }
                $groups[$key].Ids.Add([string]$it.Id)
            }
        }
        foreach ($g in $groups.Values) {
            for ($i = 0; $i -lt $g.Ids.Count; $i += $batchSize) {
                $ids = @($g.Ids[$i..([Math]::Min($i + $batchSize - 1, $g.Ids.Count - 1))])
                $work.Add(@{ Mode = 'Batch'; SubscriptionId = $g.Sub; Region = $g.Region; Namespace = $g.Ns; SubPath = $g.Call.SubPath; Filter = $g.Call.Filter
                        Names = $g.Call.Names; Specs = $g.Call.Specs; ResourceIds = @($ids | ForEach-Object { "$_$($g.Call.SubPath)" }); OriginalIds = $ids })
            }
        }
    }
    else { foreach ($it in $Items) { $work.Add(@{ Mode = 'Single'; Id = $it.Id; Calls = $it.Calls }) } }

    $slice = 200
    $done = 0
    for ($b = 0; $b -lt $work.Count; $b += $slice) {
        $part = @($work[$b..([Math]::Min($b + $slice - 1, $work.Count - 1))])
        $ctx = @{
            ArmToken = (Get-BearerToken); MetricsToken = $(if ($metricsToken) { Get-BearerToken -Resource $script:MetricsResource } else { $null })
            ArmBase = $script:ArmBase; Timespan = $Timespan; StartIso = $StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); EndIso = $EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); StartUtc = $StartUtc
            TotalHours = $TotalHours; Days = $LookbackDays; DayIdx = $HourDayIndex; LocalHour = $HourLocalHour; LocalDow = $HourLocalDow
            WeekendDows = $WeekendDowInts; BizStart = $BusinessHoursStart; BizEnd = $BusinessHoursEnd
        }
        $out = $part | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $it = $_
            try { $sb = [scriptblock]::Create($using:workerText); & $sb $it $using:ctx }
            catch {
                $msg = "Worker failure: $($_.Exception.Message)"
                $ids = if ($it.Mode -eq 'Batch') { @($it.OriginalIds) } else { @($it.Id) }
                [pscustomobject]@{ Requests = 0; Results = @($ids | ForEach-Object { @{ Id = ([string]$_).ToLower(); Series = @{}; Errors = [System.Collections.Generic.List[string]]@($msg); Calls = 0; CallsOk = 0 } }) }
            }
        }
        foreach ($o in $out) {
            if ($null -eq $o) { continue }
            $script:MetricRequests += [int]$o.Requests
            foreach ($a in @($o.Results)) {
                if ($null -eq $a) { continue }
                $id = ([string]$a.Id).ToLower()
                $acc = $results[$id]
                if (-not $acc) { $acc = @{ Id = $id; Series = @{}; ErrorList = [System.Collections.Generic.List[string]]::new(); Calls = 0; CallsOk = 0 }; $results[$id] = $acc }
                foreach ($k in $a.Series.Keys) { $acc.Series[$k] = $a.Series[$k] }
                foreach ($e in @($a.Errors)) { if ($e) { $acc.ErrorList.Add([string]$e) } }
                $acc.Calls += [int]$a.Calls; $acc.CallsOk += [int]$a.CallsOk
            }
        }
        $done += $part.Count
        Write-Log "    $Label metrics: $done/$($work.Count) $(if ($metricsToken) { 'batch requests' } else { 'resources' }) completed" Info
    }
    foreach ($acc in $results.Values) {
        $acc.Status = if ($acc.Series.Count -gt 0 -and $acc.ErrorList.Count -eq 0) { 'OK' } elseif ($acc.Series.Count -gt 0) { 'Partial' } elseif ($acc.ErrorList.Count -gt 0 -and $acc.CallsOk -eq 0) { 'Error' } else { 'NoData' }
        $acc.Errors = ($acc.ErrorList | Select-Object -Unique) -join ' | '
    }
    return $results
}

function Format-Num($v) {
    if ($null -eq $v) { return '' }
    try { return ([double]$v).ToString('0.####', [Globalization.CultureInfo]::InvariantCulture) } catch { return [string]$v }
}

function Format-CsvField($v) {
    if ($null -eq $v) { return '' }
    $s = [string]$v
    if ($s -match '[",\r\n]') { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

$script:DayNames = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')

function Write-TypeCsv {
    param([System.Collections.IEnumerable]$Rows, [string]$Path)
    $tail = @('MetricsStatus', 'MetricsErrors', 'MetricCalls', 'Tags', 'ResourceId')
    $final = foreach ($row in $Rows) {
        $o = [ordered]@{}
        foreach ($k in $row.Keys) { if (-not ([string]$k).StartsWith('_') -and $k -notin $tail) { $o[$k] = $row[$k] } }
        foreach ($t in $tail) { if ($t -eq 'Tags') { $o.Tags = $row._Tags } elseif ($t -eq 'ResourceId') { $o.ResourceId = $row._Id } else { $o[$t] = $row[$t] } }
        $o
    }
    $final = @($final)
    # Column order: identity / configuration / metrics -> Tag.* (priority keys first, capped) -> Prop.* (most populated, capped) -> tail
    $core = [System.Collections.Generic.List[string]]::new(); $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $tagFill = @{}; $propFill = @{}
    foreach ($o in $final) {
        foreach ($k in $o.Keys) {
            if ($k -in $tail) { continue }
            if ($k.StartsWith('Tag.')) { $tagFill[$k] = [int]($tagFill[$k] ?? 0) + 1; continue }
            if ($k.StartsWith('Prop.') -or $k.StartsWith('Plan.')) { $propFill[$k] = [int]($propFill[$k] ?? 0) + 1; continue }
            if ($seen.Add($k)) { $core.Add($k) }
        }
    }
    $prio = @{}; for ($i = 0; $i -lt $script:PriorityTagKeys.Count; $i++) { $prio["Tag.$($script:PriorityTagKeys[$i])".ToLower()] = $i }
    $tagCols = @($tagFill.Keys | Sort-Object @{ e = { if ($prio.ContainsKey($_.ToLower())) { $prio[$_.ToLower()] } else { 1000 } } }, @{ e = { $tagFill[$_] }; Descending = $true }, @{ e = { $_ } } | Select-Object -First $script:MaxTagColumns)
    $propCols = @($propFill.Keys | Sort-Object @{ e = { $propFill[$_] }; Descending = $true }, @{ e = { $_ } } | Select-Object -First $script:MaxPropColumns | Sort-Object)
    $cols = @($core) + $tagCols + $propCols + $tail
    $out = foreach ($o in $final) {
        $p = [ordered]@{}
        foreach ($c in $cols) {
            $v = $o[$c]
            if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { $v = [Math]::Round([double]$v, 4) }
            elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = (@($v) -join '; ') }
            $p[$c] = $v
        }
        [pscustomobject]$p
    }
    $out | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8BOM
}

#endregion

#region ============================== SUBSCRIPTIONS ==============================

Write-Log ('=' * 78) Section
Write-Log 'PHASE 1 - COST OPTIMIZATION & PERFORMANCE UTILIZATION DATA COLLECTION' Section
Write-Log ('=' * 78) Section
Write-Log "Output folder : $OutputPath"
Write-Log "Window        : $($StartLocal.ToString('yyyy-MM-dd')) -> $($EndLocal.AddDays(-1).ToString('yyyy-MM-dd')) ($LookbackDays full days, $TimeZoneId)"
Write-Log "UTC timespan  : $Timespan ($TotalHours hourly buckets)"
Write-Log "Usage pattern : business hours $($BusinessHoursStart):00-$($BusinessHoursEnd):00, weekend = $($WeekendDays -join ', ')"

$allSubs = @(Get-AllPages "$script:ArmBase/subscriptions?api-version=2022-12-01" | Where-Object { $_.state -eq 'Enabled' -and (-not $_.tenantId -or $_.tenantId -eq $TenantId) } |
        ForEach-Object { [pscustomobject]@{ Id = [string]$_.subscriptionId; Name = [string]$_.displayName; State = [string]$_.state; TenantId = [string]$_.tenantId } })
if ($SubscriptionIds) {
    $wanted = @($SubscriptionIds | ForEach-Object { $_.ToLower() })
    $allSubs = @($allSubs | Where-Object { $wanted -contains $_.Id.ToLower() })
    $missing = @($wanted | Where-Object { $_ -notin @($allSubs | ForEach-Object { $_.Id.ToLower() }) })
    foreach ($m in $missing) { Write-Log "Subscription $m not found / not enabled in tenant $TenantId - skipped" Warning }
}
if ($allSubs.Count -eq 0) { throw "No enabled subscriptions found in scope (tenant $TenantId)." }
Write-Log "Authentication  : $script:AuthSource ($(if ($script:AuthSource -eq 'AzureCLI') { $script:CliAccount.user.name } else { $script:AzContext.Account.Id }))"

$script:SubIds = @($allSubs | ForEach-Object { $_.Id })
$script:SubNames = @{}; foreach ($s in $allSubs) { $script:SubNames[$s.Id] = $s.Name }
Write-Log "Subscriptions in scope: $($allSubs.Count)" Success
$allSubs | ForEach-Object { [pscustomobject]@{ SubscriptionName = $_.Name; SubscriptionId = $_.Id; State = $_.State; TenantId = $_.TenantId } } |
    Export-Csv -Path (Join-Path $RawPath 'Subscriptions.csv') -NoTypeInformation -Encoding utf8BOM

#endregion

#region ============================== LOOKUPS & INVENTORY ==============================

Write-Log 'Building lookups (autoscale, gateways, Sentinel, backups, registries)...' Section
$script:Autoscale = @{}; $script:GwConnections = @{}; $script:SentinelByWs = @{}; $script:DiskIdSet = [System.Collections.Generic.HashSet[string]]::new()
$script:AcrReplications = @{}; $script:BackupItemsByVault = @{}; $script:PoolDbCount = @{}; $script:AdfIRs = @{}
$script:BatchPools = @{}; $script:HostViews = @{}; $script:CapResViews = @{}
$script:AiDeployments = @{}; $script:AiDeploymentRows = [System.Collections.Generic.List[object]]::new()
$script:MlComputeRows = [System.Collections.Generic.List[object]]::new()

function Invoke-Lookup([string]$Name, [string]$Query) {
    try { return Invoke-ARGQuery $Query } catch { Add-CollectionError 'Lookup' $Name $_.Exception.Message; Write-Log "  Lookup '$Name' failed: $($_.Exception.Message)" Warning; return @() }
}

foreach ($a in (Invoke-Lookup 'Autoscale' "resources | where type =~ 'microsoft.insights/autoscalesettings' | project target = tolower(tostring(properties.targetResourceUri)), enabled = tobool(properties.enabled), profiles = properties.profiles")) {
    $cap = @($a.profiles)[0].capacity
    $script:Autoscale[$a.target] = @{ Enabled = $a.enabled; Min = $cap.minimum; Max = $cap.maximum; Default = $cap.default }
}
foreach ($c in (Invoke-Lookup 'Connections' "resources | where type =~ 'microsoft.network/connections' | project gw1 = tolower(tostring(properties.virtualNetworkGateway1.id)), gw2 = tolower(tostring(properties.virtualNetworkGateway2.id))")) {
    foreach ($g in @($c.gw1, $c.gw2)) { if ($g) { $script:GwConnections[$g] = [int]($script:GwConnections[$g] ?? 0) + 1 } }
}
foreach ($s in (Invoke-Lookup 'Sentinel' "resources | where type =~ 'microsoft.operationsmanagement/solutions' and name startswith 'SecurityInsights' | project ws = tolower(tostring(properties.workspaceResourceId)), skuName = tostring(properties.sku.name), level = properties.sku.capacityReservationLevel")) {
    if (-not $s.ws) { continue }
    $pricing = switch -Regex ([string]$s.skuName) {
        '^(?i)unified$' { 'Simplified (unified)' } '^(?i)capacityreservation$' { 'Classic - Commitment tier' } '^(?i)pergb' { 'Classic - Pay-as-you-go' } default { 'Not specified' }
    }
    $script:SentinelByWs[$s.ws] = @{ Pricing = $pricing; Level = $s.level }
}
foreach ($d in (Invoke-Lookup 'Disks' "resources | where type =~ 'microsoft.compute/disks' | project id = tolower(id)")) { [void]$script:DiskIdSet.Add([string]$d.id) }
foreach ($x in (Invoke-Lookup 'AcrReplications' "resources | where type =~ 'microsoft.containerregistry/registries/replications' | extend registry = tolower(substring(id, 0, indexof(tolower(id), '/replications/'))) | summarize n = count() by registry")) { $script:AcrReplications[$x.registry] = $x.n }
foreach ($x in (Invoke-Lookup 'ElasticPoolDbs' "resources | where type =~ 'microsoft.sql/servers/databases' and isnotempty(tostring(properties.elasticPoolId)) | summarize n = count() by pool = tolower(tostring(properties.elasticPoolId))")) { $script:PoolDbCount[$x.pool] = $x.n }

Write-Log 'Querying full resource inventory...' Section
$inventory = Invoke-ARGQuery "resources | project id, name, type, kind, location, resourceGroup, subscriptionId, skuName = tostring(sku.name), skuTier = tostring(sku.tier), tags | order by id asc"
$script:AllIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($r in $inventory) { [void]$script:AllIds.Add(([string]$r.id).ToLower()) }

# Resource group tags (inherited context for filtering) and resource created / changed times (resource age)
$script:RgTags = @{}
foreach ($x in (Invoke-Lookup 'ResourceGroupTags' "resourcecontainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups' | project subscriptionId, name, tags")) {
    $script:RgTags["$($x.subscriptionId)|$($x.name)".ToLower()] = Get-TagString $x.tags
}
try {
    $tok = Get-BearerToken; $armBase = $script:ArmBase
    $times = $script:SubIds | ForEach-Object -ThrottleLimit ([Math]::Min(8, $ThrottleLimit)) -Parallel {
        $u = "$($using:armBase)/subscriptions/$_/resources?`$expand=createdTime,changedTime&api-version=2021-04-01"
        $pages = 0
        while ($u -and $pages -lt 1000) {
            $pages++; $resp = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                try { $resp = Invoke-RestMethod -Uri $u -Headers @{ Authorization = ('Bearer ' + $using:tok) } -TimeoutSec 180; break } catch { Start-Sleep -Seconds ([Math]::Pow(2, $attempt)) }
            }
            if (-not $resp) { break }
            foreach ($v in $resp.value) { [pscustomobject]@{ Id = ([string]$v.id).ToLower(); Created = $v.createdTime; Changed = $v.changedTime } }
            $u = $resp.nextLink
        }
    }
    foreach ($t in $times) {
        $fmt = { param($d) if ($null -eq $d) { $null } elseif ($d -is [DateTime]) { $d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } else { [string]$d } }
        $script:ResourceTimes[$t.Id] = @{ Created = (& $fmt $t.Created); Changed = (& $fmt $t.Changed) }
    }
    Write-Log "Resource created/changed times: $($script:ResourceTimes.Count) resources" Info
}
catch { Add-CollectionError 'ResourceTimes' 'all' $_.Exception.Message; Write-Log "Resource created/changed times unavailable: $($_.Exception.Message)" Warning }

$typeToKey = @{}; foreach ($c in $script:Catalog) { foreach ($t in $c.Types) { $typeToKey[$t] = $c.Key } }
$inventory | ForEach-Object {
    $tm = $script:ResourceTimes[([string]$_.id).ToLower()]
    [pscustomobject][ordered]@{
        SubscriptionName = $script:SubNames[[string]$_.subscriptionId]; SubscriptionId = $_.subscriptionId; ResourceGroup = $_.resourceGroup
        ResourceName = $_.name; ArmType = $_.type; Kind = $_.kind; Location = $_.location; SkuName = $_.skuName; SkuTier = $_.skuTier
        CatalogKey = $typeToKey[[string]$_.type]; EnvironmentTag = Get-TagValue $_.tags @('environment', 'env', 'stage', 'environmenttype')
        CreatedTime = $tm.Created; LastChangedTime = $tm.Changed; AgeDays = if ($tm -and $tm.Created) { Get-AgeDays $tm.Created } else { $null }
        Tags = Get-TagString $_.tags; ResourceGroupTags = $script:RgTags["$($_.subscriptionId)|$($_.resourceGroup)".ToLower()]; ResourceId = $_.id
    }
} | Export-Csv -Path (Join-Path $RawPath 'Inventory_AllResources.csv') -NoTypeInformation -Encoding utf8BOM
Write-Log "Inventory: $($inventory.Count) resources, $(@($inventory | Select-Object -ExpandProperty type -Unique).Count) resource types" Success

# Backup protected items (orphaned sources = source resource no longer exists in scope)
$backupItems = Invoke-Lookup 'BackupItems' @"
recoveryservicesresources
| where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'
| extend vaultId = tolower(substring(id, 0, indexof(tolower(id), '/backupfabrics/')))
| project vaultId, subscriptionId, resourceGroup, sourceId = tolower(tostring(properties.sourceResourceId)), workload = tostring(properties.workloadType),
          state = tostring(properties.protectionState), friendlyName = tostring(properties.friendlyName), lastBackup = tostring(properties.lastBackupTime)
"@
$script:OrphanBackupRows = [System.Collections.Generic.List[object]]::new()
$subIdSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach ($s in $script:SubIds) { [void]$subIdSet.Add($s) }
foreach ($g in ($backupItems | Group-Object vaultId)) {
    $orph = 0
    foreach ($bi in $g.Group) {
        $srcSub = if ($bi.sourceId -match '^/subscriptions/([^/]+)/') { $Matches[1] } else { $null }
        if ($bi.sourceId -and $srcSub -and $subIdSet.Contains($srcSub) -and -not $script:AllIds.Contains($bi.sourceId)) {
            $orph++
            $script:OrphanBackupRows.Add([pscustomobject]@{
                    Finding = 'Backup item for deleted source resource'; CostImpact = 'Paid'; SubscriptionName = $script:SubNames[[string]$bi.subscriptionId]
                    ResourceGroup = $bi.resourceGroup; ResourceName = $bi.friendlyName; ArmType = "Backup item ($($bi.workload))"; Location = ''
                    Sku = ''; Detail = "Vault: $(Get-NameFromId $g.Name); state: $($bi.state); last backup: $($bi.lastBackup)"; ResourceId = $bi.sourceId
                })
        }
    }
    $script:BackupItemsByVault[$g.Name] = @{
        Count = $g.Count; Orphaned = $orph
        Stopped = @($g.Group | Where-Object { $_.state -match 'ProtectionStopped' }).Count
        ByWorkload = (@($g.Group | Group-Object workload | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '; ')
    }
}

#endregion

#region ============================== LOG ANALYTICS INGESTION ==============================

function ConvertFrom-LaResponse($Data) {
    $t = @($Data.tables)[0]
    if (-not $t) { return @() }
    $cols = @($t.columns | ForEach-Object { [string]$_.name })
    foreach ($rw in @($t.rows)) {
        $o = [ordered]@{}
        for ($i = 0; $i -lt $cols.Count; $i++) { $o[$cols[$i]] = $rw[$i] }
        [pscustomobject]$o
    }
}

function Get-DayLabel($v) {
    if ($v -is [DateTime]) { $dt = if ($v.Kind -eq [DateTimeKind]::Local) { $v.ToUniversalTime() } else { $v }; return $dt.ToString('yyyy-MM-dd') }
    $s = [string]$v; if ($s.Length -ge 10) { return $s.Substring(0, 10) }; return $s
}

function Add-LogAnalyticsIngestion($rows) {
    <# Daily billable ingestion per workspace from the Usage table -> commitment-tier utilization. #>
    $ws = @($rows | Where-Object { $_.WorkspaceId })
    if ($ws.Count -eq 0) { return }
    Write-Log "  Querying daily ingestion (Usage table) for $($ws.Count) workspace(s)..." Info
    $offMin = [int][Math]::Round($TzOffsetHours * 60)
    $s = $StartUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'); $e = $EndUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $qDaily = "Usage | where TimeGenerated >= datetime($s) and TimeGenerated < datetime($e) | summarize GB = sum(Quantity) / 1000.0 by Day = startofday(TimeGenerated + ${offMin}m), IsBillable, Solution"
    $qTop = "Usage | where TimeGenerated >= datetime($s) and TimeGenerated < datetime($e) | where IsBillable == true | summarize GB = sum(Quantity) / 1000.0 by DataType, Solution | top 25 by GB desc"
    $qTable = "Usage | where TimeGenerated >= datetime($s) and TimeGenerated < datetime($e) | where IsBillable == true | summarize GB = sum(Quantity) / 1000.0 by DataType, Day = startofday(TimeGenerated + ${offMin}m)"
    $items = foreach ($w in $ws) {
        $url = "$script:LaEndpoint/workspaces/$($w.WorkspaceId)/query"
        @{ Key = "$($w.WorkspaceId)|daily"; Url = $url; Method = 'POST'; TokenKey = 'la'; Body = (@{ query = $qDaily; timespan = $Timespan } | ConvertTo-Json -Compress) }
        @{ Key = "$($w.WorkspaceId)|top"; Url = $url; Method = 'POST'; TokenKey = 'la'; Body = (@{ query = $qTop; timespan = $Timespan } | ConvertTo-Json -Compress) }
        @{ Key = "$($w.WorkspaceId)|table"; Url = $url; Method = 'POST'; TokenKey = 'la'; Body = (@{ query = $qTable; timespan = $Timespan } | ConvertTo-Json -Compress) }
    }
    $res = Invoke-ParallelRest -Items @($items) -Label 'Log Analytics usage queries'
    $secSolutions = @('Security', 'SecurityInsights', 'SecurityCenterFree', 'SecurityCenter')
    $ingRows = [System.Collections.Generic.List[object]]::new()
    $topRows = [System.Collections.Generic.List[object]]::new()
    $tableRows = [System.Collections.Generic.List[object]]::new()
    $dailyWriter = [System.IO.StreamWriter]::new((Join-Path $DailyPath 'LogAnalyticsWorkspaces_Daily.csv'), $false, $script:Utf8Bom)
    $dailyWriter.WriteLine('SubscriptionName,ResourceGroup,ResourceName,ResourceType,Metric,Dimension,Date,Average,Maximum,Minimum,Total,HoursWithData,ResourceId')
    try {
        foreach ($w in $ws) {
            $d = $res["$($w.WorkspaceId)|daily"]; $t = $res["$($w.WorkspaceId)|top"]
            if (-not $d -or -not $d.Ok) {
                $w.IngestionQueryStatus = "Failed: $(if ($d) { $d.Error } else { 'no response' })"
                Add-CollectionError 'LogAnalyticsQuery' $w._Id $w.IngestionQueryStatus
                continue
            }
            $bill = @{}; $nonBill = @{}; $sec = @{}
            foreach ($x in (ConvertFrom-LaResponse $d.Data)) {
                $day = Get-DayLabel $x.Day; $gb = [double]($x.GB ?? 0)
                if (ConvertTo-Bool $x.IsBillable) {
                    $bill[$day] = [double]($bill[$day] ?? 0) + $gb
                    if ([string]$x.Solution -in $secSolutions) { $sec[$day] = [double]($sec[$day] ?? 0) + $gb }
                }
                else { $nonBill[$day] = [double]($nonBill[$day] ?? 0) + $gb }
            }
            $first = -1
            for ($i = 0; $i -lt $DayLabels.Count; $i++) { if ($bill.ContainsKey($DayLabels[$i]) -or $nonBill.ContainsKey($DayLabels[$i])) { $first = $i; break } }
            $commit = if ([string]$w.PricingTier -match '(?i)capacityreservation' -and $w.CommitmentGBPerDay) { [double]$w.CommitmentGBPerDay } else { $null }
            $sentCommit = if ($w.SentinelCommitmentGBPerDay) { [double]$w.SentinelCommitmentGBPerDay } else { $null }
            $prefix = (Format-CsvField $w.SubscriptionName) + ',' + (Format-CsvField $w.ResourceGroup) + ',' + (Format-CsvField $w.ResourceName) + ',' + (Format-CsvField $w.ResourceType) + ','
            $vals = [System.Collections.Generic.List[double]]::new(); $wk = [System.Collections.Generic.List[double]]::new(); $we = [System.Collections.Generic.List[double]]::new()
            for ($i = [Math]::Max(0, $first); $i -lt $DayLabels.Count -and $first -ge 0; $i++) {
                $lbl = $DayLabels[$i]; $b = [double]($bill[$lbl] ?? 0); $nb = [double]($nonBill[$lbl] ?? 0); $sg = [double]($sec[$lbl] ?? 0)
                $vals.Add($b)
                $dow = [int]$StartLocal.AddDays($i).DayOfWeek
                if ($WeekendDowInts -contains $dow) { $we.Add($b) } else { $wk.Add($b) }
                foreach ($pair in @(@('BillableIngestionGB', $b), @('NonBillableIngestionGB', $nb), @('SecurityBillableGB', $sg))) {
                    $dailyWriter.WriteLine($prefix + $pair[0] + ',,' + $lbl + ',' + (Format-Num $pair[1]) + ',' + (Format-Num $pair[1]) + ',' + (Format-Num $pair[1]) + ',' + (Format-Num $pair[1]) + ',24,' + (Format-CsvField $w._Id))
                }
                $ingRows.Add([pscustomobject][ordered]@{
                        SubscriptionName = $w.SubscriptionName; ResourceGroup = $w.ResourceGroup; Workspace = $w.ResourceName; Date = $lbl
                        DayOfWeek = $script:DayNames[$dow]; BillableGB = [Math]::Round($b, 3); NonBillableGB = [Math]::Round($nb, 3); SecurityGB = [Math]::Round($sg, 3)
                        PricingTier = $w.PricingTier; CommitmentGBPerDay = $commit; SentinelEnabled = $w.SentinelEnabled; SentinelPricing = $w.SentinelPricing
                        SentinelCommitmentGBPerDay = $sentCommit; ResourceId = $w._Id
                    })
            }
            $w.IngestionQueryStatus = if ($first -lt 0) { 'OK - no ingestion in window' } else { 'OK' }
            $w.ActiveDays = $vals.Count
            if ($vals.Count -gt 0) {
                $arr = $vals.ToArray(); $sorted = $vals.ToArray(); [Array]::Sort($sorted)
                $avg = ($arr | Measure-Object -Average).Average
                $w.TotalBillableGB = [Math]::Round(($arr | Measure-Object -Sum).Sum, 2)
                $w.AvgDailyBillableGB = [Math]::Round($avg, 2)
                $w.MedianDailyBillableGB = [Math]::Round($sorted[[int][Math]::Floor(($sorted.Length - 1) / 2)], 2)
                $w.P95DailyBillableGB = [Math]::Round($sorted[[Math]::Max(0, [int][Math]::Ceiling(0.95 * $sorted.Length) - 1)], 2)
                $w.MaxDailyBillableGB = [Math]::Round($sorted[-1], 2)
                $w.MinDailyBillableGB = [Math]::Round($sorted[0], 2)
                $w.WeekdayAvgGB = if ($wk.Count) { [Math]::Round(($wk | Measure-Object -Average).Average, 2) } else { $null }
                $w.WeekendAvgGB = if ($we.Count) { [Math]::Round(($we | Measure-Object -Average).Average, 2) } else { $null }
                if ($arr.Length -ge 14) {
                    $f7 = ($arr[0..6] | Measure-Object -Average).Average; $l7 = ($arr[($arr.Length - 7)..($arr.Length - 1)] | Measure-Object -Average).Average
                    $w.TrendPct_Last7vsFirst7 = if ($f7 -gt 0) { [Math]::Round(100 * ($l7 - $f7) / $f7, 1) } else { $null }
                }
                $w.AvgDailyNonBillableGB = [Math]::Round((@($nonBill.Values) | Measure-Object -Sum).Sum / $vals.Count, 2)
                $w.AvgDailySecurityGB = [Math]::Round((@($sec.Values) | Measure-Object -Sum).Sum / $vals.Count, 2)
                if ($commit) {
                    $w.CommitmentUtilPct = [Math]::Round(100 * $avg / $commit, 1)
                    $w.DaysBelowCommitment = @($arr | Where-Object { $_ -lt $commit }).Count
                    $w.DaysAboveCommitment = @($arr | Where-Object { $_ -gt $commit }).Count
                }
                if ($sentCommit) { $w.SentinelCommitmentUtilPct = [Math]::Round(100 * $avg / $sentCommit, 1) }
            }
            if ($t -and $t.Ok) {
                $tops = @(ConvertFrom-LaResponse $t.Data)
                $tot = [double](($tops | Measure-Object -Property GB -Sum).Sum)
                $w.TopTables = (@($tops | Select-Object -First 5 | ForEach-Object {
                            "$($_.DataType) ($([Math]::Round([double]$_.GB / [Math]::Max(1, $vals.Count), 2)) GB/day)"
                        }) -join '; ')
                foreach ($x in $tops) {
                    $topRows.Add([pscustomobject][ordered]@{
                            SubscriptionName = $w.SubscriptionName; ResourceGroup = $w.ResourceGroup; Workspace = $w.ResourceName
                            DataType = $x.DataType; Solution = $x.Solution; BillableGB = [Math]::Round([double]$x.GB, 3)
                            AvgDailyGB = [Math]::Round([double]$x.GB / [Math]::Max(1, $vals.Count), 3)
                            PctOfTop25 = if ($tot -gt 0) { [Math]::Round(100 * [double]$x.GB / $tot, 1) } else { $null }
                            SentinelEnabled = $w.SentinelEnabled; ResourceId = $w._Id
                        })
                }
            }
            $tb = $res["$($w.WorkspaceId)|table"]
            if ($tb -and $tb.Ok) {
                $trows = @(ConvertFrom-LaResponse $tb.Data)
                $topNames = @($trows | Group-Object DataType | Sort-Object { ($_.Group | Measure-Object -Property GB -Sum).Sum } -Descending | Select-Object -First 10 | ForEach-Object Name)
                foreach ($x in $trows) {
                    $tableRows.Add([pscustomobject][ordered]@{
                            SubscriptionName = $w.SubscriptionName; ResourceGroup = $w.ResourceGroup; Workspace = $w.ResourceName
                            Date = Get-DayLabel $x.Day; DataType = $(if ($x.DataType -in $topNames) { $x.DataType } else { 'Other tables' }); BillableGB = [Math]::Round([double]$x.GB, 4); ResourceId = $w._Id
                        })
                }
            }
        }
    }
    finally { $dailyWriter.Dispose() }
    if ($ingRows.Count) { $ingRows | Export-Csv -Path (Join-Path $RawPath 'LogAnalytics_DailyIngestion.csv') -NoTypeInformation -Encoding utf8BOM }
    if ($topRows.Count) { $topRows | Export-Csv -Path (Join-Path $RawPath 'LogAnalytics_TopTables.csv') -NoTypeInformation -Encoding utf8BOM }
    if ($tableRows.Count) {
        $tableRows | Group-Object { "$($_.ResourceId)|$($_.Date)|$($_.DataType)" } | ForEach-Object {
            $f = $_.Group[0]
            [pscustomobject][ordered]@{ SubscriptionName = $f.SubscriptionName; ResourceGroup = $f.ResourceGroup; Workspace = $f.Workspace; Date = $f.Date; DataType = $f.DataType; BillableGB = [Math]::Round(($_.Group | Measure-Object -Property BillableGB -Sum).Sum, 3); ResourceId = $f.ResourceId }
        } | Sort-Object Workspace, Date, DataType | Export-Csv -Path (Join-Path $RawPath 'LogAnalytics_DailyByTable.csv') -NoTypeInformation -Encoding utf8BOM
    }
    Write-Log "  Log Analytics: $(@($ws | Where-Object { $_.IngestionQueryStatus -like 'OK*' }).Count)/$($ws.Count) workspace(s) queried successfully" Success
}

#endregion

#region ============================== PER-TYPE COLLECTION ==============================

$typeSummary = [System.Collections.Generic.List[object]]::new()
$script:MetricCatalogRows = [System.Collections.Generic.List[object]]::new()
$script:MetricCatalogKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:MetricCallsTotal = 0

function Invoke-PremiumFileShareCollection($storageResources) {
    $premium = @($storageResources | Where-Object { $_.kind -eq 'FileStorage' })
    if ($premium.Count -eq 0) { return }
    Write-Log "  Collecting provisioned file shares for $($premium.Count) FileStorage (premium/provisioned) account(s)..." Info
    $items = foreach ($r in $premium) { @{ Key = ([string]$r.id).ToLower(); Url = "$script:ArmBase$($r.id)/fileServices/default/shares?api-version=2023-05-01"; Method = 'GET'; TokenKey = 'arm' } }
    $lists = Invoke-ParallelRest -Items @($items) -Label 'File share lists'
    $statItems = [System.Collections.Generic.List[object]]::new(); $shareMeta = @{}
    foreach ($r in $premium) {
        $x = $lists[([string]$r.id).ToLower()]
        if (-not $x -or -not $x.Ok) { if ($x) { Add-CollectionError 'FileShares' $r.id $x.Error }; continue }
        foreach ($s in @($x.Data.value | Where-Object { $_ } | Select-Object -First 200)) {
            $k = ([string]$s.id).ToLower(); $shareMeta[$k] = @{ Account = $r; Share = $s }
            $statItems.Add(@{ Key = $k; Url = "$script:ArmBase$($s.id)?api-version=2023-05-01&`$expand=stats"; Method = 'GET'; TokenKey = 'arm' })
        }
    }
    $stats = Invoke-ParallelRest -Items $statItems.ToArray() -Label 'File share usage'
    $rows = foreach ($k in $shareMeta.Keys) {
        $r = $shareMeta[$k].Account; $s = $shareMeta[$k].Share; $st = $stats[$k]
        $p = if ($st -and $st.Ok) { $st.Data.properties } else { $s.properties }
        $quota = [double]($p.shareQuota ?? 0)
        $usedGiB = if ($null -ne $p.shareUsageBytes) { [Math]::Round([double]$p.shareUsageBytes / 1GB, 2) } else { $null }
        [ordered]@{
            SubscriptionName = $script:SubNames[[string]$r.subscriptionId]; SubscriptionId = [string]$r.subscriptionId; ResourceGroup = [string]$r.resourceGroup
            ResourceName = [string]$r.name; ResourceType = 'Provisioned File Shares'; Location = [string]$r.location
            ShareName = [string]$s.name; StorageSku = [string]$r.sku.name
            BillingModel = if ([string]$r.sku.name -match 'V2') { 'Provisioned v2' } else { 'Provisioned v1 (Premium)' }
            AccessTier = [string]$p.accessTier; Protocol = [string]$p.enabledProtocols
            ProvisionedGiB = $quota; UsedGiB = $usedGiB; UsedPct = Get-Ratio $usedGiB $quota
            ProvisionedIops = $p.provisionedIops; ProvisionedBandwidthMiBps = $p.provisionedBandwidthMibps
            LastModified = [string]$p.lastModifiedTime
            MetricsStatus = if ($st -and $st.Ok) { 'OK' } else { "Usage unavailable: $($st.Error)" }
            _Id = [string]$r.id; _Tags = Get-TagString $r.tags
        }
    }
    if ($rows) {
        Write-TypeCsv -Rows @($rows) -Path (Join-Path $TypesPath 'PremiumFileShares.csv')
        $typeSummary.Add([pscustomobject]@{ Key = 'PremiumFileShares'; Display = 'Provisioned File Shares'; Category = 'Storage'; Resources = $premium.Count; Rows = @($rows).Count; MetricsOK = @($rows | Where-Object { $_.MetricsStatus -eq 'OK' }).Count; MetricsNoData = 0; MetricsError = 0; MetricCalls = 0; Generic = $false; ArmTypes = 'microsoft.storage/storageaccounts/fileservices/shares' })
        Write-Log "  Provisioned file shares: $(@($rows).Count) share(s)" Success
    }
}

function Invoke-TypeCollection($cat, [object[]]$resources) {
    Write-Log "[$($cat.Category)] $($cat.Display): $($resources.Count) resource(s)" Section

    if ($cat.PreCollect) { try { & $cat.PreCollect $resources } catch { Add-CollectionError 'PreCollect' $cat.Key $_.Exception.Message; Write-Log "  Pre-collection failed: $($_.Exception.Message)" Warning } }

    # ---- Build rows (provisioned configuration + all properties + all tags)
    $rows = [System.Collections.Generic.List[object]]::new()
    $rowsById = @{}
    $metricItems = [System.Collections.Generic.List[object]]::new()
    $specsById = @{}
    $extraSpecs = @(foreach ($x in @(if ($script:CatalogExtraMetrics -is [System.Collections.IDictionary] -and $script:CatalogExtraMetrics.Contains($cat.Key)) { $script:CatalogExtraMetrics[$cat.Key] })) {
            if ($x -and $x.Name) { M ([string]$x.Name) ([string]($x.Alias ?? ($x.Name -replace '[^A-Za-z0-9]', ''))) -Out ([string]($x.Out ?? 'Avg,P95,Max')) -Filter ([string]($x.Filter ?? '')) -SubPath ([string]($x.SubPath ?? '')) }
        })
    foreach ($r in $resources) {
        $id = ([string]$r.id).ToLower()
        $base = New-BaseRow $r $cat
        try { $props = & $cat.Props $r; foreach ($kv in $props.GetEnumerator()) { $base[$kv.Key] = $kv.Value } }
        catch { Add-CollectionError 'Properties' $r.id $_.Exception.Message }
        try { Add-ResourceDetails $base $r } catch { Add-CollectionError 'PropertyFlatten' $r.id $_.Exception.Message }
        $base._Id = [string]$r.id; $base._Tags = Get-TagString $r.tags
        $resRows = @($base)
        if ($cat.Expand) {
            try { $exp = @(& $cat.Expand $r $base | Where-Object { $_ }); if ($exp.Count -gt 0) { $resRows = $exp } }
            catch { Add-CollectionError 'Expand' $r.id $_.Exception.Message }
        }
        foreach ($rr in $resRows) { $rows.Add($rr) }
        $rowsById[$id] = $resRows
        if (-not $SkipMetrics -and ($cat.Metrics -or $cat.GenericSpecs -or $extraSpecs.Count)) {
            try {
                $specs = @(if ($cat.Metrics) { & $cat.Metrics $r $resRows[0] | Where-Object { $_ } } elseif ($cat.GenericSpecs) { $cat.GenericSpecs }) + $extraSpecs
                foreach ($s in $specs) {
                    $mk = "$($cat.Key)|$($s.Alias)"
                    if ($script:MetricCatalogKeys.Add($mk)) { $script:MetricCatalogRows.Add([pscustomobject][ordered]@{ TypeKey = $cat.Key; ResourceType = $cat.Display; AzureMetric = $s.Name; Alias = $s.Alias; Outputs = $s.Out; Scale = $s.Scale; Basis = $s.Basis; SubResource = $s.SubPath; DimensionFilter = $s.Filter; UsagePattern = [bool]$s.Pattern }) }
                }
                if ($specs.Count -gt 0) { $specsById[$id] = $specs; $metricItems.Add((New-MetricItem $r $specs)) }
            }
            catch { Add-CollectionError 'MetricPlan' $r.id $_.Exception.Message }
        }
    }

    # ---- Collect metrics (parallel)
    $results = @{}
    if ($metricItems.Count -gt 0) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $results = Invoke-MetricCollection -Items $metricItems.ToArray() -Label $cat.Display
        Write-Log "  Metrics collected for $($results.Count) resource(s) in $([int]$sw.Elapsed.TotalSeconds)s" Info
    }

    # ---- Merge statistics, write daily series and usage-pattern accumulators
    $dailyFile = Join-Path $DailyPath "$($cat.Key)_Daily.csv"
    $dailyWriter = $null
    $how = @{}
    $okCount = 0; $noData = 0; $errCount = 0; $calls = 0
    foreach ($id in $rowsById.Keys) {
        $resRows = $rowsById[$id]; $res = $results[$id]; $specs = $specsById[$id]
        if ($res) {
            $calls += [int]$res.Calls
            switch ($res.Status) { 'OK' { $okCount++ } 'Partial' { $okCount++ } 'NoData' { $noData++ } default { $errCount++ } }
            if ($res.Errors) { Add-CollectionError 'Metrics' $resRows[0]._Id $res.Errors }
            # daily rows (one set per resource)
            foreach ($sk in $res.Series.Keys) {
                $st = $res.Series[$sk]; if (-not $st.Daily -or $st.Daily.Count -eq 0) { continue }
                if (-not $dailyWriter) {
                    $dailyWriter = [System.IO.StreamWriter]::new($dailyFile, $false, $script:Utf8Bom)
                    $dailyWriter.WriteLine('SubscriptionName,ResourceGroup,ResourceName,ResourceType,Metric,Dimension,Date,Average,Maximum,Minimum,Total,HoursWithData,ResourceId')
                }
                $parts = $sk -split '\|', 2; $metric = $parts[0]; $dim = if ($parts.Count -gt 1) { $parts[1] } else { '' }
                $r0 = $resRows[0]
                $prefix = (Format-CsvField $r0.SubscriptionName) + ',' + (Format-CsvField $r0.ResourceGroup) + ',' + (Format-CsvField $r0.ResourceName) + ',' + (Format-CsvField $cat.Display) + ',' + (Format-CsvField $metric) + ',' + (Format-CsvField $dim) + ','
                foreach ($dy in $st.Daily) {
                    $dailyWriter.WriteLine($prefix + $DayLabels[$dy[0]] + ',' + (Format-Num $dy[1]) + ',' + (Format-Num $dy[2]) + ',' + (Format-Num $dy[3]) + ',' + (Format-Num $dy[4]) + ',' + $dy[5] + ',' + (Format-CsvField $r0._Id))
                }
            }
        }
        foreach ($row in $resRows) {
            if (-not $specs) {
                $row.MetricsStatus = if ($SkipMetrics) { 'Skipped (-SkipMetrics)' } elseif ($cat.Generic) { 'N/A (no platform metrics - cost only)' } elseif ($cat.Metrics) { 'Not collected (state/SKU not applicable)' } else { 'N/A (configuration-based analysis)' }
                continue
            }
            foreach ($s in $specs) {
                $isSplit = -not [string]::IsNullOrEmpty($s.Filter)
                $key = if ($isSplit) { if ($row._Dim) { "$($s.Alias)|$($row._Dim)" } else { $null } } else { $s.Alias }
                $st = if ($key -and $res) { $res.Series[$key] } else { $null }
                if (-not $isSplit -or $row._Dim) {
                    $metricFailed = $res -and $res.Errors -and (([string]$res.Errors) -match ([regex]::Escape($s.Name) + '(,| ->)'))
                    $callOk = $res -and $res.Status -in @('OK', 'Partial', 'NoData') -and -not $metricFailed
                    foreach ($o in ($s.Out -split ',')) {
                        $row["$($s.Alias)_$o"] = if ($st) { Get-Round $st[$o] 3 } elseif ($callOk -and $o -in @('Total', 'MaxT')) { 0 } else { $null }
                    }
                }
                if ($s.Pattern -and $key) {
                    $pt = if ($st) { $st.Pattern } else { $null }
                    $row.PrimaryMetric = $s.Alias
                    $row.Pattern_ActiveHoursPct = if ($pt) { Get-Round $pt.ActivePct 1 } else { 0 }
                    $row.Pattern_IdleHoursPct = if ($pt) { Get-Round $pt.IdlePct 1 } else { $null }
                    $row.Pattern_BizHoursAvg = if ($pt) { Get-Round $pt.BizAvg 3 } else { $null }
                    $row.Pattern_OffHoursAvg = if ($pt) { Get-Round $pt.OffAvg 3 } else { $null }
                    $row.Pattern_WeekdayAvg = if ($pt) { Get-Round $pt.WeekdayAvg 3 } else { $null }
                    $row.Pattern_WeekendAvg = if ($pt) { Get-Round $pt.WeekendAvg 3 } else { $null }
                    $row.Pattern_PeakHour = if ($pt -and $null -ne $pt.PeakHour) { '{0:00}:00' -f $pt.PeakHour } else { $null }
                    $row.Pattern_PeakDay = if ($pt -and $null -ne $pt.PeakDow) { $script:DayNames[$pt.PeakDow] } else { $null }
                    if ($pt) {
                        if (-not $how.ContainsKey($s.Alias)) { $how[$s.Alias] = @{ S = [double[]]::new(168); C = [int[]]::new(168); R = 0 } }
                        $acc = $how[$s.Alias]; $acc.R++
                        for ($h = 0; $h -lt 168; $h++) { $acc.S[$h] += $pt.HowSum[$h]; $acc.C[$h] += $pt.HowCnt[$h] }
                    }
                }
            }
            if ($cat.Derive) { try { & $cat.Derive $row $(if ($res) { $res.Series } else { @{} }) $cat } catch { Add-CollectionError 'Derive' $row._Id $_.Exception.Message } }
            $row.MetricsStatus = if ($res) { $res.Status } else { 'Error' }
            $row.MetricsErrors = if ($res -and $res.Errors) { ([string]$res.Errors).Substring(0, [Math]::Min(300, ([string]$res.Errors).Length)) } else { $null }
            $row.MetricCalls = if ($res) { $res.Calls } else { 0 }
        }
    }
    if ($dailyWriter) { $dailyWriter.Dispose() }
    $script:MetricCallsTotal += $calls

    # ---- Type-specific post collection
    switch ($cat.Key) {
        'StorageAccounts' { try { Invoke-PremiumFileShareCollection $resources } catch { Add-CollectionError 'FileShares' 'all' $_.Exception.Message } }
        'CognitiveServices' {
            foreach ($d in $script:AiDeploymentRows) {
                $res = $results[([string]$d._Id).ToLower()]
                $st = if ($res -and $d.IsProvisioned) { $res.Series["PTUUtilPct|$($d.Deployment)"] } else { $null }
                $d.PTUUtilPct_Avg = if ($st) { Get-Round $st.Avg 1 } else { $null }
                $d.PTUUtilPct_P95 = if ($st) { Get-Round $st.P95 1 } else { $null }
                $d.PTUUtilPct_Max = if ($st) { Get-Round $st.Max 1 } else { $null }
                $d.MetricsStatus = if ($d.IsProvisioned) { if ($st) { 'OK' } else { 'NoData' } } else { 'N/A (pay-per-token)' }
            }
        }
        'LogAnalyticsWorkspaces' { if (-not $SkipLogAnalytics) { try { Add-LogAnalyticsIngestion $rows } catch { Add-CollectionError 'LogAnalytics' 'all' $_.Exception.Message; Write-Log "  Log Analytics ingestion collection failed: $($_.Exception.Message)" Warning } } }
    }

    # ---- Export per-type CSV + hour-of-week usage pattern
    Write-TypeCsv -Rows $rows -Path (Join-Path $TypesPath "$($cat.Key).csv")
    if ($how.Count -gt 0) {
        $hrows = foreach ($alias in $how.Keys) {
            $acc = $how[$alias]
            for ($h = 0; $h -lt 168; $h++) {
                [pscustomobject]@{
                    ResourceType = $cat.Display; Metric = $alias; DayIndex = [int][Math]::Floor($h / 24); DayOfWeek = $script:DayNames[[int][Math]::Floor($h / 24)]
                    Hour = $h % 24; AvgValue = if ($acc.C[$h]) { [Math]::Round($acc.S[$h] / $acc.C[$h], 4) } else { $null }; Samples = $acc.C[$h]; Resources = $acc.R
                }
            }
        }
        $hrows | Export-Csv -Path (Join-Path $PatternPath "$($cat.Key)_HourOfWeek.csv") -NoTypeInformation -Encoding utf8BOM
    }
    $typeSummary.Add([pscustomobject]@{
            Key = $cat.Key; Display = $cat.Display; Category = $cat.Category; Resources = $resources.Count; Rows = $rows.Count
            MetricsOK = $okCount; MetricsNoData = $noData; MetricsError = $errCount; MetricCalls = $calls; Generic = [bool]$cat.Generic; ArmTypes = ($cat.Types -join ',')
        })
    Write-Log "  Exported $($rows.Count) row(s) -> ResourceTypes\$($cat.Key).csv (metrics OK: $okCount, no data: $noData, errors: $errCount)" Success
}

#endregion

#region ============================== ORPHANED / IDLE RESOURCES ==============================

Write-Log 'Detecting orphaned & unused resources (Resource Graph)...' Section
$orphanChecks = @(
    @{ Finding = 'Unattached managed disk'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.compute/disks' and tostring(properties.diskState) =~ 'Unattached' | extend detail = strcat(tostring(sku.name), ' ', tostring(properties.diskSizeGB), ' GB, last detached ', tostring(properties.LastOwnershipUpdateTime))" }
    @{ Finding = 'Unattached public IP address'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.network/publicipaddresses' and isempty(tostring(properties.ipConfiguration.id)) and isempty(tostring(properties.natGateway.id)) | extend detail = strcat(tostring(sku.name), ' / ', tostring(properties.publicIPAllocationMethod), ' ', tostring(properties.ipAddress))" }
    @{ Finding = 'App Service plan with no apps'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.web/serverfarms' and toint(properties.numberOfSites) == 0 and tostring(sku.tier) !in~ ('Free','Shared','Dynamic','FlexConsumption') | extend detail = strcat(tostring(sku.name), ' x', tostring(sku.capacity), ' instance(s)')" }
    @{ Finding = 'Load balancer with no backend pools'; Cost = 'Paid (Standard)'; Q = "resources | where type =~ 'microsoft.network/loadbalancers' and array_length(properties.backendAddressPools) == 0 | extend detail = strcat(tostring(sku.name), ' SKU, ', tostring(array_length(properties.frontendIPConfigurations)), ' frontend(s)')" }
    @{ Finding = 'NAT gateway not attached to any subnet'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.network/natgateways' and array_length(properties.subnets) == 0 | extend detail = 'No subnets associated'" }
    @{ Finding = 'Virtual network gateway with no connections'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.network/virtualnetworkgateways' | extend gid = tolower(id) | join kind=leftouter (resources | where type =~ 'microsoft.network/connections' | mv-expand gw = pack_array(tolower(tostring(properties.virtualNetworkGateway1.id)), tolower(tostring(properties.virtualNetworkGateway2.id))) | extend gid = tostring(gw) | where isnotempty(gid) | summarize conns = count() by gid) on gid | where isnull(conns) or conns == 0 | extend detail = strcat(tostring(properties.gatewayType), ' ', tostring(properties.sku.name), ' - no site-to-site/ER connections (check P2S usage)')" }
    @{ Finding = 'Private endpoint with disconnected/rejected connection'; Cost = 'Paid'; Q = "resources | where type =~ 'microsoft.network/privateendpoints' | mv-expand c = array_concat(properties.privateLinkServiceConnections, properties.manualPrivateLinkServiceConnections) | extend st = tostring(c.properties.privateLinkServiceConnectionState.status) | where st in~ ('Disconnected','Rejected') | extend detail = strcat('Connection state: ', st, ' -> ', tostring(split(tostring(c.properties.privateLinkServiceId), '/')[-1]))" }
    @{ Finding = 'Private DNS zone without VNet links'; Cost = 'Paid (small)'; Q = "resources | where type =~ 'microsoft.network/privatednszones' and toint(properties.numberOfVirtualNetworkLinks) == 0 | extend detail = strcat(tostring(properties.numberOfRecordSets), ' record set(s), 0 VNet links')" }
    @{ Finding = 'Traffic Manager profile without endpoints'; Cost = 'Paid (small)'; Q = "resources | where type =~ 'microsoft.network/trafficmanagerprofiles' and array_length(properties.endpoints) == 0 | extend detail = 'No endpoints'" }
    @{ Finding = 'DDoS protection plan with no protected VNets'; Cost = 'Paid (high)'; Q = "resources | where type =~ 'microsoft.network/ddosprotectionplans' and array_length(properties.virtualNetworks) == 0 | extend detail = 'Plan billed monthly with no VNets attached'" }
    @{ Finding = 'Unattached network interface'; Cost = 'Free (hygiene)'; Q = "resources | where type =~ 'microsoft.network/networkinterfaces' and isempty(tostring(properties.virtualMachine.id)) and isempty(tostring(properties.privateEndpoint.id)) and isempty(tostring(properties.privateLinkService.id)) | extend detail = 'Not attached to a VM / private endpoint'" }
    @{ Finding = 'Network security group not associated'; Cost = 'Free (hygiene)'; Q = "resources | where type =~ 'microsoft.network/networksecuritygroups' and isnull(properties.networkInterfaces) and isnull(properties.subnets) | extend detail = 'No subnet or NIC association'" }
    @{ Finding = 'Route table not associated'; Cost = 'Free (hygiene)'; Q = "resources | where type =~ 'microsoft.network/routetables' and isnull(properties.subnets) | extend detail = 'No subnet association'" }
    @{ Finding = 'Empty resource group'; Cost = 'Free (hygiene)'; Q = "resourcecontainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups' | extend rgKey = strcat(subscriptionId, '/', tolower(name)) | join kind=leftouter (resources | summarize n = count() by rgKey = strcat(subscriptionId, '/', tolower(resourceGroup))) on rgKey | where isnull(n) or n == 0 | extend resourceGroup = name, detail = 'Resource group contains no resources'" }
)
$orphanRows = [System.Collections.Generic.List[object]]::new()
foreach ($chk in $orphanChecks) {
    try {
        $q = "$($chk.Q) | project id, name, type, location, resourceGroup, subscriptionId, sku = tostring(sku.name), detail"
        foreach ($x in (Invoke-ARGQuery $q)) {
            $orphanRows.Add([pscustomobject][ordered]@{
                    Finding = $chk.Finding; CostImpact = $chk.Cost; SubscriptionName = $script:SubNames[[string]$x.subscriptionId]; ResourceGroup = $x.resourceGroup
                    ResourceName = $x.name; ArmType = $x.type; Location = $x.location; Sku = $x.sku; Detail = $x.detail; ResourceId = $x.id
                })
        }
    }
    catch { Add-CollectionError 'Orphans' $chk.Finding $_.Exception.Message; Write-Log "  Check '$($chk.Finding)' failed: $($_.Exception.Message)" Warning }
}
foreach ($o in $script:OrphanBackupRows) { $orphanRows.Add($o) }
$orphanRows | Export-Csv -Path (Join-Path $RawPath 'Orphaned_Resources.csv') -NoTypeInformation -Encoding utf8BOM
Write-Log "Orphaned/unused findings: $($orphanRows.Count) ($(@($orphanRows | Where-Object { $_.CostImpact -like 'Paid*' }).Count) cost-bearing)" Success

#endregion

#region ============================== ADVISOR & DEFENDER ==============================

Write-Log 'Collecting Azure Advisor Cost & Performance recommendations...' Section
$advisorStatus = 'OK'
try {
    $adv = Invoke-ARGQuery @"
advisorresources
| where type =~ 'microsoft.advisor/recommendations' and tostring(properties.category) in~ ('Cost', 'Performance')
| extend ep = properties.extendedProperties
| project subscriptionId, resourceGroup, category = tostring(properties.category), impactedType = tostring(properties.impactedField), impactedName = tostring(properties.impactedValue),
          impact = tostring(properties.impact), problem = tostring(properties.shortDescription.problem), solution = tostring(properties.shortDescription.solution),
          resourceId = tostring(properties.resourceMetadata.resourceId), lastUpdated = tostring(properties.lastUpdated), recommendationTypeId = tostring(properties.recommendationTypeId),
          annualSavings = todouble(ep.annualSavingsAmount), monthlySavings = todouble(ep.savingsAmount), currency = tostring(ep.savingsCurrency),
          currentSku = tostring(ep.currentSku), targetSku = tostring(ep.targetSku), term = tostring(ep.term), lookback = tostring(ep.lookbackPeriod),
          region = tostring(ep.region), details = tostring(ep)
"@
    $adv | ForEach-Object {
        [pscustomobject][ordered]@{
            Category = $_.category; SubscriptionName = $script:SubNames[[string]$_.subscriptionId]; ResourceGroup = $_.resourceGroup; ImpactedType = $_.impactedType; ImpactedResource = $_.impactedName
            Impact = $_.impact; Problem = $_.problem; Solution = $_.solution; CurrentSku = $_.currentSku; TargetSku = $_.targetSku; Term = $_.term
            LookbackDays = $_.lookback; MonthlySavings = $_.monthlySavings; AnnualSavings = $_.annualSavings; Currency = $_.currency
            LastUpdated = $_.lastUpdated; RecommendationTypeId = $_.recommendationTypeId; ResourceId = $_.resourceId; Details = $_.details
        }
    } | Export-Csv -Path (Join-Path $RawPath 'Advisor_CostRecommendations.csv') -NoTypeInformation -Encoding utf8BOM
    Write-Log "Advisor recommendations: $($adv.Count) ($(@($adv | Where-Object category -eq 'Cost').Count) cost, $(@($adv | Where-Object category -eq 'Performance').Count) performance)" Success
}
catch { $advisorStatus = "Failed: $($_.Exception.Message)"; Add-CollectionError 'Advisor' 'all' $_.Exception.Message; Write-Log "Advisor query failed: $($_.Exception.Message)" Warning }

try {
    $def = Invoke-ARGQuery "securityresources | where type =~ 'microsoft.security/pricings' | project subscriptionId, plan = name, tier = tostring(properties.pricingTier), subPlan = tostring(properties.subPlan), enablementTime = tostring(properties.enablementTime), deprecated = tostring(properties.deprecated)"
    $def | ForEach-Object { [pscustomobject][ordered]@{ SubscriptionName = $script:SubNames[[string]$_.subscriptionId]; SubscriptionId = $_.subscriptionId; Plan = $_.plan; Tier = $_.tier; SubPlan = $_.subPlan; EnablementTime = $_.enablementTime; Deprecated = $_.deprecated } } |
        Export-Csv -Path (Join-Path $RawPath 'Defender_Plans.csv') -NoTypeInformation -Encoding utf8BOM
    Write-Log "Defender for Cloud plan settings: $($def.Count)" Success
}
catch { Add-CollectionError 'Defender' 'all' $_.Exception.Message }

#endregion

#region ============================== COST MANAGEMENT ==============================

$costStatus = 'Skipped'
$script:CostById = @{}
if (-not $SkipCostData) {
    Write-Log "Collecting $CostType per resource and daily cost per service (Cost Management)..." Section
    $fromDate = $StartLocal.ToString('yyyy-MM-dd'); $toDate = $EndLocal.AddDays(-1).ToString('yyyy-MM-dd')
    function Invoke-CostQuery([string]$SubId, [string]$Granularity, [object[]]$Grouping, [string]$From, [string]$To) {
        if (-not $From) { $From = $fromDate }; if (-not $To) { $To = $toDate }
        $variants = @(
            @{ totalCost = @{ name = 'Cost'; function = 'Sum' }; totalCostUSD = @{ name = 'CostUSD'; function = 'Sum' } },
            @{ totalCost = @{ name = 'Cost'; function = 'Sum' } },
            @{ totalCost = @{ name = 'PreTaxCost'; function = 'Sum' } })
        $lastErr = $null
        foreach ($agg in $variants) {
            $body = @{ type = $CostType; timeframe = 'Custom'; timePeriod = @{ from = "${From}T00:00:00Z"; to = "${To}T23:59:59Z" }
                dataset = @{ granularity = $Granularity; aggregation = $agg; grouping = $Grouping } }
            try {
                $rows = [System.Collections.Generic.List[object]]::new(); $cols = $null; $page = 0
                $uri = "$script:ArmBase/subscriptions/$SubId/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
                while ($uri -and $page -lt 100) {
                    $page++
                    $resp = Invoke-AzApi -Uri $uri -Method POST -Body $body -MaxRetries 8
                    if (-not $cols) { $cols = @($resp.properties.columns | ForEach-Object { [string]$_.name }) }
                    foreach ($rw in @($resp.properties.rows)) { $o = @{}; for ($i = 0; $i -lt $cols.Count; $i++) { $o[$cols[$i]] = $rw[$i] }; $rows.Add($o) }
                    $uri = $resp.properties.nextLink
                }
                return @{ Ok = $true; Rows = $rows }
            }
            catch { $lastErr = $_.Exception.Message; if ($lastErr -notmatch 'HTTP 400') { break } }
        }
        return @{ Ok = $false; Error = $lastErr }
    }
    $costRows = [System.Collections.Generic.List[object]]::new(); $dailyCostRows = [System.Collections.Generic.List[object]]::new(); $monthlyCostRows = [System.Collections.Generic.List[object]]::new()
    $okSubs = 0; $failSubs = [System.Collections.Generic.List[string]]::new()
    foreach ($sub in $allSubs) {
        $r1 = Invoke-CostQuery $sub.Id 'None' @(@{ type = 'Dimension'; name = 'ResourceId' }, @{ type = 'Dimension'; name = 'ServiceName' })
        if ($r1.Ok) {
            $okSubs++
            foreach ($x in $r1.Rows) {
                $rid = ([string]$x.ResourceId).ToLower()
                $costRows.Add([pscustomobject][ordered]@{
                        SubscriptionName = $sub.Name; SubscriptionId = $sub.Id
                        ResourceGroup = if ($rid -match '/resourcegroups/([^/]+)') { $Matches[1] } else { '' }
                        ResourceName = Get-NameFromId $rid; ServiceName = $x.ServiceName
                        Cost = [Math]::Round([double]($x.Cost ?? $x.PreTaxCost ?? 0), 4); CostUSD = if ($null -ne $x.CostUSD) { [Math]::Round([double]$x.CostUSD, 4) } else { $null }
                        Currency = $x.Currency; ResourceId = $rid
                    })
            }
        }
        else { $failSubs.Add("$($sub.Name): $($r1.Error)"); Add-CollectionError 'CostByResource' $sub.Name $r1.Error }
        $r2 = Invoke-CostQuery $sub.Id 'Daily' @(@{ type = 'Dimension'; name = 'ServiceName' })
        if ($r2.Ok) {
            foreach ($x in $r2.Rows) {
                $ds = [string]$x.UsageDate
                $dateLabel = if ($ds -match '^(\d{4})(\d{2})(\d{2})') { "$($Matches[1])-$($Matches[2])-$($Matches[3])" } else { $ds }
                $dailyCostRows.Add([pscustomobject][ordered]@{
                        Date = $dateLabel; SubscriptionName = $sub.Name; ServiceName = $x.ServiceName
                        Cost = [Math]::Round([double]($x.Cost ?? $x.PreTaxCost ?? 0), 4); CostUSD = if ($null -ne $x.CostUSD) { [Math]::Round([double]$x.CostUSD, 4) } else { $null }
                        Currency = $x.Currency
                    })
            }
        }
        else { Add-CollectionError 'CostDaily' $sub.Name $r2.Error }
        if ($CostHistoryMonths -gt 0) {
            $histFrom = (Get-Date -Year $StartLocal.Year -Month $StartLocal.Month -Day 1).AddMonths(-$CostHistoryMonths).ToString('yyyy-MM-dd')
            $r3 = Invoke-CostQuery $sub.Id 'Monthly' @(@{ type = 'Dimension'; name = 'ServiceName' }) $histFrom $toDate
            if ($r3.Ok) {
                foreach ($x in $r3.Rows) {
                    $bm = $x.BillingMonth ?? $x.UsageDate
                    $month = if ($bm -is [DateTime]) { $bm.ToString('yyyy-MM') } elseif ([string]$bm -match '^(\d{4})-?(\d{2})') { "$($Matches[1])-$($Matches[2])" } else { [string]$bm }
                    $monthlyCostRows.Add([pscustomobject][ordered]@{
                            Month = $month; SubscriptionName = $sub.Name; ServiceName = $x.ServiceName
                            Cost = [Math]::Round([double]($x.Cost ?? $x.PreTaxCost ?? 0), 2); CostUSD = if ($null -ne $x.CostUSD) { [Math]::Round([double]$x.CostUSD, 2) } else { $null }
                            Currency = $x.Currency; PartialMonth = ($month -eq $EndLocal.AddDays(-1).ToString('yyyy-MM') -and $EndLocal.Day -ne 1)
                        })
                }
            }
            else { Add-CollectionError 'CostMonthly' $sub.Name $r3.Error }
        }
        Start-Sleep -Milliseconds 500
    }
    foreach ($c in $costRows) { $script:CostById[$c.ResourceId] = [double]($script:CostById[$c.ResourceId] ?? 0) + [double]($c.CostUSD ?? $c.Cost ?? 0) }
    if ($costRows.Count) { $costRows | Export-Csv -Path (Join-Path $RawPath 'Cost_ByResource.csv') -NoTypeInformation -Encoding utf8BOM }
    if ($dailyCostRows.Count) { $dailyCostRows | Export-Csv -Path (Join-Path $RawPath 'Cost_DailyByService.csv') -NoTypeInformation -Encoding utf8BOM }
    if ($monthlyCostRows.Count) { $monthlyCostRows | Export-Csv -Path (Join-Path $RawPath 'Cost_MonthlyTrend.csv') -NoTypeInformation -Encoding utf8BOM }
    $costStatus = if ($okSubs -eq $allSubs.Count) { 'OK' } elseif ($okSubs -gt 0) { "Partial ($okSubs/$($allSubs.Count) subscriptions)" } else { "Unavailable: $(@($failSubs)[0])" }
    Write-Log "Cost data: $costStatus - $($costRows.Count) resource/service rows, $($dailyCostRows.Count) daily rows, $($monthlyCostRows.Count) monthly history rows" $(if ($okSubs) { 'Success' } else { 'Warning' })
}

#endregion

#region ============================== CATALOG COLLECTION ==============================

Write-Log 'Collecting configuration, properties, tags & utilization per resource type...' Section
$script:CatalogTypes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($c in $script:Catalog) { foreach ($t in $c.Types) { [void]$script:CatalogTypes.Add($t) } }
foreach ($cat in $script:Catalog) {
    if ($IncludeResourceTypes -and $cat.Key -notin $IncludeResourceTypes) { continue }
    $typeList = ($cat.Types | ForEach-Object { "'$_'" }) -join ','
    $q = "resources | where type in~ ($typeList) $($cat.Where) | order by id asc"
    try { $resources = Invoke-ARGQuery $q }
    catch { Add-CollectionError 'ARG' $cat.Key $_.Exception.Message; Write-Log "[$($cat.Key)] Resource Graph query failed: $($_.Exception.Message)" Error; continue }
    if ($resources.Count -eq 0) { continue }
    try { Invoke-TypeCollection $cat $resources }
    catch { Add-CollectionError 'TypeCollection' $cat.Key $_.Exception.Message; Write-Log "[$($cat.Key)] collection failed: $($_.Exception.Message)" Error }
}

if ($script:AiDeploymentRows.Count -gt 0) {
    Write-TypeCsv -Rows $script:AiDeploymentRows -Path (Join-Path $TypesPath 'AIModelDeployments.csv')
    $typeSummary.Add([pscustomobject]@{ Key = 'AIModelDeployments'; Display = 'AI Model Deployments'; Category = 'Analytics & AI'; Resources = @($script:AiDeploymentRows | ForEach-Object { $_._Id } | Select-Object -Unique).Count; Rows = $script:AiDeploymentRows.Count; MetricsOK = @($script:AiDeploymentRows | Where-Object { $_.MetricsStatus -eq 'OK' }).Count; MetricsNoData = 0; MetricsError = 0; MetricCalls = 0; Generic = $false })
}
if ($script:MlComputeRows.Count -gt 0) {
    Write-TypeCsv -Rows $script:MlComputeRows -Path (Join-Path $TypesPath 'MLComputes.csv')
    $typeSummary.Add([pscustomobject]@{ Key = 'MLComputes'; Display = 'ML Computes'; Category = 'Analytics & AI'; Resources = @($script:MlComputeRows | ForEach-Object { $_._Id } | Select-Object -Unique).Count; Rows = $script:MlComputeRows.Count; MetricsOK = 0; MetricsNoData = 0; MetricsError = 0; MetricCalls = 0; Generic = $false })
}

#endregion

#region ============================== GENERIC COVERAGE (EVERY OTHER BILLED RESOURCE TYPE) ==============================

$GenericDisplay = @{
    'microsoft.keyvault/vaults' = 'Key Vaults'; 'microsoft.signalrservice/signalr' = 'SignalR Service'; 'microsoft.signalrservice/webpubsub' = 'Web PubSub'
    'microsoft.eventgrid/topics' = 'Event Grid Topics'; 'microsoft.eventgrid/systemtopics' = 'Event Grid System Topics'; 'microsoft.eventgrid/namespaces' = 'Event Grid Namespaces'; 'microsoft.eventgrid/domains' = 'Event Grid Domains'
    'microsoft.automation/automationaccounts' = 'Automation Accounts'; 'microsoft.network/dnszones' = 'DNS Zones'; 'microsoft.network/privatednszones' = 'Private DNS Zones'
    'microsoft.network/trafficmanagerprofiles' = 'Traffic Manager Profiles'; 'microsoft.network/virtualnetworks' = 'Virtual Networks'; 'microsoft.network/publicipprefixes' = 'Public IP Prefixes'
    'microsoft.network/firewallpolicies' = 'Firewall Policies'; 'microsoft.network/vpngateways' = 'vWAN VPN Gateways'; 'microsoft.network/expressroutegateways' = 'vWAN ExpressRoute Gateways'
    'microsoft.network/p2svpngateways' = 'vWAN Point-to-Site Gateways'; 'microsoft.network/virtualwans' = 'Virtual WANs'; 'microsoft.network/privatelinkservices' = 'Private Link Services'
    'microsoft.network/connections' = 'VPN / ExpressRoute Connections'; 'microsoft.network/localnetworkgateways' = 'Local Network Gateways'; 'microsoft.network/networkwatchers/flowlogs' = 'Flow Logs'
    'microsoft.network/frontdoorwebapplicationfirewallpolicies' = 'Front Door WAF Policies'; 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies' = 'Application Gateway WAF Policies'
    'microsoft.network/expressrouteports' = 'ExpressRoute Direct Ports'; 'microsoft.network/networkmanagers' = 'Virtual Network Managers'
    'microsoft.appconfiguration/configurationstores' = 'App Configuration Stores'; 'microsoft.relay/namespaces' = 'Relay Namespaces'
    'microsoft.notificationhubs/namespaces' = 'Notification Hub Namespaces'; 'microsoft.notificationhubs/namespaces/notificationhubs' = 'Notification Hubs'
    'microsoft.communication/communicationservices' = 'Communication Services'; 'microsoft.maps/accounts' = 'Azure Maps Accounts'
    'microsoft.avs/privateclouds' = 'Azure VMware Solution Private Clouds'; 'microsoft.compute/hostgroups' = 'Dedicated Host Groups'
    'microsoft.compute/capacityreservationgroups' = 'Capacity Reservation Groups'; 'microsoft.compute/galleries/images/versions' = 'Gallery Image Versions'
    'microsoft.compute/images' = 'Managed Images'; 'microsoft.compute/restorepointcollections' = 'Restore Point Collections'; 'microsoft.compute/galleries' = 'Compute Galleries'
    'microsoft.web/staticsites' = 'Static Web Apps'; 'microsoft.web/hostingenvironments' = 'App Service Environments'; 'microsoft.web/sites/slots' = 'App Service Deployment Slots'
    'microsoft.appplatform/spring' = 'Azure Spring Apps'; 'microsoft.servicefabric/clusters' = 'Service Fabric Clusters'; 'microsoft.servicefabric/managedclusters' = 'Service Fabric Managed Clusters'
    'microsoft.redhatopenshift/openshiftclusters' = 'Azure Red Hat OpenShift'; 'microsoft.kubernetes/connectedclusters' = 'Arc-enabled Kubernetes'
    'microsoft.hybridcompute/machines' = 'Arc-enabled Servers'; 'microsoft.azurearcdata/sqlserverinstances' = 'Arc-enabled SQL Server'
    'microsoft.dataprotection/backupvaults' = 'Backup Vaults'; 'microsoft.storagesync/storagesyncservices' = 'Storage Sync Services'
    'microsoft.storagecache/caches' = 'HPC Cache'; 'microsoft.storagecache/amlfilesystems' = 'Managed Lustre'
    'microsoft.netapp/netappaccounts/capacitypools/volumes' = 'NetApp Volumes'; 'microsoft.dbforpostgresql/servergroupsv2' = 'Cosmos DB for PostgreSQL'
    'microsoft.documentdb/cassandraclusters' = 'Managed Apache Cassandra'; 'microsoft.documentdb/mongoclusters' = 'Cosmos DB for MongoDB (vCore)'
    'microsoft.purview/accounts' = 'Microsoft Purview'; 'microsoft.dashboard/grafana' = 'Managed Grafana'; 'microsoft.monitor/accounts' = 'Azure Monitor Workspaces'
    'microsoft.insights/webtests' = 'Availability Tests'; 'microsoft.insights/metricalerts' = 'Metric Alert Rules'; 'microsoft.insights/scheduledqueryrules' = 'Log Search Alert Rules'
    'microsoft.insights/actiongroups' = 'Action Groups'; 'microsoft.insights/activitylogalerts' = 'Activity Log Alerts'
    'microsoft.machinelearningservices/workspaces/onlineendpoints' = 'ML Online Endpoints'; 'microsoft.machinelearningservices/workspaces/onlineendpoints/deployments' = 'ML Online Deployments'
    'microsoft.machinelearningservices/workspaces/batchendpoints' = 'ML Batch Endpoints'; 'microsoft.botservice/botservices' = 'Bot Services'
    'microsoft.devcenter/devcenters' = 'Dev Centers'; 'microsoft.devcenter/projects/pools' = 'Dev Box Pools'; 'microsoft.labservices/labs' = 'Lab Services Labs'; 'microsoft.devtestlab/labs' = 'DevTest Labs'
    'microsoft.loadtestservice/loadtests' = 'Load Testing'; 'microsoft.digitaltwins/digitaltwinsinstances' = 'Digital Twins'; 'microsoft.iotcentral/iotapps' = 'IoT Central Apps'
    'microsoft.devices/provisioningservices' = 'IoT Hub Device Provisioning'; 'microsoft.media/mediaservices' = 'Media Services'; 'microsoft.datashare/accounts' = 'Data Share'
    'microsoft.synapse/workspaces' = 'Synapse Workspaces'; 'microsoft.synapse/workspaces/kustopools' = 'Synapse Data Explorer Pools'; 'microsoft.healthcareapis/workspaces' = 'Health Data Services'
    'microsoft.logic/integrationaccounts' = 'Integration Accounts'; 'microsoft.apimanagement/gateways' = 'API Management Gateways'; 'microsoft.cdn/profiles/afdendpoints' = 'Front Door Endpoints'
    'microsoft.kusto/clusters/databases' = 'Data Explorer Databases'; 'microsoft.securitydevops/azuredevopsconnectors' = 'Defender DevOps Connectors'
    'microsoft.containerinstance/containergroupprofiles' = 'Container Group Profiles'; 'microsoft.app/jobs' = 'Container Apps Jobs'; 'microsoft.app/sessionpools' = 'Container Apps Session Pools'
    'microsoft.chaos/experiments' = 'Chaos Experiments'; 'microsoft.cognitiveservices/accounts/projects' = 'AI Foundry Projects'
}
$ProviderCategory = [ordered]@{
    '^microsoft\.(compute|avs|desktopvirtualization|batch|servicefabric|hybridcompute|devcenter|labservices|devtestlab|vmware)' = 'Compute'
    '^microsoft\.(containerservice|containerinstance|containerregistry|app|redhatopenshift|kubernetes|appplatform)' = 'Containers'
    '^microsoft\.(web|certificateregistration|domainregistration)' = 'App Service'
    '^microsoft\.(sql|dbfor|documentdb|cache|azurearcdata)' = 'Databases'
    '^microsoft\.(storage|netapp|elasticsan|storagecache|storagesync|dataprotection|recoveryservices|classicstorage|datalakestore)' = 'Storage'
    '^microsoft\.(network|cdn)' = 'Networking'
    '^microsoft\.(operationalinsights|insights|operationsmanagement|monitor|dashboard|security|securityinsights|keyvault|alertsmanagement|automation|attestation)' = 'Monitoring & Security'
    '^microsoft\.(eventhub|servicebus|logic|apimanagement|devices|eventgrid|relay|notificationhubs|signalrservice|streamanalytics|communication|iotcentral|digitaltwins|appconfiguration|maps)' = 'Messaging & Integration'
    '^microsoft\.(synapse|datafactory|kusto|databricks|hdinsight|analysisservices|powerbidedicated|fabric|cognitiveservices|search|machinelearningservices|purview|botservice|datashare|media|healthcareapis)' = 'Analytics & AI'
}
function Get-GenericDisplay([string]$Type) {
    $t = $Type.ToLower()
    if ($GenericDisplay.ContainsKey($t)) { return $GenericDisplay[$t] }
    $parts = $t -split '/'
    $prov = ($parts[0] -replace '^microsoft\.', '')
    $leaf = ($parts[1..($parts.Count - 1)] -join ' / ')
    $cap = { param($s) if ($s) { $s.Substring(0, 1).ToUpper() + $s.Substring(1) } else { $s } }
    return "$(& $cap $prov) - $(& $cap $leaf)"
}
function Get-GenericCategory([string]$Type) {
    foreach ($k in $ProviderCategory.Keys) { if ($Type -match $k) { return $ProviderCategory[$k] } }
    return 'Other Services'
}
function Select-GenericMetrics([string]$Type, [object[]]$Defs) {
    if ($script:MetricOverrides.ContainsKey($Type.ToLower()) -and @($script:MetricOverrides[$Type.ToLower()] | Where-Object { $_ }).Count -eq 0) { return @() }
    $override = @($script:MetricOverrides[$Type.ToLower()] | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLower() })
    $cands = foreach ($d in $Defs) {
        $name = [string]$d.name.value; $disp = [string]$d.name.localizedValue; $unit = [string]$d.unit; $prim = [string]$d.primaryAggregationType
        if (-not $name -or $d.isDimensionRequired -eq $true) { continue }
        if ("$name $disp" -match '(?i)deprecated|ddos|pingmesh') { continue }
        $text = "$name $disp"
        $isErr = $text -match '(?i)error|fail|throttl|5xx|4xx|dropped|reject|deadletter|abandon'
        # Percent-of-service-limit metrics (quota / saturation) measure headroom, not utilization of a billed SKU
        $isLimit = $unit -eq 'Percent' -and $name -match '(?i)capacityutilization|percentutilization|saturation|shoebox|limit|quota'
        # Configuration gauges (record / link / object counts, limits) are not usage activity
        $isGauge = $name -match '(?i)(limit|quota)$' -or ($name -match '(?i)(count|size|capacity)$' -and $name -notmatch '(?i)request|message|call|transaction|event|quer|connection|session|hit|job|run|operation|invocation|deliver|publish|api|incoming|outgoing')
        # Latency / duration metrics measure performance, never volume of activity
        $isTime = $unit -match '(?i)^(milliseconds|seconds)$' -or $text -match '(?i)latency|duration|response ?time'
        $isUtil = -not $isLimit -and $unit -eq 'Percent' -and $text -match '(?i)cpu|memory|utiliz|load|capacity|consum|usage|saturation|dtu|disk used|percentage'
        $isAct = -not $isGauge -and -not $isTime -and $text -match '(?i)request|transaction|call|message|connection|quer|execution|run|event|operation|session|hit|job|invocation|deliver|publish|incoming|outgoing|api|count'
        $isTraffic = -not $isTime -and $text -match '(?i)byte|throughput|bandwidth|ingress|egress|traffic|data'
        $isPerf = $text -match '(?i)latency|duration|time'
        $isAvail = $text -match '(?i)availability|health|success rate|hit rate'
        $score = 0
        if ($isUtil -and -not $isErr) { $score += 150 } elseif ($isLimit -and -not $isErr) { $score += 110 } elseif ($unit -eq 'Percent') { $score += 25 }
        if ($isAct -and -not $isErr) { $score += 60 }
        if ($isTraffic -and -not $isErr) { $score += 40 }
        if ($isPerf) { $score += 20 }
        if ($isErr) { $score += 15 }
        if ($isAvail) { $score += 5 }
        if ($isGauge -and -not $isLimit) { $score = [Math]::Min($score, 10) }
        $role = if ($isUtil -and -not $isErr) { 'Utilization' } elseif ($isLimit -and -not $isErr) { 'Limit' } elseif ($isErr) { 'Errors' } elseif ($isGauge) { 'Inventory' } elseif ($isAct -or $isTraffic) { 'Activity' } elseif ($isPerf) { 'Performance' } elseif ($isAvail) { 'Availability' } else { 'Other' }
        if ($override.Count) { $idx = [array]::IndexOf($override, $name.ToLower()); $score = if ($idx -ge 0) { 1000 - $idx } else { -1 } }
        [pscustomobject]@{ Name = $name; Display = $disp; Unit = $unit; Primary = $prim; Score = $score; Role = $role }
    }
    $cands = @($cands | Where-Object { $_.Score -gt 0 } | Sort-Object Score -Descending)
    $sel = @($cands | Select-Object -First $script:GenericMaxMetrics)
    if (-not ($sel | Where-Object Role -eq 'Activity')) {
        $act = $cands | Where-Object { $_.Role -eq 'Activity' -and $_.Name -notin @($sel | ForEach-Object Name) } | Select-Object -First 1
        if ($act) { $sel = @($sel | Select-Object -First ([Math]::Max(0, $script:GenericMaxMetrics - 1))) + $act }
    }
    return $sel
}

$genericSelection = [System.Collections.Generic.List[object]]::new()
$genericTypeCount = 0
if ($script:GenericEnabled -and -not $IncludeResourceTypes) {
    Write-Log 'Generic coverage: discovering utilization metrics for every other billed resource type...' Section
    $monthFactor1 = 30.4375 / [Math]::Max(1, $LookbackDays)
    $candidates = foreach ($g in ($inventory | Group-Object { ([string]$_.type).ToLower() })) {
        $t = $g.Name
        if ($script:CatalogTypes.Contains($t) -or $script:ExcludedTypes.Contains($t)) { continue }
        $cost = 0.0; foreach ($x in $g.Group) { $cost += [double]($script:CostById[([string]$x.id).ToLower()] ?? 0) }
        [pscustomobject]@{ Type = $t; Ids = @($g.Group | ForEach-Object { [string]$_.id }); MonthlyCost = [Math]::Round($cost * $monthFactor1, 2) }
    }
    $candidates = @($candidates)
    $defItems = foreach ($c in $candidates) {
        $i = 0
        foreach ($id in ($c.Ids | Select-Object -First $script:GenericSamples)) { @{ Key = "$($c.Type)|$i"; Url = "$script:ArmBase$id/providers/Microsoft.Insights/metricDefinitions?api-version=2023-10-01"; Method = 'GET'; TokenKey = 'arm' }; $i++ }
    }
    $defRes = if (-not $SkipMetrics -and @($defItems).Count) { Invoke-ParallelRest -Items @($defItems) -Label 'Metric definitions (generic types)' } else { @{} }
    $selectedTypes = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $candidates) {
        $defs = @{}
        for ($i = 0; $i -lt $script:GenericSamples; $i++) {
            $x = $defRes["$($c.Type)|$i"]
            if ($x -and $x.Ok) { foreach ($d in @($x.Data.value)) { if ($d -and -not $defs.ContainsKey([string]$d.name.value)) { $defs[[string]$d.name.value] = $d } } }
        }
        $sel = @(if ($defs.Count) { Select-GenericMetrics $c.Type @($defs.Values) })
        if ($sel.Count -eq 0 -and $c.MonthlyCost -lt $script:GenericMinCost) { continue }
        $selectedTypes.Add([pscustomobject]@{ Type = $c.Type; Selection = $sel; MonthlyCost = $c.MonthlyCost; DefinitionCount = $defs.Count })
    }
    Write-Log "  Generic types selected: $($selectedTypes.Count) of $($candidates.Count) uncovered types (with platform metrics or cost >= $($script:GenericMinCost)/month)" Info
    if ($selectedTypes.Count) {
        $typeList = ($selectedTypes | ForEach-Object { "'$($_.Type)'" }) -join ','
        $allGeneric = @()
        try { $allGeneric = Invoke-ARGQuery "resources | where type in~ ($typeList) | order by id asc" }
        catch { Add-CollectionError 'ARG' 'GenericTypes' $_.Exception.Message }
        $byType = @{}; foreach ($r in $allGeneric) { $k = ([string]$r.type).ToLower(); if (-not $byType.ContainsKey($k)) { $byType[$k] = [System.Collections.Generic.List[object]]::new() }; $byType[$k].Add($r) }
        foreach ($st in $selectedTypes) {
            $res = $byType[$st.Type]; if (-not $res -or $res.Count -eq 0) { continue }
            $specs = [System.Collections.Generic.List[object]]::new(); $roles = @{}; $patternSet = $false
            $usedAliases = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($m in $st.Selection) {
                $alias = (($m.Name -replace '[^A-Za-z0-9]+', ' ').Trim() -split ' ' | ForEach-Object { if ($_) { $_.Substring(0, 1).ToUpper() + $_.Substring(1) } }) -join ''
                if (-not $alias) { $alias = 'Metric' + $specs.Count }
                while (-not $usedAliases.Add($alias)) { $alias = "$alias" + '_' }
                $isTotal = $m.Primary -in @('Total', 'Count')
                $pattern = $false
                if (-not $patternSet -and ($m.Role -eq 'Utilization' -or ($m.Role -eq 'Activity' -and -not ($st.Selection | Where-Object Role -eq 'Utilization')))) { $pattern = $true; $patternSet = $true }
                $spec = if ($isTotal) { M $m.Name $alias -Out 'Total,MaxT' -Basis 'Total' -PatternAgg 'Total' -Idle 0 -Pattern:$pattern }
                else { M $m.Name $alias -Out 'Avg,P95,Max' -Idle $(if ($m.Role -eq 'Utilization') { 5 } else { 0 }) -Pattern:$pattern }
                $specs.Add($spec); $roles[$alias] = $m.Role
                $genericSelection.Add([pscustomobject]@{ ArmType = $st.Type; Display = Get-GenericDisplay $st.Type; Category = Get-GenericCategory $st.Type; Metric = $m.Name; DisplayName = $m.Display; Alias = $alias; Unit = $m.Unit; PrimaryAggregation = $m.Primary; Role = $m.Role; Score = $m.Score })
            }
            if ($st.Selection.Count -eq 0) { $genericSelection.Add([pscustomobject]@{ ArmType = $st.Type; Display = Get-GenericDisplay $st.Type; Category = Get-GenericCategory $st.Type; Metric = ''; DisplayName = '(no platform metrics - cost only)'; Alias = ''; Unit = ''; PrimaryAggregation = ''; Role = 'CostOnly'; Score = 0 }) }
            $key = 'Gen_' + (($st.Type -replace '^microsoft\.', '') -replace '[^A-Za-z0-9]+', '_')
            $gcat = @{
                Key = $key; Display = Get-GenericDisplay $st.Type; Category = Get-GenericCategory $st.Type; Types = @($st.Type); Generic = $true
                GenericSpecs = $specs.ToArray(); GenericRoles = $roles
                Props = { param($r) [ordered]@{} }
                Derive = { param($row, $series, $cat)
                    $util = $null; $utilAlias = $null
                    foreach ($s in $cat.GenericSpecs) {
                        if ($cat.GenericRoles[$s.Alias] -ne 'Utilization') { continue }
                        $v = $row["$($s.Alias)_P95"]
                        if ($null -ne $v -and ($null -eq $util -or [double]$v -gt $util)) { $util = [double]$v; $utilAlias = $s.Alias }
                    }
                    $row.GenericUtilMetric = $utilAlias
                    $row.GenericUtilAvg = if ($utilAlias) { $row["$($utilAlias)_Avg"] } else { $null }
                    $row.GenericUtilP95 = $util
                    $row.GenericUtilMax = if ($utilAlias) { $row["$($utilAlias)_Max"] } else { $null }
                    $lim = $null; $limAlias = $null
                    foreach ($s in $cat.GenericSpecs) {
                        if ($cat.GenericRoles[$s.Alias] -ne 'Limit') { continue }
                        $v = $row["$($s.Alias)_P95"]
                        if ($null -ne $v -and ($null -eq $lim -or [double]$v -gt $lim)) { $lim = [double]$v; $limAlias = $s.Alias }
                    }
                    $row.GenericLimitMetric = $limAlias
                    $row.GenericLimitP95 = $lim
                    $act = 0.0; $actN = 0; $names = [System.Collections.Generic.List[string]]::new()
                    foreach ($s in $cat.GenericSpecs) {
                        if ($cat.GenericRoles[$s.Alias] -ne 'Activity') { continue }
                        $v = if ($s.Out -like '*Total*') { $row["$($s.Alias)_Total"] } else { $row["$($s.Alias)_Max"] }
                        if ($null -ne $v) { $act += [double]$v; $actN++; $names.Add($s.Alias) }
                    }
                    $row.GenericActivityTotal = if ($actN) { [Math]::Round($act, 3) } else { $null }
                    $row.GenericActivityMetrics = ($names -join ', ')
                    $row.GenericMetricCount = @($cat.GenericSpecs).Count
                }
            }
            if ($specs.Count -eq 0) { $gcat.Metrics = $null }
            try { Invoke-TypeCollection $gcat $res.ToArray(); $genericTypeCount++ }
            catch { Add-CollectionError 'GenericType' $st.Type $_.Exception.Message; Write-Log "  Generic type $($st.Type) failed: $($_.Exception.Message)" Warning }
        }
    }
    if ($genericSelection.Count) { $genericSelection | Export-Csv -Path (Join-Path $RawPath 'Reference_GenericMetrics.csv') -NoTypeInformation -Encoding utf8BOM }
    Write-Log "Generic coverage complete: $genericTypeCount additional resource type(s)" Success
}

#endregion

#region ============================== RESERVATIONS & SAVINGS PLANS ==============================

Write-Log 'Collecting reservation & savings plan utilization (best effort - requires billing benefit read access)...' Section
function Get-UtilAgg($util, [int]$Grain) {
    $a = @($util.aggregates | Where-Object { [int]$_.grain -eq $Grain }) | Select-Object -First 1
    if ($a) { return [Math]::Round([double]$a.value, 1) } else { return $null }
}
$riStatus = 'OK'; $spStatus = 'OK'
try {
    $ris = Get-AllPages "$script:ArmBase/providers/Microsoft.Capacity/reservations?api-version=2022-11-01"
    $ris | ForEach-Object {
        $p = $_.properties
        [pscustomobject][ordered]@{
            ReservationName = $p.displayName; ResourceType = $p.reservedResourceType; Sku = $_.sku.name; Quantity = $p.quantity; Term = $p.term
            AppliedScopeType = $p.appliedScopeType; ProvisioningState = $p.provisioningState; ExpiryDate = ($p.expiryDateTime ?? $p.expiryDate)
            Renew = $p.renew; Utilization1DayPct = Get-UtilAgg $p.utilization 1; Utilization7DayPct = Get-UtilAgg $p.utilization 7
            Utilization30DayPct = Get-UtilAgg $p.utilization 30; UtilizationTrend = $p.utilization.trend; Id = $_.id
        }
    } | Export-Csv -Path (Join-Path $RawPath 'Commitments_Reservations.csv') -NoTypeInformation -Encoding utf8BOM
    Write-Log "Reservations visible: $($ris.Count)" Success
}
catch { $riStatus = "Unavailable: $($_.Exception.Message)"; Add-CollectionError 'Reservations' 'all' $_.Exception.Message; Write-Log "Reservations not accessible: $($_.Exception.Message)" Warning }
try {
    $sps = Get-AllPages "$script:ArmBase/providers/Microsoft.BillingBenefits/savingsPlans?api-version=2022-11-01"
    $sps | ForEach-Object {
        $p = $_.properties
        [pscustomobject][ordered]@{
            SavingsPlanName = $p.displayName; Term = $p.term; CommitmentAmount = $p.commitment.amount; CommitmentCurrency = $p.commitment.currencyCode
            CommitmentGrain = $p.commitment.grain; AppliedScopeType = $p.appliedScopeType; ProvisioningState = $p.provisioningState
            ExpiryDate = $p.expiryDateTime; Renew = $p.renew; Utilization1DayPct = Get-UtilAgg $p.utilization 1
            Utilization7DayPct = Get-UtilAgg $p.utilization 7; Utilization30DayPct = Get-UtilAgg $p.utilization 30; UtilizationTrend = $p.utilization.trend; Id = $_.id
        }
    } | Export-Csv -Path (Join-Path $RawPath 'Commitments_SavingsPlans.csv') -NoTypeInformation -Encoding utf8BOM
    Write-Log "Savings plans visible: $($sps.Count)" Success
}
catch { $spStatus = "Unavailable: $($_.Exception.Message)"; Add-CollectionError 'SavingsPlans' 'all' $_.Exception.Message; Write-Log "Savings plans not accessible: $($_.Exception.Message)" Warning }

#endregion

#region ============================== RETAIL PRICING ==============================

$pricingStatus = 'Skipped'
if (-not $SkipPricing) {
    Write-Log 'Downloading public retail prices for savings estimates (prices.azure.com)...' Section
    function Get-RetailPrices([string]$Filter, [string]$ApiVersion) {
        $items = [System.Collections.Generic.List[object]]::new()
        $u = 'https://prices.azure.com/api/retail/prices?' + $(if ($ApiVersion) { "api-version=$ApiVersion&" }) + '$filter=' + [uri]::EscapeDataString($Filter); $page = 0
        while ($u -and $page -lt 80) {
            $page++; $attempt = 0; $r = $null
            while ($true) {
                $attempt++
                try { $r = Invoke-RestMethod -Uri $u -TimeoutSec 120; break }
                catch { if ($attempt -ge 4) { throw }; Start-Sleep -Seconds (3 * $attempt) }
            }
            foreach ($i in $r.Items) { $items.Add($i) }
            $u = $r.NextPageLink
        }
        return $items.ToArray()
    }
    function Get-RegionsFor([string[]]$Types, [scriptblock]$Filter) {
        @($inventory | Where-Object { $_.type -in $Types -and $_.location -and $_.location -ne 'global' -and (-not $Filter -or (& $Filter $_)) } | ForEach-Object { ([string]$_.location).ToLower() } | Select-Object -Unique)
    }
    $pricingErrors = 0
    try {
        $vmRegions = Get-RegionsFor @('microsoft.compute/virtualmachines', 'microsoft.compute/virtualmachinescalesets', 'microsoft.containerservice/managedclusters')
        $vmPrices = foreach ($rg in $vmRegions) {
            # Pay-as-you-go (+ savings plan rates via preview API) and reservation term prices
            $all = Get-RetailPrices "serviceName eq 'Virtual Machines' and armRegionName eq '$rg' and priceType eq 'Consumption'" '2023-01-01-preview'
            $resv = @{}
            try {
                foreach ($x in (Get-RetailPrices "serviceName eq 'Virtual Machines' and armRegionName eq '$rg' and priceType eq 'Reservation'")) {
                    if (-not $x.armSkuName -or $x.productName -match 'Cloud Services|Dedicated Host') { continue }
                    $hours = if ($x.reservationTerm -eq '3 Years') { 26280 } elseif ($x.reservationTerm -eq '1 Year') { 8760 } else { 0 }
                    if ($hours -gt 0) { $resv["$($x.armSkuName)|$($x.reservationTerm)"] = [double]$x.retailPrice / $hours }
                }
            }
            catch { Add-CollectionError 'Pricing' "VM reservations $rg" $_.Exception.Message }
            $all | Where-Object { $_.unitOfMeasure -eq '1 Hour' -and $_.skuName -notmatch 'Spot|Low Priority' -and $_.productName -notmatch 'Cloud Services|Dedicated Host|Red Hat|RHEL|SUSE|SLES|Ubuntu Pro|SQL|BizTalk' } |
                Group-Object { "$($_.armSkuName)|$(if ($_.productName -match 'Windows') { 'Windows' } else { 'Linux' })" } | ForEach-Object {
                    $min = $_.Group | Sort-Object retailPrice | Select-Object -First 1
                    $sp1 = @($min.savingsPlan | Where-Object { $_.term -eq '1 Year' } | Select-Object -First 1).retailPrice
                    $sp3 = @($min.savingsPlan | Where-Object { $_.term -eq '3 Years' } | Select-Object -First 1).retailPrice
                    [pscustomobject][ordered]@{
                        Region = $rg; VmSize = $min.armSkuName; OS = ($_.Name -split '\|')[1]; HourlyPrice = $min.retailPrice
                        RI1YHourly = if ($resv.ContainsKey("$($min.armSkuName)|1 Year")) { [Math]::Round($resv["$($min.armSkuName)|1 Year"], 5) } else { $null }
                        RI3YHourly = if ($resv.ContainsKey("$($min.armSkuName)|3 Years")) { [Math]::Round($resv["$($min.armSkuName)|3 Years"], 5) } else { $null }
                        SP1YHourly = $sp1; SP3YHourly = $sp3; Currency = $min.currencyCode; ProductName = $min.productName
                    }
                }
            Write-Log "  VM prices loaded (PAYG, reservation, savings plan): $rg" Info
        }
        if ($vmPrices) { $vmPrices | Export-Csv -Path (Join-Path $RawPath 'Pricing_VirtualMachines.csv') -NoTypeInformation -Encoding utf8BOM }
    }
    catch { $pricingErrors++; Add-CollectionError 'Pricing' 'VirtualMachines' $_.Exception.Message }

    $priceSets = @(
        @{ Name = 'ManagedDisks'; Types = @('microsoft.compute/disks'); Filter = "serviceName eq 'Storage' and armRegionName eq '{0}' and priceType eq 'Consumption' and contains(productName,'Managed Disks')"; Keep = { param($i) $i.meterName -match 'Disk$' -or $i.productName -match 'v2|Ultra' } }
        @{ Name = 'AppService'; Types = @('microsoft.web/serverfarms'); Filter = "serviceName eq 'Azure App Service' and armRegionName eq '{0}' and priceType eq 'Consumption'"; Keep = { param($i) $i.unitOfMeasure -eq '1 Hour' } }
        @{ Name = 'LogAnalyticsSentinel'; Types = @('microsoft.operationalinsights/workspaces'); Filter = "armRegionName eq '{0}' and priceType eq 'Consumption' and (serviceName eq 'Log Analytics' or serviceName eq 'Azure Monitor' or serviceName eq 'Sentinel')"; Keep = { param($i) $true } }
        @{ Name = 'Files'; Types = @('microsoft.storage/storageaccounts'); Filter = "serviceName eq 'Storage' and armRegionName eq '{0}' and priceType eq 'Consumption' and contains(productName,'Files')"; Keep = { param($i) $true } }
        @{ Name = 'SqlVCore'; Types = @('microsoft.sql/servers/databases', 'microsoft.sql/servers/elasticpools', 'microsoft.sql/managedinstances'); Filter = "armRegionName eq '{0}' and priceType eq 'Consumption' and (serviceName eq 'SQL Database' or serviceName eq 'SQL Managed Instance') and contains(meterName,'vCore')"; Keep = { param($i) $true } }
    )
    foreach ($ps in $priceSets) {
        try {
            $rows = foreach ($rg in (Get-RegionsFor $ps.Types $null)) {
                Get-RetailPrices ($ps.Filter -f $rg) | Where-Object { & $ps.Keep $_ } | ForEach-Object {
                    [pscustomobject][ordered]@{ Region = $rg; ServiceName = $_.serviceName; ProductName = $_.productName; SkuName = $_.skuName; MeterName = $_.meterName; ArmSkuName = $_.armSkuName; Unit = $_.unitOfMeasure; Price = $_.retailPrice; Currency = $_.currencyCode; TierMinimumUnits = $_.tierMinimumUnits }
                }
            }
            if ($rows) { $rows | Export-Csv -Path (Join-Path $RawPath "Pricing_$($ps.Name).csv") -NoTypeInformation -Encoding utf8BOM }
            Write-Log "  $($ps.Name) prices: $(@($rows).Count) meter(s)" Info
        }
        catch { $pricingErrors++; Add-CollectionError 'Pricing' $ps.Name $_.Exception.Message; Write-Log "  $($ps.Name) prices failed: $($_.Exception.Message)" Warning }
    }
    $pricingStatus = if ($pricingErrors -eq 0) { 'OK' } else { "Partial ($pricingErrors price set(s) failed)" }
}

if ($script:VmSkuRows.Count) { $script:VmSkuRows | Export-Csv -Path (Join-Path $RawPath 'Reference_VmSizes.csv') -NoTypeInformation -Encoding utf8BOM }

#endregion

#region ============================== METADATA & SUMMARY ==============================

$typeSummary | Export-Csv -Path (Join-Path $RawPath 'Collection_Summary.csv') -NoTypeInformation -Encoding utf8BOM
if ($script:MetricCatalogRows.Count) { $script:MetricCatalogRows | Export-Csv -Path (Join-Path $RawPath 'Reference_MetricCatalog.csv') -NoTypeInformation -Encoding utf8BOM }
if ($script:CollectionErrors.Count) { $script:CollectionErrors | Export-Csv -Path (Join-Path $RawPath 'Collection_Errors.csv') -NoTypeInformation -Encoding utf8BOM }
$duration = (Get-Date) - $script:RunStart
$meta = [ordered]@{
    ScriptVersion      = $script:ScriptVersion
    CollectedAtUtc     = [DateTime]::UtcNow.ToString('s') + 'Z'
    CollectedBy        = if ($script:AuthSource -eq 'AzureCLI') { [string]$script:CliAccount.user.name } else { [string]$script:AzContext.Account.Id }
    AuthSource         = $script:AuthSource
    TenantId           = $TenantId
    SubscriptionCount  = $allSubs.Count
    LookbackDays       = $LookbackDays
    WindowStartLocal   = $StartLocal.ToString('yyyy-MM-dd')
    WindowEndLocal     = $EndLocal.AddDays(-1).ToString('yyyy-MM-dd')
    WindowStartUtc     = $StartUtc.ToString('s') + 'Z'
    WindowEndUtc       = $EndUtc.ToString('s') + 'Z'
    TotalHours         = $TotalHours
    TimeZoneId         = $TimeZoneId
    TzOffsetHours      = $TzOffsetHours
    WeekendDays        = $WeekendDays
    BusinessHoursStart = $BusinessHoursStart
    BusinessHoursEnd   = $BusinessHoursEnd
    CostType           = $CostType
    InventoryCount     = $inventory.Count
    MetricCallsTotal   = $script:MetricRequests
    MetricsBatchApi    = [bool]$script:BatchMetrics
    MetricsBatchSize   = $MetricsBatchSize
    GenericTypeCount   = $genericTypeCount
    CostHistoryMonths  = $CostHistoryMonths
    ConfigPath         = if ($script:Config) { $ConfigPath } else { $null }
    DurationMinutes    = [Math]::Round($duration.TotalMinutes, 1)
    DataSources        = [ordered]@{
        Metrics      = if ($SkipMetrics) { 'Skipped' } elseif ($script:BatchMetrics) { 'OK (Metrics Batch API)' } else { 'OK (per-resource ARM metrics)' }
        LogAnalytics = if ($SkipLogAnalytics) { 'Skipped' } else { 'OK' }
        CostData     = $costStatus
        Advisor      = $advisorStatus
        Reservations = $riStatus
        SavingsPlans = $spStatus
        Pricing      = $pricingStatus
        GenericTypes = if ($script:GenericEnabled) { "OK ($genericTypeCount type(s))" } else { 'Skipped' }
    }
    ResourceTypes      = @($typeSummary)
    ErrorCount         = $script:CollectionErrors.Count
}
$meta | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $RawPath 'Collection_Metadata.json') -Encoding utf8

Write-Log ('=' * 78) Section
Write-Log "PHASE 1 COMPLETE in $([Math]::Round($duration.TotalMinutes, 1)) min" Success
Write-Log "  Resources inventoried  : $($inventory.Count)"
Write-Log "  Resource types analysed: $($typeSummary.Count) ($(($typeSummary | Measure-Object -Property Rows -Sum).Sum) rows; $genericTypeCount via generic discovery)"
Write-Log "  Metric API requests    : $($script:MetricRequests) ($(if ($script:BatchMetrics) { 'Metrics Batch API' } else { 'per-resource' }))"
Write-Log "  Cost data              : $costStatus"
Write-Log "  Collection warnings    : $($script:CollectionErrors.Count) (see 01_RawData\Collection_Errors.csv)"
Write-Log ('=' * 78) Section

#endregion
