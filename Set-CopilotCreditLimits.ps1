#requires -Version 5.1
#requires -Modules Az.Accounts

<#
.SYNOPSIS
Applies Copilot Credit (MCSMessages) monthly limits to the agents in one or more Power Platform
environments, or in every environment of an environment group, and reports what was done.

.DESCRIPTION
Signs in interactively as a Power Platform admin through Az.Accounts, resolves the target
environments, inventories the Copilot Studio agents in each one, and writes a per-agent Copilot
Credit threshold through the Power Platform licensing API:

    PUT licensing/environments/{environmentId}/entitlements/MCSMessages/resources/{resourceId}/threshold

IMPORTANT - the api-version matters. The published licensing spec documents api-version=2024-10-01.
The Power Platform admin center calls this route with api-version=1, and only that version takes
effect: a limit written with 2024-10-01 returns 200 and persists in resourceThresholds, but never
appears in Licensing > Copilot Studio > Manage Agents and does not enforce. This script writes with
api-version=1 and falls back to the documented version.

Agents are discovered from three complementary sources:

  * Inventory  - the admin-scoped Power Platform inventory that backs Manage > Inventory. It is the
                 only source that sees inside environments the caller is not a member of, such as
                 personal developer environments, where Dataverse returns HTTP 403.
  * Licensing  - agents that already have Copilot Credit consumption records. Adds month-to-date
                 consumption.
  * Dataverse  - the 'bots' table of the environment. Adds display names.

.PARAMETER EnvironmentGroup
Environment group ID or exact display name. Every environment of the group is targeted.

.PARAMETER Environment
One or more environment IDs or exact environment display names.

.PARAMETER AgentCreditLimit
Monthly Copilot Credit limit to apply to each agent. Required unless -Discover or -ReportOnly is used.

.PARAMETER AgentNotificationThreshold
Percentage of the limit at which administrators are notified. Defaults to 80. The API treats this as
a percentage, and the admin center restricts it to 50-100.

.PARAMETER NotifyIfOverCapacity
Send a notification when an agent exceeds its limit. On by default; use -NotifyIfOverCapacity:$false
to disable.

.PARAMETER StopAgentIfOverLimit
Turn the agent off when it reaches its limit (hard stop).

.PARAMETER ExcludeAgentId
One or more agent/resource IDs to skip.

.PARAMETER ExcludeAgentIdCsv
Path to a CSV of agent/resource IDs to skip. A header named AgentId, ResourceId, BotId or Id is
honoured; a headerless single-column file also works. An exclusion file that yields zero exclusions
stops the run, because it would otherwise silently protect nothing.

.PARAMETER AgentSource
Where to discover agents: Inventory, Licensing, Dataverse, or All (default).

.PARAMETER IncludeFlows
Also write limits to agent flows and cloud flows, not only Copilot Studio agents.

Flows can consume Copilot Credits, so they are always inventoried and reported. They are NOT written
to by default: the threshold API is proven for agents, and enforcement for flows is not yet
confirmed. Use this switch to target them deliberately and verify the result in the admin center.

.PARAMETER LookbackDays
How far back the consumption snapshot is queried. Default 90.

.PARAMETER ReportPath
Folder for the CSV report. Default: .\reports

.PARAMETER Discover
Read-only mode: dump the raw API responses for the first target environment.

.PARAMETER ReportOnly
Inventory and report current limits without writing anything.

.PARAMETER TenantId
Optional tenant ID to pass to Connect-AzAccount.

.PARAMETER Force
Skip confirmation prompts and rewrite limits even when the current value already matches.

.PARAMETER Diagnostics
Show extra detail about discovery and API fallbacks.

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -ReportOnly

Inventories the agents and shows their current limits. Writes nothing.

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -AgentCreditLimit 1000 -WhatIf

Shows exactly which agents would be changed.

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" -AgentCreditLimit 500 -StopAgentIfOverLimit -ExcludeAgentIdCsv .\samples\agent-exclusions.csv -Force

Applies a 500 credit limit with a hard stop to every agent in the group, except those listed.

.LINK
https://github.com/LaureVDH/powerplatform-copilot-credit-limits
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Group')]
param
(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Group')]
    [ValidateNotNullOrEmpty()]
    [string] $EnvironmentGroup,

    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Environment')]
    [ValidateNotNullOrEmpty()]
    [string[]] $Environment,


    [Parameter()]
    [ValidateRange(0, 2147483647)]
    [int] $AgentCreditLimit,

    [Parameter()]
    [ValidateRange(50, 100)]
    [int] $AgentNotificationThreshold = 80,

    [Parameter()]
    [switch] $NotifyIfOverCapacity = $true,

    [Parameter()]
    [switch] $StopAgentIfOverLimit,


    [Parameter()]
    [string[]] $ExcludeAgentId,

    [Parameter()]
    [string] $ExcludeAgentIdCsv,


    [Parameter()]
    [ValidateSet('Licensing', 'Dataverse', 'Inventory', 'All')]
    [string] $AgentSource = 'All',

    [Parameter()]
    [switch] $IncludeFlows,

    [Parameter()]
    [ValidateRange(1, 365)]
    [int] $LookbackDays = 90,

    [Parameter()]
    [string] $ReportPath = (Join-Path -Path $PSScriptRoot -ChildPath 'reports'),

    [Parameter()]
    [switch] $Discover,

    [Parameter()]
    [switch] $ReportOnly,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $TenantId,

    [Parameter()]
    [switch] $Force,

    [Parameter()]
    [switch] $Diagnostics
)

$ErrorActionPreference = 'Stop'
$PowerPlatformApiRoot = 'https://api.powerplatform.com'

# The Power Platform API versions its namespaces independently: licensing is GA at 2024-10-01,
# while environmentmanagement is still on the 2022-03-01-preview contract. Calls fall back through
# the candidate list until one is accepted, so a version bump in either namespace keeps working.
$ApiVersion = '2024-10-01'
$LicensingApiVersions = @('2024-10-01')
$EnvironmentApiVersions = @('2024-10-01', '2022-03-01-preview', '2021-04-01')

# The Power Platform admin center calls the resource threshold route with api-version=1, not the
# 2024-10-01 documented in the licensing spec. A limit written with 2024-10-01 persists in
# resourceThresholds but does NOT appear in Licensing > Copilot Studio > Manage Agents, whereas a
# limit set in the admin center does. Writes therefore use the admin center's version first.
$ThresholdApiVersions = @('1', '2024-10-01')

$EntitlementId = 'MCSMessages'
$GraphApiRoot = 'https://graph.microsoft.com'

#region Infrastructure

function Get-AccessToken
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $ResourceUrl,

        [Parameter()]
        [string] $TenantId
    )

    if (-not (Get-AzContext -ErrorAction SilentlyContinue))
    {
        Write-Host 'No Azure context found. Launching sign-in...' -ForegroundColor Yellow

        if ($TenantId)
        {
            Connect-AzAccount -Tenant $TenantId -WarningAction SilentlyContinue -WhatIf:$false | Out-Null
        }
        else
        {
            Connect-AzAccount -WarningAction SilentlyContinue -WhatIf:$false | Out-Null
        }
    }

    $azAccessToken = Get-AzAccessToken -ResourceUrl $ResourceUrl -WarningAction SilentlyContinue

    if ($azAccessToken.Token -is [System.Security.SecureString])
    {
        return [System.Net.NetworkCredential]::new('', $azAccessToken.Token).Password
    }

    return $azAccessToken.Token
}

function Get-PropertyValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object] $InputObject,

        [Parameter(Mandatory = $true)]
        [string[]] $Paths
    )

    foreach ($path in $Paths)
    {
        $current = $InputObject
        $found = $true

        foreach ($part in $path.Split('.'))
        {
            if ($null -eq $current)
            {
                $found = $false
                break
            }

            $property = $current.PSObject.Properties[$part]
            if ($null -eq $property)
            {
                $found = $false
                break
            }

            $current = $property.Value
        }

        if ($found -and $null -ne $current -and "$current" -ne '')
        {
            return $current
        }
    }

    return $null
}

function Invoke-PowerPlatformApi
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'PUT', 'POST', 'PATCH', 'DELETE')]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [string] $PathOrUri,

        [Parameter()]
        [object] $Body,

        [Parameter()]
        [string] $ApiVersionOverride,

        [Parameter()]
        [switch] $AllPages
    )

    $token = Get-AccessToken -ResourceUrl $PowerPlatformApiRoot -TenantId $TenantId

    $effectiveApiVersion = if ($ApiVersionOverride) { $ApiVersionOverride } else { $ApiVersion }

    if ($PathOrUri -match '^https://')
    {
        $uri = $PathOrUri
    }
    else
    {
        $separator = if ($PathOrUri.Contains('?')) { '&' } else { '?' }
        $uri = '{0}/{1}{2}api-version={3}' -f $PowerPlatformApiRoot.TrimEnd('/'), $PathOrUri.TrimStart('/'), $separator, $effectiveApiVersion
    }

    $headers = @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
    }

    if ($null -ne $Body)
    {
        $headers['Content-Type'] = 'application/json'
        $jsonBody = $Body | ConvertTo-Json -Depth 20
    }

    $results = New-Object System.Collections.Generic.List[object]

    do
    {
        $parameters = @{
            Method      = $Method
            Uri         = $uri
            Headers     = $headers
            ErrorAction = 'Stop'
        }

        if ($null -ne $Body)
        {
            $parameters['Body'] = $jsonBody
        }

        $response = Invoke-RestMethod @parameters

        if (-not $AllPages)
        {
            return $response
        }

        $page = $null

        foreach ($collectionName in @('value', 'items', 'resources', 'users', 'environments'))
        {
            $property = $response.PSObject.Properties[$collectionName]
            if ($property -and $null -ne $property.Value)
            {
                $page = $property.Value
                break
            }
        }

        if ($null -eq $page)
        {
            $page = $response
        }

        foreach ($item in @($page))
        {
            if ($null -ne $item)
            {
                $results.Add($item) | Out-Null
            }
        }

        $nextUri = $response.'@odata.nextLink'

        if (-not $nextUri)
        {
            $nextUri = $response.nextLink
        }

        $continuationToken = Get-PropertyValue -InputObject $response -Paths @('continuationToken', 'ContinuationToken')

        if ($nextUri)
        {
            $uri = "$nextUri"
        }
        elseif ($continuationToken)
        {
            $encoded = [uri]::EscapeDataString("$continuationToken")

            if ($uri -match 'continuation-token=')
            {
                $uri = $uri -replace 'continuation-token=[^&]*', "continuation-token=$encoded"
            }
            else
            {
                $uri = "$uri&continuation-token=$encoded"
            }
        }
        else
        {
            $uri = $null
        }
    }
    while ($uri)

    return $results.ToArray()
}

function Invoke-PowerPlatformApiWithVersions
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $PathOrUri,

        [Parameter(Mandatory = $true)]
        [string[]] $ApiVersions,

        [Parameter()]
        [switch] $AllPages
    )

    $lastError = $null

    foreach ($version in $ApiVersions)
    {
        try
        {
            $result = if ($AllPages)
            {
                @(Invoke-PowerPlatformApi -Method GET -PathOrUri $PathOrUri -ApiVersionOverride $version -AllPages)
            }
            else
            {
                Invoke-PowerPlatformApi -Method GET -PathOrUri $PathOrUri -ApiVersionOverride $version
            }

            Write-Verbose "'$PathOrUri' succeeded with api-version=$version."
            return $result
        }
        catch
        {
            $lastError = $_
            Write-Verbose "'$PathOrUri' failed with api-version=$version. $($_.Exception.Message)"
        }
    }

    if ($lastError)
    {
        throw $lastError
    }
}

function Invoke-BapApi
{
    <#
        The Business Application Platform admin API is the documented, stable way to list every
        environment in a tenant with its Dataverse instance URL and environment group. See
        https://learn.microsoft.com/power-platform/admin/list-environments
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $root = 'https://api.bap.microsoft.com'
    $token = Get-AccessToken -ResourceUrl $root -TenantId $TenantId

    $headers = @{
        Authorization = "Bearer $token"
        Accept        = 'application/json'
    }

    $uri = '{0}/{1}' -f $root, $Path.TrimStart('/')
    $results = New-Object System.Collections.Generic.List[object]

    do
    {
        $response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop

        foreach ($item in @($response.value))
        {
            if ($null -ne $item)
            {
                $results.Add($item) | Out-Null
            }
        }

        $uri = $response.nextLink
    }
    while ($uri)

    return $results.ToArray()
}

function Invoke-DataverseApi
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $InstanceApiUrl,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $root = $InstanceApiUrl.TrimEnd('/')
    $token = Get-AccessToken -ResourceUrl $root -TenantId $TenantId

    $headers = @{
        Authorization      = "Bearer $token"
        Accept             = 'application/json'
        'OData-Version'    = '4.0'
        'OData-MaxVersion' = '4.0'
        Prefer             = 'odata.include-annotations="*"'
    }

    $uri = '{0}/api/data/v9.2/{1}' -f $root, $Path.TrimStart('/')
    $results = New-Object System.Collections.Generic.List[object]

    do
    {
        $response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop

        foreach ($item in @($response.value))
        {
            if ($null -ne $item)
            {
                $results.Add($item) | Out-Null
            }
        }

        $uri = $response.'@odata.nextLink'
    }
    while ($uri)

    return $results.ToArray()
}

function Write-Section
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $Text
    )

    if ($script:SuppressSectionOutput)
    {
        return
    }

    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * $Text.Length) -ForegroundColor DarkCyan
}

function Write-Detail
{
    <#
        Console output that is suppressed during the exclusion preflight pass, which re-runs discovery
        purely to validate exclusions and should not duplicate the real run's output.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $Text,

        [Parameter()]
        [string] $ForegroundColor = 'White'
    )

    if ($script:SuppressSectionOutput)
    {
        return
    }

    Write-Host $Text -ForegroundColor $ForegroundColor
}

#endregion

#region Exclusions

function Get-ExclusionSet
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [AllowNull()]
        [string[]] $InlineValues,

        [Parameter()]
        [AllowNull()]
        [string] $CsvPath,

        [Parameter(Mandatory = $true)]
        [string[]] $ColumnCandidates,

        [Parameter(Mandatory = $true)]
        [string] $Label
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($value in @($InlineValues))
    {
        if ($value -and "$value".Trim())
        {
            $set.Add("$value".Trim()) | Out-Null
        }
    }

    if ($CsvPath)
    {
        if (-not (Test-Path -LiteralPath $CsvPath))
        {
            throw "$Label exclusion file '$CsvPath' was not found."
        }

        $rows = @(Import-Csv -LiteralPath $CsvPath -ErrorAction Stop)
        $matchedColumn = $null

        if ($rows.Count -gt 0)
        {
            foreach ($candidate in $ColumnCandidates)
            {
                $property = $rows[0].PSObject.Properties | Where-Object { ($_.Name -replace '\s', '') -eq $candidate }

                if ($property)
                {
                    $matchedColumn = @($property)[0].Name
                    break
                }
            }
        }

        if ($matchedColumn)
        {
            foreach ($row in $rows)
            {
                $value = "$($row.$matchedColumn)".Trim()

                if ($value)
                {
                    $set.Add($value) | Out-Null
                }
            }
        }
        else
        {
            # Headerless or unrecognised header: treat every non-empty line as an ID.
            foreach ($line in @(Get-Content -LiteralPath $CsvPath))
            {
                $value = "$line".Split(',')[0].Trim().Trim('"')

                if ($value -and $value -notmatch '^\s*#')
                {
                    $set.Add($value) | Out-Null
                }
            }
        }

        Write-Verbose "Loaded $($set.Count) $Label exclusion(s) from '$CsvPath'."

        if ($set.Count -eq 0)
        {
            # Passing an exclusion file means the caller intends to protect something. Silently
            # treating an empty or unparsable file as "exclude nothing" would apply the limit to
            # EVERY target, which is the opposite of the caller's intent.
            throw ("The $Label exclusion file '$CsvPath' produced 0 exclusions. " +
                   "The file is empty, or its ID column was not recognised (expected one of: " +
                   ($ColumnCandidates -join ', ') + "). " +
                   "Refusing to continue, because running without the intended exclusions would apply " +
                   "the limit to every $Label. Fix the file, or drop the parameter to target everything.")
        }
    }

    return ,$set
}

#endregion

#region Environment resolution

function Get-EnvironmentId
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object] $EnvironmentObject
    )

    $candidate = Get-PropertyValue -InputObject $EnvironmentObject -Paths @(
        'EnvironmentName',
        'environmentName',
        'EnvironmentId',
        'environmentId',
        'name',
        'Internal.name',
        'properties.environmentId',
        'Internal.properties.environmentId',
        'id',
        'Internal.id',
        'properties.id'
    )

    if (-not $candidate)
    {
        return $null
    }

    $candidateText = "$candidate"

    if ($candidateText -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
    {
        return $candidateText
    }

    if ($candidateText -match '/environments/([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})')
    {
        return $Matches[1]
    }

    return $candidateText
}

function Get-EnvironmentDisplayName
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object] $EnvironmentObject
    )

    return Get-PropertyValue -InputObject $EnvironmentObject -Paths @(
        'DisplayName',
        'displayName',
        'EnvironmentDisplayName',
        'Internal.properties.displayName',
        'properties.displayName',
        'name',
        'Internal.name',
        'EnvironmentName'
    )
}

function Get-EnvironmentInstanceApiUrl
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object] $EnvironmentObject
    )

    $url = Get-PropertyValue -InputObject $EnvironmentObject -Paths @(
        'Internal.properties.linkedEnvironmentMetadata.instanceApiUrl',
        'properties.linkedEnvironmentMetadata.instanceApiUrl',
        'Internal.properties.linkedEnvironmentMetadata.instanceUrl',
        'properties.linkedEnvironmentMetadata.instanceUrl',
        'linkedEnvironmentMetadata.instanceApiUrl',
        'linkedEnvironmentMetadata.instanceUrl',
        'instanceApiUrl',
        'instanceUrl'
    )

    if ($url)
    {
        return "$url".TrimEnd('/')
    }

    return $null
}

function Get-AllEnvironments
{
    [CmdletBinding()]
    param()

    if ($script:AllEnvironmentsCache)
    {
        return $script:AllEnvironmentsCache
    }

    $environments = $null
    $failures = New-Object System.Collections.Generic.List[string]

    # 1. Documented BAP admin API. Returns instanceApiUrl and environment group membership.
    try
    {
        $candidate = @(Invoke-BapApi -Path 'providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2020-10-01&$expand=properties.capacity,properties.addons')

        if ($candidate.Count -gt 0)
        {
            $environments = $candidate
            Write-Verbose "Listed $($candidate.Count) environment(s) from the BAP admin API."
        }
    }
    catch
    {
        $failures.Add("BAP admin API: $($_.Exception.Message)") | Out-Null
        Write-Verbose "Could not list environments from the BAP admin API. $($_.Exception.Message)"
    }

    # 2. Power Platform API environmentmanagement namespace.
    if ($null -eq $environments -or $environments.Count -eq 0)
    {
        foreach ($path in @('environmentmanagement/environments', 'appmanagement/environments'))
        {
            try
            {
                $candidate = @(Invoke-PowerPlatformApiWithVersions -PathOrUri $path -ApiVersions $EnvironmentApiVersions -AllPages)

                if ($candidate.Count -gt 0)
                {
                    $environments = $candidate
                    Write-Verbose "Listed $($candidate.Count) environment(s) from '$path'."
                    break
                }
            }
            catch
            {
                $failures.Add("${path}: $($_.Exception.Message)") | Out-Null
                Write-Verbose "Could not list environments from '$path'. $($_.Exception.Message)"
            }
        }
    }

    # 3. Optional Power Apps administration module, only if it is actually installed.
    if ($null -eq $environments -or $environments.Count -eq 0)
    {
        if (Get-Module -ListAvailable -Name Microsoft.PowerApps.Administration.PowerShell)
        {
            Write-Verbose 'Falling back to the Power Apps administration module for the environment list.'

            try
            {
                Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop

                if (-not $global:currentSession)
                {
                    if ($TenantId)
                    {
                        Add-PowerAppsAccount -Endpoint prod -TenantID $TenantId | Out-Null
                    }
                    else
                    {
                        Add-PowerAppsAccount -Endpoint prod | Out-Null
                    }
                }

                $environments = @(Get-AdminPowerAppEnvironment)
            }
            catch
            {
                $failures.Add("Microsoft.PowerApps.Administration.PowerShell: $($_.Exception.Message)") | Out-Null
            }
        }
        else
        {
            Write-Verbose 'Microsoft.PowerApps.Administration.PowerShell is not installed; skipping that fallback.'
        }
    }

    if ($null -eq $environments -or $environments.Count -eq 0)
    {
        $detail = if ($failures.Count -gt 0) { [Environment]::NewLine + '  - ' + ($failures -join ([Environment]::NewLine + '  - ')) } else { '' }

        throw ("Could not list Power Platform environments. Confirm you signed in as a Power Platform or Global Administrator " +
               "in tenant '$TenantId', then rerun with -Verbose for the full API trace.$detail")
    }

    $script:AllEnvironmentsCache = $environments
    return $environments
}

function Get-EnvironmentGroups
{
    [CmdletBinding()]
    param()

    foreach ($path in @('environmentmanagement/environmentGroups', 'environmentmanagement/groups'))
    {
        try
        {
            $groups = @(Invoke-PowerPlatformApiWithVersions -PathOrUri $path -ApiVersions $EnvironmentApiVersions -AllPages)

            if ($groups.Count -gt 0)
            {
                return $groups
            }
        }
        catch
        {
            Write-Verbose "Could not list environment groups from '$path'. $($_.Exception.Message)"
        }
    }

    # Fallback: derive the distinct groups referenced by the environments themselves.
    Write-Verbose 'Deriving environment groups from environment membership.'

    $derived = @{}

    foreach ($environmentObject in Get-AllEnvironments)
    {
        $groupId = Get-PropertyValue -InputObject $environmentObject -Paths @(
            'properties.parentEnvironmentGroup.id',
            'Internal.properties.parentEnvironmentGroup.id',
            'properties.environmentGroupId',
            'Internal.properties.environmentGroupId',
            'environmentGroupId'
        )

        if (-not $groupId)
        {
            continue
        }

        $groupName = Get-PropertyValue -InputObject $environmentObject -Paths @(
            'properties.parentEnvironmentGroup.displayName',
            'Internal.properties.parentEnvironmentGroup.displayName',
            'properties.parentEnvironmentGroup.name',
            'Internal.properties.parentEnvironmentGroup.name',
            'properties.environmentGroupName',
            'Internal.properties.environmentGroupName'
        )

        if (-not $derived.ContainsKey("$groupId"))
        {
            $derived["$groupId"] = [PSCustomObject]@{
                id          = "$groupId"
                displayName = if ($groupName) { "$groupName" } else { "$groupId" }
            }
        }
    }

    if ($derived.Count -gt 0)
    {
        return @($derived.Values)
    }

    throw 'Could not list environment groups from the Power Platform API, and no environment reported group membership.'
}

function Resolve-EnvironmentGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $GroupIdOrName
    )

    $groups = Get-EnvironmentGroups
    $isGuid = $GroupIdOrName -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

    if ($isGuid)
    {
        $matched = @($groups | Where-Object {
            (Get-PropertyValue -InputObject $_ -Paths @('id', 'name', 'groupId')) -eq $GroupIdOrName
        })
    }
    else
    {
        $matched = @($groups | Where-Object {
            (Get-PropertyValue -InputObject $_ -Paths @('displayName', 'properties.displayName', 'name')) -eq $GroupIdOrName
        })
    }

    if ($matched.Count -eq 0)
    {
        throw "Environment group '$GroupIdOrName' was not found."
    }

    if ($matched.Count -gt 1)
    {
        throw "Environment group name '$GroupIdOrName' matched more than one group. Use the group ID."
    }

    return $matched[0]
}

function Get-EnvironmentGroupValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object] $EnvironmentObject
    )

    return Get-PropertyValue -InputObject $EnvironmentObject -Paths @(
        'environmentGroupId',
        'environmentGroupName',
        'groupId',
        'groupName',
        'EnvironmentGroupId',
        'EnvironmentGroupName',
        'Internal.properties.environmentGroupId',
        'Internal.properties.environmentGroupName',
        'Internal.properties.parentEnvironmentGroup.id',
        'Internal.properties.parentEnvironmentGroup.name',
        'Internal.properties.linkedEnvironmentGroup.id',
        'Internal.properties.linkedEnvironmentGroup.name',
        'Internal.properties.governanceConfiguration.environmentGroupId',
        'properties.environmentGroupId',
        'properties.environmentGroupName',
        'properties.parentEnvironmentGroup.id',
        'properties.parentEnvironmentGroup.name',
        'properties.linkedEnvironmentGroup.id',
        'properties.linkedEnvironmentGroup.name',
        'properties.governanceConfiguration.environmentGroupId'
    )
}

function Resolve-TargetEnvironments
{
    [CmdletBinding()]
    param()

    # NOTE: $PSCmdlet inside a function refers to THAT function, not to the script, so
    # $PSCmdlet.ParameterSetName here would return '__AllParameterSets' and never 'Group'.
    # The bound script parameter is the reliable signal.
    if ($EnvironmentGroup)
    {
        $group = Resolve-EnvironmentGroup -GroupIdOrName $EnvironmentGroup
        $groupId = "$(Get-PropertyValue -InputObject $group -Paths @('id', 'name', 'groupId'))"
        $groupName = "$(Get-PropertyValue -InputObject $group -Paths @('displayName', 'properties.displayName', 'name'))"

        Write-Host "Environment group: $groupName ($groupId)" -ForegroundColor Green

        $environments = $null

        foreach ($pathTemplate in @(
            'environmentmanagement/environmentGroups/{0}/environments',
            'environmentmanagement/groups/{0}/environments'
        ))
        {
            $path = $pathTemplate -f $groupId

            try
            {
                $candidate = @(Invoke-PowerPlatformApiWithVersions -PathOrUri $path -ApiVersions $EnvironmentApiVersions -AllPages)

                if ($candidate.Count -gt 0)
                {
                    $environments = $candidate
                    break
                }
            }
            catch
            {
                Write-Verbose "Could not list environments from '$path'. $($_.Exception.Message)"
            }
        }

        if ($null -eq $environments -or $environments.Count -eq 0)
        {
            Write-Verbose 'Falling back to filtering the full environment list by group ID or name.'

            $all = Get-AllEnvironments
            $environments = @($all | Where-Object {
                $value = Get-EnvironmentGroupValue -EnvironmentObject $_
                $value -eq $groupId -or $value -eq $groupName
            })
        }

        if ($environments.Count -eq 0)
        {
            # Group membership is exposed inconsistently across API surfaces, so show what the
            # environments actually reported rather than just failing.
            Write-Warning "No environments matched group '$groupName' ($groupId)."

            $observed = Get-AllEnvironments |
                Select-Object @{ Name = 'Environment'; Expression = { Get-EnvironmentDisplayName -EnvironmentObject $_ } },
                              @{ Name = 'GroupValueSeen'; Expression = { Get-EnvironmentGroupValue -EnvironmentObject $_ } } |
                Where-Object { $_.GroupValueSeen }

            if ($observed)
            {
                Write-Host 'Group values reported by environments in this tenant:' -ForegroundColor Yellow
                $observed | Format-Table -AutoSize | Out-Host
            }
            else
            {
                Write-Host 'No environment reported any group membership field.' -ForegroundColor Yellow
                Write-Host 'The group may be empty, or membership is not exposed on this API surface.' -ForegroundColor Yellow
            }

            throw ("No environments were found in environment group '$groupName'. " +
                   'Confirm the group contains environments, or target them with -Environment instead.')
        }

        # Group endpoints can return reduced objects; re-hydrate from the full environment list so the
        # Dataverse instance URL is available for agent-name enrichment and the user inventory.
        $all = Get-AllEnvironments

        $hydrated = foreach ($environmentObject in $environments)
        {
            $environmentId = Get-EnvironmentId -EnvironmentObject $environmentObject
            $full = $all | Where-Object { (Get-EnvironmentId -EnvironmentObject $_) -eq $environmentId } | Select-Object -First 1

            if ($full) { $full } else { $environmentObject }
        }

        return @($hydrated)
    }

    $all = Get-AllEnvironments
    $resolved = New-Object System.Collections.Generic.List[object]

    foreach ($item in $Environment)
    {
        $match = @($all | Where-Object {
            (Get-EnvironmentId -EnvironmentObject $_) -eq $item -or
            (Get-EnvironmentDisplayName -EnvironmentObject $_) -eq $item
        })

        if ($match.Count -eq 0)
        {
            throw "Environment '$item' was not found."
        }

        if ($match.Count -gt 1)
        {
            throw "Environment name '$item' matched more than one environment. Use the environment ID."
        }

        $resolved.Add($match[0]) | Out-Null
    }

    return $resolved.ToArray()
}

#endregion

#region Agents

function Get-LicensingAgents
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $EnvironmentId
    )

    $fromDate = (Get-Date).AddDays(-$LookbackDays).ToString('yyyy-MM-dd')
    $toDate = (Get-Date).ToString('yyyy-MM-dd')

    # The REST API expects camelCase fromDate/toDate. The kebab-case spelling (from-date/to-date) is
    # the pac CLI FLAG name and is rejected with HTTP 400, so camelCase is attempted first.
    $pathTemplates = @(
        'licensing/entitlements/{0}/environments/{1}/resources?fromDate={2}&toDate={3}',
        'licensing/environments/{1}/entitlements/{0}/resources?fromDate={2}&toDate={3}',
        'licensing/entitlements/{0}/environments/{1}/resources?from-date={2}&to-date={3}&page-size=200',
        'licensing/environments/{1}/entitlements/{0}/resources?from-date={2}&to-date={3}&page-size=200'
    )

    $sawForbidden = $false

    foreach ($template in $pathTemplates)
    {
        $path = $template -f $EntitlementId, $EnvironmentId, $fromDate, $toDate

        try
        {
            $response = @(Invoke-PowerPlatformApi -Method GET -PathOrUri $path -AllPages)
            Write-Verbose "Licensing resources for $EnvironmentId came from '$path' ($($response.Count) row(s))."
            return $response
        }
        catch
        {
            $statusCode = $null

            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response)
            {
                $statusCode = [int] $_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 403)
            {
                # This endpoint is consistently forbidden for delegated admin tokens in some tenants.
                # Retrying the remaining spellings only adds noise.
                Write-Verbose "Licensing resources are forbidden for $EnvironmentId (HTTP 403); skipping remaining path variants."
                $sawForbidden = $true
                break
            }

            Write-Verbose "Licensing resource path failed: '$path'. $($_.Exception.Message)"
        }
    }

    if ($sawForbidden)
    {
        Write-Verbose ("Per-agent consumption is not readable for environment $EnvironmentId (HTTP 403). " +
                       "Agents are still discovered and limits can still be written; only the 'Consumed' column is blank.")
        return @()
    }

    Write-Warning "Could not read licensing resources for environment $EnvironmentId. Agent discovery falls back to Dataverse."
    return @()
}

function Get-InventoryResources
{
    <#
        Queries the Power Platform inventory that backs Manage > Inventory in the admin center:

            POST resourcequery/resources/query

        This is an ADMIN-SCOPED control-plane API (KQL over Azure Resource Graph). Unlike reading the
        Dataverse bots table, it does not require the caller to be a member of each environment, so it
        returns agents in personal developer environments that would otherwise fail with HTTP 403.

        Resource types that can consume Copilot Credits:
          microsoft.copilotstudio/agents      - Copilot Studio agents
          microsoft.powerautomate/agentflows  - agent flows
          microsoft.powerautomate/cloudflows  - cloud flows (AI Builder actions consume credits)
    #>
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [string[]] $ResourceTypes = @(
            'microsoft.copilotstudio/agents',
            'microsoft.powerautomate/agentflows',
            'microsoft.powerautomate/cloudflows'
        ),

        [Parameter()]
        [string] $EnvironmentIdFilter
    )

    $quotedTypes = @($ResourceTypes | ForEach-Object { "'$_'" })

    # The inventory is tenant-wide and unchanged within a run, so cache it. Without this, any second
    # discovery pass (such as the exclusion preflight) would repeat the whole query.
    $cacheKey = "{0}|{1}" -f ($quotedTypes -join ','), $EnvironmentIdFilter

    if ($script:InventoryCache -and $script:InventoryCache.ContainsKey($cacheKey))
    {
        return $script:InventoryCache[$cacheKey]
    }

    if (-not $script:InventoryCache)
    {
        $script:InventoryCache = @{}
    }

    $clauses = New-Object System.Collections.Generic.List[object]

    $clauses.Add([ordered]@{
        '$type'   = 'where'
        FieldName = 'type'
        Operator  = 'in~'
        Values    = $quotedTypes
    }) | Out-Null

    if ($EnvironmentIdFilter)
    {
        $clauses.Add([ordered]@{
            '$type'   = 'where'
            FieldName = 'properties.environmentId'
            Operator  = '=~'
            Values    = @("'$EnvironmentIdFilter'")
        }) | Out-Null
    }

    $results = New-Object System.Collections.Generic.List[object]
    $skipToken = ''
    $page = 0

    do
    {
        $page++

        $options = [ordered]@{ Top = 1000 }

        if ($skipToken)
        {
            $options['SkipToken'] = $skipToken
        }
        else
        {
            $options['Skip'] = 0
        }

        $body = [ordered]@{
            Options   = $options
            TableName = 'PowerPlatformResources'
            Clauses   = $clauses.ToArray()
        }

        try
        {
            $response = Invoke-PowerPlatformApi -Method POST -PathOrUri 'resourcequery/resources/query' -Body $body
        }
        catch
        {
            $statusCode = $null

            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response)
            {
                $statusCode = [int] $_.Exception.Response.StatusCode
            }

            Write-Verbose "Inventory query failed (HTTP $statusCode). $($_.Exception.Message)"

            return [PSCustomObject]@{
                Succeeded  = $false
                Resources  = @()
                Error      = "inventory query failed (HTTP $statusCode): $($_.Exception.Message)"
                StatusCode = $statusCode
            }
        }

        foreach ($item in @($response.data))
        {
            if ($null -ne $item)
            {
                $results.Add($item) | Out-Null
            }
        }

        $skipToken = "$($response.skipToken)"
        Write-Verbose "Inventory page $page returned $(@($response.data).Count) row(s); total reported $($response.totalRecords)."
    }
    while ($skipToken -and $page -lt 50)

    $inventoryResult = [PSCustomObject]@{
        Succeeded  = $true
        Resources  = $results.ToArray()
        Error      = $null
        StatusCode = 200
    }

    $script:InventoryCache[$cacheKey] = $inventoryResult
    return $inventoryResult
}

function Get-DataverseAgents
{
    <#
        Returns a result object rather than a bare list, because "this environment has no agents" and
        "this environment could not be inspected" must not look the same to the caller. Reporting an
        inaccessible environment as empty would understate the estate in a governance report.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $InstanceApiUrl
    )

    $query = 'bots?$select=botid,name,schemaname,statecode,createdon,_ownerid_value&$orderby=name asc'

    if ($script:DataverseCache -and $script:DataverseCache.ContainsKey($InstanceApiUrl))
    {
        return $script:DataverseCache[$InstanceApiUrl]
    }

    if (-not $script:DataverseCache)
    {
        $script:DataverseCache = @{}
    }

    try
    {
        $bots = @(Invoke-DataverseApi -InstanceApiUrl $InstanceApiUrl -Path $query)

        $dataverseResult = [PSCustomObject]@{
            Succeeded  = $true
            Agents     = $bots
            Error      = $null
            StatusCode = 200
        }

        $script:DataverseCache[$InstanceApiUrl] = $dataverseResult
        return $dataverseResult
    }
    catch
    {
        $statusCode = $null

        if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response)
        {
            $statusCode = [int] $_.Exception.Response.StatusCode
        }

        $reason = if ($statusCode -eq 403)
        {
            "access denied (HTTP 403) - the calling account is not a member of this environment"
        }
        elseif ($statusCode -eq 404)
        {
            'no Dataverse instance found (HTTP 404)'
        }
        else
        {
            $_.Exception.Message
        }

        $failureResult = [PSCustomObject]@{
            Succeeded  = $false
            Agents     = @()
            Error      = $reason
            StatusCode = $statusCode
        }

        $script:DataverseCache[$InstanceApiUrl] = $failureResult
        return $failureResult
    }
}

function Get-ResourceThresholdMap
{
    <#
        The resourceThresholds endpoint is TENANT-WIDE: it returns thresholds for every environment.
        Each record carries its own environmentId, so the map must be keyed on environmentId + resourceId.
        Keying on resourceId alone lets a threshold from environment A masquerade as the current limit
        for a same-named resource in environment B, which would silently skip or misreport agents.
    #>
    [CmdletBinding()]
    param()

    $map = @{}

    try
    {
        $thresholds = @(Invoke-PowerPlatformApiWithVersions -PathOrUri ('licensing/entitlements/{0}/resourceThresholds' -f $EntitlementId) -ApiVersions $ThresholdApiVersions)
    }
    catch
    {
        Write-Warning "Could not read existing resource thresholds. $($_.Exception.Message)"
        return $map
    }

    foreach ($threshold in $thresholds)
    {
        $resourceId = "$(Get-PropertyValue -InputObject $threshold -Paths @('resourceId', 'ResourceId'))"
        $thresholdEnvironmentId = "$(Get-PropertyValue -InputObject $threshold -Paths @('environmentId', 'EnvironmentId'))"

        if ($resourceId)
        {
            $map[(Get-ThresholdKey -EnvironmentId $thresholdEnvironmentId -ResourceId $resourceId)] = $threshold
        }
    }

    Write-Verbose "Loaded $($map.Count) existing resource threshold(s) across all environments."
    return $map
}

function Get-ThresholdKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $EnvironmentId,

        [Parameter(Mandatory = $true)]
        [string] $ResourceId
    )

    return ('{0}|{1}' -f $EnvironmentId, $ResourceId).ToLowerInvariant()
}

function Set-AgentThreshold
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $TargetEnvironmentId,

        [Parameter(Mandatory = $true)]
        [string] $ResourceId,

        [Parameter(Mandatory = $true)]
        [int] $Limit,

        [Parameter(Mandatory = $true)]
        [int] $NotificationThreshold,

        [Parameter(Mandatory = $true)]
        [bool] $NotifyOverCapacity,

        [Parameter(Mandatory = $true)]
        [bool] $StopOverCapacity
    )

    $path = 'licensing/environments/{0}/entitlements/{1}/resources/{2}/threshold' -f $TargetEnvironmentId, $EntitlementId, $ResourceId

    $body = @{
        entitlementId         = $EntitlementId
        environmentId         = $TargetEnvironmentId
        resourceId            = $ResourceId
        limit                 = $Limit
        notificationThreshold = $NotificationThreshold
        notifyIfOverCapacity  = $NotifyOverCapacity
        stopIfOverCapacity    = $StopOverCapacity
        stopResource          = $false
    }

    # Use the api-version the admin center itself uses, so the limit lands where Manage Agents
    # reads it. Falls back to the documented version if that is rejected.
    $lastError = $null

    foreach ($version in $ThresholdApiVersions)
    {
        try
        {
            $result = Invoke-PowerPlatformApi -Method PUT -PathOrUri $path -Body $body -ApiVersionOverride $version
            Write-Verbose "Threshold write for $ResourceId succeeded with api-version=$version."
            return $result
        }
        catch
        {
            $lastError = $_
            Write-Verbose "Threshold write failed with api-version=$version. $($_.Exception.Message)"
        }
    }

    throw $lastError
}

function Invoke-AgentScope
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object[]] $Environments,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]] $Exclusions,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]] $Report
    )

    $thresholdMap = Get-ResourceThresholdMap

    # Track which exclusions actually matched something. An exclusion that never matches is usually a
    # typo or a stale ID, and it silently leaves the resource it was meant to protect unprotected.
    $script:MatchedExclusions = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    # The inventory API is tenant-wide and admin-scoped, so it is fetched once and indexed by
    # environment. It is the only source that reliably sees inside environments the caller is not a
    # member of, such as personal developer environments.
    $inventoryByEnvironment = @{}

    if ($AgentSource -eq 'Inventory' -or $AgentSource -eq 'All')
    {
        Write-Detail 'Querying the Power Platform inventory (admin-scoped)...' -ForegroundColor Cyan
        $inventoryResult = Get-InventoryResources

        if ($inventoryResult.Succeeded)
        {
            foreach ($resource in $inventoryResult.Resources)
            {
                $resourceEnvironmentId = "$(Get-PropertyValue -InputObject $resource -Paths @('properties.environmentId', 'environmentId'))"

                if (-not $resourceEnvironmentId)
                {
                    continue
                }

                $key = $resourceEnvironmentId.ToLowerInvariant()

                if (-not $inventoryByEnvironment.ContainsKey($key))
                {
                    $inventoryByEnvironment[$key] = New-Object System.Collections.Generic.List[object]
                }

                $inventoryByEnvironment[$key].Add($resource) | Out-Null
            }

            Write-Detail ("Inventory returned {0} credit-consuming resource(s) across {1} environment(s)." -f
                        $inventoryResult.Resources.Count, $inventoryByEnvironment.Keys.Count) -ForegroundColor Green
        }
        else
        {
            Write-Warning "Inventory query unavailable: $($inventoryResult.Error)"
            Write-Warning 'Falling back to licensing and Dataverse discovery only.'
        }
    }

    foreach ($environmentObject in $Environments)
    {
        $environmentId = Get-EnvironmentId -EnvironmentObject $environmentObject
        $environmentName = Get-EnvironmentDisplayName -EnvironmentObject $environmentObject
        $instanceUrl = Get-EnvironmentInstanceApiUrl -EnvironmentObject $environmentObject

        Write-Section "Environment: $environmentName ($environmentId)"

        $agents = @{}

        if ($AgentSource -eq 'Inventory' -or $AgentSource -eq 'All')
        {
            $key = "$environmentId".ToLowerInvariant()

            if ($inventoryByEnvironment.ContainsKey($key))
            {
                foreach ($resource in $inventoryByEnvironment[$key])
                {
                    $resourceId = "$(Get-PropertyValue -InputObject $resource -Paths @('name', 'id'))"

                    if (-not $resourceId)
                    {
                        continue
                    }

                    # The inventory id is a full ARM-style path; the last segment is the resource GUID.
                    if ($resourceId -match '/([^/]+)$')
                    {
                        $resourceId = $Matches[1]
                    }

                    $resourceType = "$(Get-PropertyValue -InputObject $resource -Paths @('type'))"

                    $agents[$resourceId] = [PSCustomObject]@{
                        ResourceId   = $resourceId
                        DisplayName  = "$(Get-PropertyValue -InputObject $resource -Paths @('properties.displayName', 'properties.name', 'name'))"
                        Consumed     = $null
                        Source       = 'Inventory'
                        ResourceType = $resourceType
                    }
                }
            }
        }

        if ($AgentSource -eq 'Licensing' -or $AgentSource -eq 'All')
        {
            foreach ($resource in Get-LicensingAgents -EnvironmentId $environmentId)
            {
                $resourceId = "$(Get-PropertyValue -InputObject $resource -Paths @('resourceId', 'ResourceId', 'id'))"

                if (-not $resourceId)
                {
                    continue
                }

                if ($agents.ContainsKey($resourceId))
                {
                    # Keep the inventory display name, add the consumption figure.
                    $agents[$resourceId].Consumed = Get-PropertyValue -InputObject $resource -Paths @('consumed', 'Consumed')
                    $agents[$resourceId].Source = "$($agents[$resourceId].Source)+Licensing"
                }
                else
                {
                    $agents[$resourceId] = [PSCustomObject]@{
                        ResourceId   = $resourceId
                        DisplayName  = "$(Get-PropertyValue -InputObject $resource -Paths @('metadata.ProductName', 'metadata.productName', 'metadata.Feature', 'name'))"
                        Consumed     = Get-PropertyValue -InputObject $resource -Paths @('consumed', 'Consumed')
                        Source       = 'Licensing'
                        ResourceType = ''
                    }
                }
            }
        }

        $discoveryFailed = $false
        $discoveryError = $null

        if ($AgentSource -eq 'Dataverse' -or $AgentSource -eq 'All')
        {
            if (-not $instanceUrl)
            {
                $discoveryFailed = $true
                $discoveryError = 'no Dataverse instance URL was found for this environment'
                Write-Warning "No Dataverse instance URL was found for '$environmentName'; agents cannot be enumerated."
            }
            else
            {
                $dataverseResult = Get-DataverseAgents -InstanceApiUrl $instanceUrl

                if (-not $dataverseResult.Succeeded)
                {
                    # Inventory is admin-scoped and sees into environments the caller is not a member
                    # of, so it already covers this environment. Only flag a coverage gap when
                    # inventory did not supply anything for it.
                    $inventoryCovered = ($AgentSource -eq 'Inventory' -or $AgentSource -eq 'All') -and
                                        $inventoryByEnvironment.ContainsKey("$environmentId".ToLowerInvariant())

                    if ($inventoryCovered)
                    {
                        Write-Host ("  [info]  Dataverse is not readable here, but the admin inventory covers this environment.") -ForegroundColor DarkYellow
                    }
                    else
                    {
                        $discoveryFailed = $true
                        $discoveryError = $dataverseResult.Error

                        if ($dataverseResult.StatusCode -eq 403)
                        {
                            $discoveryError = "$($dataverseResult.Error). The admin inventory did not cover it either, so its agents are unknown."
                        }

                        Write-Host ("  [WARN]  Agents could not be enumerated in '{0}': {1}" -f $environmentName, $discoveryError) -ForegroundColor Red
                    }
                }

                foreach ($bot in $dataverseResult.Agents)
                {
                    $botId = "$($bot.botid)"

                    if (-not $botId)
                    {
                        continue
                    }

                    if ($agents.ContainsKey($botId))
                    {
                        $agents[$botId].DisplayName = "$($bot.name)"
                        $agents[$botId].Source = "$($agents[$botId].Source)+Dataverse"
                    }
                    else
                    {
                        $agents[$botId] = [PSCustomObject]@{
                            ResourceId   = $botId
                            DisplayName  = "$($bot.name)"
                            Consumed     = $null
                            Source       = 'Dataverse'
                            ResourceType = 'microsoft.copilotstudio/agents'
                        }
                    }
                }
            }
        }

        if ($agents.Count -eq 0)
        {
            if ($discoveryFailed)
            {
                $script:NotInspectedEnvironments.Add([PSCustomObject]@{
                    EnvironmentName = $environmentName
                    EnvironmentId   = $environmentId
                    Reason          = $discoveryError
                }) | Out-Null

                $action = 'NotInspected'
                $message = "Agents could NOT be enumerated ($discoveryError). This environment's agents are UNKNOWN - it is not confirmed to be empty."
            }
            else
            {
                Write-Detail 'No agents were found in this environment.' -ForegroundColor Yellow
                $action = 'NoAgentsFound'
                $message = 'No agents exist in this environment.'
            }

            $Report.Add([PSCustomObject]@{
                Timestamp             = (Get-Date).ToString('s')
                Scope                 = 'Agent'
                EnvironmentName       = $environmentName
                EnvironmentId         = $environmentId
                TargetId              = ''
                TargetName            = ''
                Source                = ''
                ResourceType          = ''
                Action                = $action
                PreviousLimit         = ''
                NewLimit              = ''
                NotificationThresholdPct = ''
                StopIfOverCapacity    = ''
                Consumed              = ''
                Message               = $message
            }) | Out-Null

            continue
        }

        if ($discoveryFailed)
        {
            # Partial visibility: licensing returned some agents but Dataverse could not be read, so
            # the list may be incomplete.
            $script:NotInspectedEnvironments.Add([PSCustomObject]@{
                EnvironmentName = $environmentName
                EnvironmentId   = $environmentId
                Reason          = "$discoveryError (partial list: $($agents.Count) agent(s) found from licensing only)"
            }) | Out-Null
        }

        $agentCount = @($agents.Values | Where-Object { -not ($_.ResourceType -like '*powerautomate*') }).Count
        $flowCount = @($agents.Values | Where-Object { $_.ResourceType -like '*powerautomate*' }).Count

        if ($flowCount -gt 0)
        {
            Write-Detail ("Discovered: {0} agent(s), {1} flow(s)" -f $agentCount, $flowCount) -ForegroundColor Green
        }
        else
        {
            Write-Detail "Agents discovered: $($agents.Count)" -ForegroundColor Green
        }

        foreach ($agent in ($agents.Values | Sort-Object DisplayName, ResourceId))
        {
            $existing = $thresholdMap[(Get-ThresholdKey -EnvironmentId $environmentId -ResourceId $agent.ResourceId)]
            $previousLimit = $null

            if ($existing)
            {
                $previousLimit = Get-PropertyValue -InputObject $existing -Paths @('limit', 'Limit')
            }

            $row = [ordered]@{
                Timestamp             = (Get-Date).ToString('s')
                Scope                 = 'Agent'
                EnvironmentName       = $environmentName
                EnvironmentId         = $environmentId
                TargetId              = $agent.ResourceId
                TargetName            = $agent.DisplayName
                Source                = $agent.Source
                ResourceType          = $agent.ResourceType
                Action                = ''
                PreviousLimit         = $previousLimit
                NewLimit              = ''
                NotificationThresholdPct = ''
                StopIfOverCapacity    = ''
                Consumed              = $agent.Consumed
                Message               = ''
            }

            if ($Exclusions.Contains($agent.ResourceId) -or ($agent.DisplayName -and $Exclusions.Contains($agent.DisplayName)))
            {
                if ($Exclusions.Contains($agent.ResourceId)) { $script:MatchedExclusions.Add($agent.ResourceId) | Out-Null }
                if ($agent.DisplayName -and $Exclusions.Contains($agent.DisplayName)) { $script:MatchedExclusions.Add($agent.DisplayName) | Out-Null }

                $row.Action = 'Skipped-Excluded'
                $row.Message = 'Listed in the agent exclusions.'
                Write-Detail ("  [skip]  {0} ({1}) - excluded" -f $agent.DisplayName, $agent.ResourceId) -ForegroundColor DarkYellow
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            # Flows can consume Copilot Credits and are therefore inventoried, but the per-resource
            # threshold API is only proven for agents. Writing to a flow is opt-in.
            $isFlow = $agent.ResourceType -and $agent.ResourceType -like '*powerautomate*'

            if ($isFlow -and -not $IncludeFlows -and -not $ReportOnly)
            {
                $row.Action = 'Skipped-Flow'
                $row.Message = "Flow ($($agent.ResourceType)). Not written to by default; use -IncludeFlows to target flows."
                Write-Detail ("  [flow]  {0} ({1}) - reported, not limited. Use -IncludeFlows to include it." -f $agent.DisplayName, $agent.ResourceId) -ForegroundColor DarkCyan
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            if ($ReportOnly)
            {
                $currentLimitText = if ($null -ne $previousLimit) { "$previousLimit" } else { 'none' }
                $row.Action = 'ReportOnly'
                $row.Message = 'Inventory only; no write requested.'
                Write-Detail ("  [read]  {0} ({1}) - current limit: {2}" -f $agent.DisplayName, $agent.ResourceId, $currentLimitText)
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            $row.NewLimit = $AgentCreditLimit
            $row.NotificationThresholdPct = $script:EffectiveNotificationThreshold
            $row.StopIfOverCapacity = [bool] $StopAgentIfOverLimit

            if ($null -ne $previousLimit -and [int] $previousLimit -eq $AgentCreditLimit -and -not $Force)
            {
                $row.Action = 'Skipped-NoChange'
                $row.Message = 'The agent already has this limit.'
                Write-Detail ("  [same]  {0} ({1}) - already {2}" -f $agent.DisplayName, $agent.ResourceId, $AgentCreditLimit) -ForegroundColor DarkGray
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            $resourceLabel = if ($isFlow)
            {
                if ($agent.ResourceType -like '*agentflows*') { 'agent flow' } else { 'cloud flow' }
            }
            else
            {
                'agent'
            }

            $target = "$resourceLabel '$($agent.DisplayName)' ($($agent.ResourceId)) in environment '$environmentName'"
            $operation = "Set Copilot Credit limit to $AgentCreditLimit"

            if (-not $PSCmdlet.ShouldProcess($target, $operation))
            {
                $row.Action = 'WhatIf'
                $row.Message = 'No change written (WhatIf).'
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            try
            {
                Set-AgentThreshold -TargetEnvironmentId $environmentId `
                                   -ResourceId $agent.ResourceId `
                                   -Limit $AgentCreditLimit `
                                   -NotificationThreshold $script:EffectiveNotificationThreshold `
                                   -NotifyOverCapacity ([bool] $NotifyIfOverCapacity) `
                                   -StopOverCapacity ([bool] $StopAgentIfOverLimit) | Out-Null

                $row.Action = 'Set'
                $row.Message = 'Limit applied.'
                Write-Detail ("  [set]   {0} ({1}) -> {2} credits" -f $agent.DisplayName, $agent.ResourceId, $AgentCreditLimit) -ForegroundColor Green
            }
            catch
            {
                $row.Action = 'Failed'
                $row.Message = $_.Exception.Message
                Write-Detail ("  [fail]  {0} ({1}) - {2}" -f $agent.DisplayName, $agent.ResourceId, $_.Exception.Message) -ForegroundColor Red
            }

            $Report.Add([PSCustomObject]$row) | Out-Null
        }
    }
}

#endregion

#region Discover

function Invoke-Discover
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object[]] $Environments
    )

    $environmentObject = $Environments[0]
    $environmentId = Get-EnvironmentId -EnvironmentObject $environmentObject
    $environmentName = Get-EnvironmentDisplayName -EnvironmentObject $environmentObject
    $instanceUrl = Get-EnvironmentInstanceApiUrl -EnvironmentObject $environmentObject

    Write-Section "Discover: $environmentName ($environmentId)"

    if ($instanceUrl)
    {
        Write-Host "Dataverse instance URL: $instanceUrl"
    }
    else
    {
        Write-Host 'Dataverse instance URL: (not found)'
    }

    Write-Section 'GET licensing/environments/{envId}/entitlements/MCSMessages'

    try
    {
        Invoke-PowerPlatformApi -Method GET -PathOrUri ('licensing/environments/{0}/entitlements/{1}' -f $environmentId, $EntitlementId) |
            ConvertTo-Json -Depth 8 |
            Write-Host
    }
    catch
    {
        Write-Warning $_.Exception.Message
    }

    Write-Section 'Licensing resources (agents with consumption)'
    $resources = @(Get-LicensingAgents -EnvironmentId $environmentId)
    Write-Host "Rows: $($resources.Count)"
    $resources | Select-Object -First 5 | ConvertTo-Json -Depth 8 | Write-Host

    Write-Section 'GET licensing/entitlements/MCSMessages/resourceThresholds'
    $thresholds = Get-ResourceThresholdMap
    Write-Host "Existing thresholds: $($thresholds.Count)"
    $thresholds.Values | Select-Object -First 5 | ConvertTo-Json -Depth 8 | Write-Host

    if ($instanceUrl)
    {
        Write-Section 'Dataverse bots (all agents)'
        $dataverseResult = Get-DataverseAgents -InstanceApiUrl $instanceUrl

        if ($dataverseResult.Succeeded)
        {
            $bots = @($dataverseResult.Agents)
            Write-Host "Agents: $($bots.Count)"
            $bots | Select-Object -First 10 -Property botid, name, schemaname, statecode | Format-Table -AutoSize | Out-Host
        }
        else
        {
            Write-Host "Agents could NOT be enumerated: $($dataverseResult.Error)" -ForegroundColor Red
        }
    }

    Write-Section 'Existing thresholds for this environment'
    $envThresholds = $thresholds.Values | Where-Object {
        "$(Get-PropertyValue -InputObject $_ -Paths @('environmentId', 'EnvironmentId'))" -eq $environmentId
    }
    Write-Host ("Rows for this environment: {0}" -f @($envThresholds).Count)
    $envThresholds | Select-Object -First 10 | ConvertTo-Json -Depth 8 | Write-Host
}

#endregion

#region Main

try
{
    if (-not $Discover -and -not $ReportOnly -and -not $PSBoundParameters.ContainsKey('AgentCreditLimit'))
    {
        throw 'Specify -AgentCreditLimit, or use -ReportOnly / -Discover for a read-only pass.'
    }

    # notificationThreshold is a percentage per the Power Platform licensing API contract.
    $script:EffectiveNotificationThreshold = $AgentNotificationThreshold

    # Parse exclusion files BEFORE any network call so a bad file fails immediately, rather than
    # after sign-in and environment enumeration.
    $agentExclusions = Get-ExclusionSet -InlineValues $ExcludeAgentId -CsvPath $ExcludeAgentIdCsv -ColumnCandidates @('AgentId', 'ResourceId', 'BotId', 'Id') -Label 'agent'

    $environments = @(Resolve-TargetEnvironments)

    if ($environments.Count -eq 0)
    {
        throw 'No target environments were resolved. Rerun with -Verbose to see which lookups were attempted.'
    }

    Write-Host "Target environments: $($environments.Count)" -ForegroundColor Green

    if ($Diagnostics)
    {
        $environments |
            Select-Object @{ Name = 'EnvironmentName'; Expression = { Get-EnvironmentDisplayName -EnvironmentObject $_ } },
                          @{ Name = 'EnvironmentId'; Expression = { Get-EnvironmentId -EnvironmentObject $_ } },
                          @{ Name = 'InstanceApiUrl'; Expression = { Get-EnvironmentInstanceApiUrl -EnvironmentObject $_ } } |
            Format-Table -AutoSize |
            Out-Host
    }

    if ($Discover)
    {
        Invoke-Discover -Environments $environments
        return
    }

    if ($agentExclusions.Count -gt 0)
    {
        Write-Host "Agent exclusions: $($agentExclusions.Count)" -ForegroundColor Yellow
    }

    if (-not $ReportOnly -and -not $Force -and -not $WhatIfPreference)
    {
        $question = "Apply a Copilot Credit limit of $AgentCreditLimit to every agent in $($environments.Count) environment(s)?"
        $answer = Read-Host "$question [y/N]"

        if ($answer -notmatch '^(y|yes)$')
        {
            Write-Host 'Cancelled.' -ForegroundColor Yellow
            return
        }
    }

    $report = New-Object System.Collections.Generic.List[object]
    $script:NotInspectedEnvironments = New-Object System.Collections.Generic.List[object]

    # Verify every exclusion matches a real resource BEFORE writing anything. An exclusion that
    # matches nothing - a typo, a stale ID, an unsubstituted placeholder - would silently leave
    # the resource it was meant to protect exposed to the write.
        if ($agentExclusions.Count -gt 0)
        {
            Write-Host ''
            Write-Host 'Verifying exclusions against discovered resources...' -ForegroundColor Cyan

            $preflight = New-Object System.Collections.Generic.List[object]
            $previousReportOnly = $ReportOnly
            $ReportOnly = $true
            $script:SuppressSectionOutput = $true

            Invoke-AgentScope -Environments $environments -Exclusions $agentExclusions -Report $preflight 6>$null | Out-Null

            $ReportOnly = $previousReportOnly
            $script:SuppressSectionOutput = $false

            $unmatched = @($agentExclusions | Where-Object { -not $script:MatchedExclusions.Contains($_) })

            if ($unmatched.Count -gt 0)
            {
                Write-Host ''
                Write-Host 'EXCLUSIONS THAT MATCHED NOTHING' -ForegroundColor Red
                Write-Host '-------------------------------' -ForegroundColor Red
                $unmatched | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
                Write-Host ''
                Write-Host 'These do not correspond to any discovered agent or flow, so they protect nothing.' -ForegroundColor Red
                Write-Host 'Check for typos or stale IDs. Anything you intended to exclude WILL be written to.' -ForegroundColor Red

                if (-not $Force -and -not $WhatIfPreference)
                {
                    $answer = Read-Host 'Continue anyway? [y/N]'

                    if ($answer -notmatch '^(y|yes)$')
                    {
                        Write-Host 'Cancelled.' -ForegroundColor Yellow
                        return
                    }
                }
            }
            else
            {
                Write-Host ("All {0} exclusion(s) matched a discovered resource." -f $agentExclusions.Count) -ForegroundColor Green
            }

            $script:MatchedExclusions.Clear()
        }

    Invoke-AgentScope -Environments $environments -Exclusions $agentExclusions -Report $report

    Write-Section 'Summary'

    if ($report.Count -eq 0)
    {
        Write-Host 'Nothing to report.' -ForegroundColor Yellow
        return
    }

    $report |
        Group-Object Scope, Action |
        Sort-Object Name |
        ForEach-Object { Write-Host ("  {0,-45} {1}" -f $_.Name, $_.Count) }

    if ($script:NotInspectedEnvironments.Count -gt 0)
    {
        Write-Host ''
        Write-Host 'INCOMPLETE COVERAGE' -ForegroundColor Red
        Write-Host '-------------------' -ForegroundColor Red
        Write-Host ("{0} environment(s) could not be fully inspected. Their agents are UNKNOWN," -f $script:NotInspectedEnvironments.Count) -ForegroundColor Red
        Write-Host 'and no limit was applied to them. Do not read this run as confirming they are clean.' -ForegroundColor Red
        Write-Host ''

        $script:NotInspectedEnvironments |
            Select-Object EnvironmentName, EnvironmentId, Reason |
            Format-Table -AutoSize |
            Out-Host

        Write-Host 'Agent discovery reads the Dataverse bots table, which requires the calling account to be' -ForegroundColor Yellow
        Write-Host 'a member of each environment. The Power Platform administrator role is not sufficient on' -ForegroundColor Yellow
        Write-Host 'its own - this commonly affects personal developer environments, and the Copilot Studio' -ForegroundColor Yellow
        Write-Host 'portal fails the same way for an administrator who is not a member.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host 'For environments you choose not to enter, the effective control is the environment-group' -ForegroundColor Yellow
        Write-Host 'rule that disables drawing from the tenant pool, combined with a zero credit allocation.' -ForegroundColor Yellow
    }

    if (-not (Test-Path -LiteralPath $ReportPath))
    {
        New-Item -ItemType Directory -Path $ReportPath -Force -WhatIf:$false | Out-Null
    }

    $reportFile = Join-Path -Path $ReportPath -ChildPath ('CopilotCreditLimits-{0}.csv' -f (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $report | Export-Csv -LiteralPath $reportFile -NoTypeInformation -Encoding UTF8 -WhatIf:$false

    Write-Host ''
    Write-Host "Report written to $reportFile" -ForegroundColor Green
}
catch
{
    Write-Error $_
    exit 1
}

#endregion

