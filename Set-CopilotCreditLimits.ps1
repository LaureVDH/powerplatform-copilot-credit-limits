#requires -Version 5.1
#requires -Modules Az.Accounts

<#
.SYNOPSIS
Applies Copilot Credit (MCSMessages) consumption limits to every agent in one or more Power Platform
environments, or in every environment of an environment group, and reports what was done.

.DESCRIPTION
Connects with an interactive Power Platform admin account through Az.Accounts, resolves the target
environments (single environment, list of environments, or all environments of an environment group),
inventories the Copilot Studio agents in each environment, and writes a per-agent Copilot Credit
threshold through the Power Platform licensing API:

    PUT licensing/environments/{environmentId}/entitlements/MCSMessages/resources/{resourceId}/threshold

Agents are discovered from two complementary sources:

  * Licensing  - agents that already have Copilot Credit consumption records. These resource IDs are
                 guaranteed to be valid targets for the threshold API.
  * Dataverse  - the 'bots' table of the environment. Adds display names and surfaces agents that have
                 never consumed a credit yet.

Users (-Scope Users or Both) are inventoried and reported, but not written to. The Power Platform
licensing API exposes user data as read-only operations (tenant user snapshots and consumption by
user); there is no supported per-user credit limit write operation. Per-user spending limits are
configured in the Microsoft 365 admin center under Copilot > Cost Management (spending policies),
which has no public REST API today. Use -ProbeUserThresholdApi to test whether a user-scoped threshold
endpoint has appeared in your tenant.

.PARAMETER EnvironmentGroup
Environment group ID or exact display name. Every environment of the group is targeted.

.PARAMETER Environment
One or more environment IDs or exact environment display names.

.PARAMETER Scope
What to process: Agents (default), Users, or Both.

.PARAMETER AgentCreditLimit
Monthly Copilot Credit limit to apply to each agent. Required when the scope includes agents, unless
-Discover or -ReportOnly is used.

.PARAMETER AgentNotificationThreshold
Percentage (1-100) of the agent limit at which administrators are notified. Defaults to 80.
The Power Platform API defines this field as a percentage, not an absolute credit count.

.PARAMETER NotifyIfOverCapacity
Send a notification when the agent exceeds its limit. Defaults to on; use -NotifyIfOverCapacity:$false to disable.

.PARAMETER StopAgentIfOverLimit
Turn the agent off when it reaches its limit (hard stop).

.PARAMETER UserCreditLimit
Per-user Copilot Credit limit to report against. Recorded in the report only; see the description.

.PARAMETER UserGroup
Entra ID group object ID or display name. When supplied, the user report is scoped to and labelled
with that group instead of listing every user of the environment.

.PARAMETER ExcludeAgentId
One or more agent/resource IDs to skip.

.PARAMETER ExcludeAgentIdCsv
Path to a CSV of agent/resource IDs to skip. A header named AgentId, ResourceId, BotId or Id is
honoured; a headerless single-column file also works.

.PARAMETER ExcludeUserId
One or more user IDs (Entra object ID, Dataverse systemuserid, or UPN) to skip in the user report.

.PARAMETER ExcludeUserIdCsv
Path to a CSV of user IDs to skip. Header UserId, ObjectId, Upn, Email or Id is honoured.

.PARAMETER AgentSource
Where to discover agents: Licensing, Dataverse, or Both (default).

.PARAMETER LookbackDays
How far back the licensing consumption snapshot is queried. Default 90.

.PARAMETER ReportPath
Folder for the CSV report. Default: .\reports

.PARAMETER Discover
Read-only mode: dump the raw API responses for the first target environment so you can verify the
shape of these preview APIs in your tenant. Nothing is written.

.PARAMETER ReportOnly
Inventory and report current limits without writing anything.

.PARAMETER ProbeUserThresholdApi
Test whether a user-scoped threshold endpoint exists in your tenant. Read-only.

.PARAMETER TenantId
Optional tenant ID to pass to Connect-AzAccount.

.PARAMETER Force
Skip the confirmation prompt and rewrite limits even when the current value already matches.

.PARAMETER Diagnostics
Show extra detail about discovery and API fallbacks.

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -Environment "Contoso Dev" -Discover

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -Environment "00000000-0000-0000-0000-000000000000" -AgentCreditLimit 1000 -WhatIf

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" -AgentCreditLimit 500 -StopAgentIfOverLimit -ExcludeAgentIdCsv .\samples\agent-exclusions.csv -Force

.EXAMPLE
.\Set-CopilotCreditLimits.ps1 -EnvironmentGroup "Personal Productivity" -Scope Both -AgentCreditLimit 500 -UserCreditLimit 100 -UserGroup "Copilot Makers"
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
    [ValidateSet('Agents', 'Users', 'Both')]
    [string] $Scope = 'Agents',

    [Parameter()]
    [ValidateRange(0, 2147483647)]
    [int] $AgentCreditLimit,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int] $AgentNotificationThreshold = 80,

    [Parameter()]
    [switch] $NotifyIfOverCapacity = $true,

    [Parameter()]
    [switch] $StopAgentIfOverLimit,

    [Parameter()]
    [ValidateRange(0, 2147483647)]
    [int] $UserCreditLimit,

    [Parameter()]
    [string] $UserGroup,

    [Parameter()]
    [string[]] $ExcludeAgentId,

    [Parameter()]
    [string] $ExcludeAgentIdCsv,

    [Parameter()]
    [string[]] $ExcludeUserId,

    [Parameter()]
    [string] $ExcludeUserIdCsv,

    [Parameter()]
    [ValidateSet('Licensing', 'Dataverse', 'Both')]
    [string] $AgentSource = 'Both',

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
    [switch] $ProbeUserThresholdApi,

    [Parameter()]
    [string] $ProbeUserId,

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

    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * $Text.Length) -ForegroundColor DarkCyan
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

    if ($PSCmdlet.ParameterSetName -eq 'Group')
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
            throw "No environments were found in environment group '$groupName'."
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
                $sawForbidden = $true
            }

            Write-Verbose "Licensing resource path failed: '$path'. $($_.Exception.Message)"
        }
    }

    if ($sawForbidden)
    {
        Write-Host ("  [info]  Per-agent consumption is not readable for this environment (HTTP 403). " +
                    "Agents are still discovered from Dataverse and limits can still be written; only the " +
                    "month-to-date 'Consumed' column will be blank.") -ForegroundColor DarkYellow
        return @()
    }

    Write-Warning "Could not read licensing resources for environment $EnvironmentId. Agent discovery falls back to Dataverse."
    return @()
}

function Get-DataverseAgents
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $InstanceApiUrl
    )

    $query = 'bots?$select=botid,name,schemaname,statecode,createdon,_ownerid_value&$orderby=name asc'

    try
    {
        return @(Invoke-DataverseApi -InstanceApiUrl $InstanceApiUrl -Path $query)
    }
    catch
    {
        Write-Warning "Could not read the Dataverse 'bots' table at $InstanceApiUrl. $($_.Exception.Message)"
        return @()
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
        $thresholds = @(Invoke-PowerPlatformApi -Method GET -PathOrUri ('licensing/entitlements/{0}/resourceThresholds' -f $EntitlementId))
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

    return Invoke-PowerPlatformApi -Method PUT -PathOrUri $path -Body $body
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

    foreach ($environmentObject in $Environments)
    {
        $environmentId = Get-EnvironmentId -EnvironmentObject $environmentObject
        $environmentName = Get-EnvironmentDisplayName -EnvironmentObject $environmentObject
        $instanceUrl = Get-EnvironmentInstanceApiUrl -EnvironmentObject $environmentObject

        Write-Section "Environment: $environmentName ($environmentId)"

        $agents = @{}

        if ($AgentSource -eq 'Licensing' -or $AgentSource -eq 'Both')
        {
            foreach ($resource in Get-LicensingAgents -EnvironmentId $environmentId)
            {
                $resourceId = "$(Get-PropertyValue -InputObject $resource -Paths @('resourceId', 'ResourceId', 'id'))"

                if (-not $resourceId)
                {
                    continue
                }

                $agents[$resourceId] = [PSCustomObject]@{
                    ResourceId  = $resourceId
                    DisplayName = "$(Get-PropertyValue -InputObject $resource -Paths @('metadata.ProductName', 'metadata.productName', 'metadata.Feature', 'name'))"
                    Consumed    = Get-PropertyValue -InputObject $resource -Paths @('consumed', 'Consumed')
                    Source      = 'Licensing'
                }
            }
        }

        if ($AgentSource -eq 'Dataverse' -or $AgentSource -eq 'Both')
        {
            if (-not $instanceUrl)
            {
                Write-Warning "No Dataverse instance URL was found for '$environmentName'; skipping Dataverse agent discovery."
            }
            else
            {
                foreach ($bot in Get-DataverseAgents -InstanceApiUrl $instanceUrl)
                {
                    $botId = "$($bot.botid)"

                    if (-not $botId)
                    {
                        continue
                    }

                    if ($agents.ContainsKey($botId))
                    {
                        $agents[$botId].DisplayName = "$($bot.name)"
                        $agents[$botId].Source = 'Licensing+Dataverse'
                    }
                    else
                    {
                        $agents[$botId] = [PSCustomObject]@{
                            ResourceId  = $botId
                            DisplayName = "$($bot.name)"
                            Consumed    = $null
                            Source      = 'Dataverse'
                        }
                    }
                }
            }
        }

        if ($agents.Count -eq 0)
        {
            Write-Host 'No agents were found in this environment.' -ForegroundColor Yellow

            $Report.Add([PSCustomObject]@{
                Timestamp             = (Get-Date).ToString('s')
                Scope                 = 'Agent'
                EnvironmentName       = $environmentName
                EnvironmentId         = $environmentId
                TargetId              = ''
                TargetName            = ''
                Source                = ''
                Action                = 'NoAgentsFound'
                PreviousLimit         = ''
                NewLimit              = ''
                NotificationThresholdPct = ''
                StopIfOverCapacity    = ''
                Consumed              = ''
                Message               = 'No agents discovered in this environment.'
            }) | Out-Null

            continue
        }

        Write-Host "Agents discovered: $($agents.Count)" -ForegroundColor Green

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
                $row.Action = 'Skipped-Excluded'
                $row.Message = 'Listed in the agent exclusions.'
                Write-Host ("  [skip]  {0} ({1}) - excluded" -f $agent.DisplayName, $agent.ResourceId) -ForegroundColor DarkYellow
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            if ($ReportOnly)
            {
                $currentLimitText = if ($null -ne $previousLimit) { "$previousLimit" } else { 'none' }
                $row.Action = 'ReportOnly'
                $row.Message = 'Inventory only; no write requested.'
                Write-Host ("  [read]  {0} ({1}) - current limit: {2}" -f $agent.DisplayName, $agent.ResourceId, $currentLimitText)
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
                Write-Host ("  [same]  {0} ({1}) - already {2}" -f $agent.DisplayName, $agent.ResourceId, $AgentCreditLimit) -ForegroundColor DarkGray
                $Report.Add([PSCustomObject]$row) | Out-Null
                continue
            }

            $target = "agent '$($agent.DisplayName)' ($($agent.ResourceId)) in environment '$environmentName'"
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
                Write-Host ("  [set]   {0} ({1}) -> {2} credits" -f $agent.DisplayName, $agent.ResourceId, $AgentCreditLimit) -ForegroundColor Green
            }
            catch
            {
                $row.Action = 'Failed'
                $row.Message = $_.Exception.Message
                Write-Host ("  [fail]  {0} ({1}) - {2}" -f $agent.DisplayName, $agent.ResourceId, $_.Exception.Message) -ForegroundColor Red
            }

            $Report.Add([PSCustomObject]$row) | Out-Null
        }
    }
}

#endregion

#region Users

function Get-EntraGroupMembers
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $GroupIdOrName
    )

    $token = Get-AccessToken -ResourceUrl $GraphApiRoot -TenantId $TenantId
    $headers = @{ Authorization = "Bearer $token"; Accept = 'application/json' }
    $groupId = $GroupIdOrName

    if ($GroupIdOrName -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
    {
        $filter = [uri]::EscapeDataString("displayName eq '$GroupIdOrName'")
        $lookupUri = '{0}/v1.0/groups?$filter={1}&$select=id,displayName' -f $GraphApiRoot, $filter
        $lookup = Invoke-RestMethod -Method GET -Uri $lookupUri -Headers $headers -ErrorAction Stop

        if (@($lookup.value).Count -eq 0)
        {
            throw "Entra group '$GroupIdOrName' was not found."
        }

        if (@($lookup.value).Count -gt 1)
        {
            throw "Entra group name '$GroupIdOrName' matched more than one group. Use the group object ID."
        }

        $groupId = $lookup.value[0].id
    }

    $members = New-Object System.Collections.Generic.List[object]
    $uri = '{0}/v1.0/groups/{1}/transitiveMembers/microsoft.graph.user?$select=id,displayName,userPrincipalName&$top=999' -f $GraphApiRoot, $groupId

    do
    {
        $response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -ErrorAction Stop

        foreach ($member in @($response.value))
        {
            $members.Add($member) | Out-Null
        }

        $uri = $response.'@odata.nextLink'
    }
    while ($uri)

    return [PSCustomObject]@{
        GroupId = $groupId
        Members = $members.ToArray()
    }
}

function Get-EnvironmentUsers
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $InstanceApiUrl
    )

    $makerIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    try
    {
        $query = 'systemusers?$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid,isdisabled&$filter=isdisabled eq false and applicationid eq null&$orderby=fullname asc'
        $users = @(Invoke-DataverseApi -InstanceApiUrl $InstanceApiUrl -Path $query)
    }
    catch
    {
        Write-Warning "Could not read Dataverse users at $InstanceApiUrl. $($_.Exception.Message)"
        return @()
    }

    try
    {
        $makerRoles = @(Invoke-DataverseApi -InstanceApiUrl $InstanceApiUrl -Path 'roles?$select=roleid,name&$filter=name eq ''Environment Maker''')

        foreach ($role in $makerRoles)
        {
            $roleQuery = 'systemusers?$select=systemuserid&$filter=systemuserroles_association/any(r:r/roleid eq {0})' -f $role.roleid
            $roleUsers = @(Invoke-DataverseApi -InstanceApiUrl $InstanceApiUrl -Path $roleQuery)

            foreach ($roleUser in $roleUsers)
            {
                $makerIds.Add("$($roleUser.systemuserid)") | Out-Null
            }
        }
    }
    catch
    {
        Write-Verbose "Could not resolve Environment Maker role membership. $($_.Exception.Message)"
    }

    return @($users | ForEach-Object {
        [PSCustomObject]@{
            SystemUserId = "$($_.systemuserid)"
            ObjectId     = "$($_.azureactivedirectoryobjectid)"
            DisplayName  = "$($_.fullname)"
            Upn          = "$($_.internalemailaddress)"
            IsMaker      = $makerIds.Contains("$($_.systemuserid)")
        }
    })
}

function Get-LicensingUserConsumption
{
    [CmdletBinding()]
    param()

    $fromDate = (Get-Date).AddDays(-$LookbackDays).ToString('yyyy-MM-dd')
    $toDate = (Get-Date).ToString('yyyy-MM-dd')
    $map = @{}

    foreach ($template in @(
        'licensing/entitlements/{0}/users?fromDate={1}&toDate={2}',
        'licensing/entitlements/{0}/users?from-date={1}&to-date={2}&page-size=200'
    ))
    {
        $path = $template -f $EntitlementId, $fromDate, $toDate

        try
        {
            $rows = @(Invoke-PowerPlatformApi -Method GET -PathOrUri $path -AllPages)

            foreach ($row in $rows)
            {
                $userRows = if ($row.users) { @($row.users) } else { @($row) }

                foreach ($userRow in $userRows)
                {
                    $userId = "$(Get-PropertyValue -InputObject $userRow -Paths @('userId', 'UserId', 'id', 'objectId'))"

                    if ($userId)
                    {
                        $map[$userId] = $userRow
                    }
                }
            }

            Write-Verbose "Loaded $($map.Count) user consumption row(s) from '$path'."
            return $map
        }
        catch
        {
            Write-Verbose "User consumption path failed: '$path'. $($_.Exception.Message)"
        }
    }

    Write-Warning 'Could not read tenant user consumption for MCSMessages.'
    return $map
}

function Test-UserThresholdApi
{
    <#
        MC1451872 (public preview 22-Aug-2026) announced per-user capacity limits INSIDE a Power
        Platform environment, alongside the existing per-agent limits. That capability is not present
        in the published licensing OpenAPI spec (2024-10-01), which exposes only GET operations for
        users plus the single resource-threshold PUT.

        This probe therefore tests, read-only, whether a user-scoped threshold endpoint has shipped in
        the caller's tenant. Candidates are built by symmetry with the documented resource shapes:

            per resource : licensing/environments/{env}/entitlements/{ent}/resources/{id}/threshold
            collection   : licensing/entitlements/{ent}/resourceThresholds

        404 means the route does not exist. 403 means the route EXISTS but the caller lacks permission,
        which is still a positive signal. Any 2xx is a hit.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string] $TargetEnvironmentId,

        [Parameter(Mandatory = $true)]
        [string] $UserId
    )

    $candidates = @(
        # Routes mirroring the PPAC Users page itself:
        #   admin.preview.powerplatform.microsoft.com/billing/licenses/agents/CopilotStudio/users
        'licensing/entitlements/{1}/users/{2}',
        'licensing/entitlements/{1}/users/{2}/threshold',
        'licensing/entitlements/{1}/userThresholds/{2}',
        'licensing/entitlements/{1}/userThresholds',
        'licensing/entitlements/{1}/users/{2}/limit',
        'licensing/entitlements/{1}/users/{2}/capacity',

        # Environment-scoped variants, mirroring the per-resource threshold route.
        'licensing/environments/{0}/entitlements/{1}/users/{2}/threshold',
        'licensing/environments/{0}/entitlements/{1}/users/{2}/userThreshold',
        'licensing/environments/{0}/entitlements/{1}/userThresholds/{2}',
        'licensing/environments/{0}/entitlements/{1}/userThresholds',
        'licensing/environments/{0}/entitlements/{1}/users/{2}/capacity',
        'licensing/environments/{0}/entitlements/{1}/users/{2}/limit',
        'licensing/environments/{0}/entitlements/{1}/userCapacityLimits',
        'licensing/entitlements/{1}/environments/{0}/users/{2}/threshold',

        # The agents/CopilotStudio wording used by the preview portal URL.
        'licensing/agents/CopilotStudio/users/{2}/threshold',
        'licensing/agents/{1}/users/{2}/threshold'
    )

    # The PPAC preview portal may serve this page from the BAP host rather than api.powerplatform.com.
    # Probing absolute BAP URLs tests that hypothesis directly.
    $bapCandidates = @(
        'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{0}/licensing/users/{2}/threshold?api-version=2020-10-01',
        'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/licensing/entitlements/{1}/users/{2}/threshold?api-version=2020-10-01',
        'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/licensing/entitlements/{1}/userThresholds?api-version=2020-10-01'
    )

    $found = New-Object System.Collections.Generic.List[string]

    Write-Host ''
    Write-Host '  Probing for a per-user threshold endpoint (MC1451872 preview)...' -ForegroundColor Cyan

    foreach ($template in $candidates)
    {
        $path = $template -f $TargetEnvironmentId, $EntitlementId, $UserId

        try
        {
            $response = Invoke-PowerPlatformApi -Method GET -PathOrUri $path
            Write-Host "  [HIT]   $path" -ForegroundColor Green
            $found.Add($path) | Out-Null

            if ($Diagnostics)
            {
                $response | ConvertTo-Json -Depth 8 | Write-Host
            }
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
                Write-Host "  [403]   $path - ROUTE EXISTS but access denied. Worth following up." -ForegroundColor Yellow
                $found.Add("$path (403 - exists, access denied)") | Out-Null
            }
            elseif ($statusCode -eq 400)
            {
                Write-Host "  [400]   $path (route may exist; parameters rejected)" -ForegroundColor DarkYellow
                $found.Add("$path (400 - route may exist)") | Out-Null
            }
            elseif ($statusCode -eq 404)
            {
                Write-Host "  [404]   $path" -ForegroundColor DarkGray
            }
            else
            {
                Write-Host "  [$statusCode]   $path" -ForegroundColor DarkGray
            }
        }
    }

    foreach ($template in $bapCandidates)
    {
        $uri = $template -f $TargetEnvironmentId, $EntitlementId, $UserId
        $shortUri = $uri -replace '^https://api\.bap\.microsoft\.com/providers/Microsoft\.BusinessAppPlatform/scopes/admin/', 'BAP:'

        try
        {
            $token = Get-AccessToken -ResourceUrl 'https://api.bap.microsoft.com' -TenantId $TenantId
            $response = Invoke-RestMethod -Method GET -Uri $uri -Headers @{ Authorization = "Bearer $token"; Accept = 'application/json' } -ErrorAction Stop

            Write-Host "  [HIT]   $shortUri" -ForegroundColor Green
            $found.Add($uri) | Out-Null

            if ($Diagnostics)
            {
                $response | ConvertTo-Json -Depth 8 | Write-Host
            }
        }
        catch
        {
            $statusCode = $null

            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response)
            {
                $statusCode = [int] $_.Exception.Response.StatusCode
            }

            if ($statusCode -eq 403 -or $statusCode -eq 400)
            {
                Write-Host "  [$statusCode]   $shortUri - route may exist. Worth following up." -ForegroundColor Yellow
                $found.Add("$uri ($statusCode)") | Out-Null
            }
            else
            {
                Write-Host "  [$statusCode]   $shortUri" -ForegroundColor DarkGray
            }
        }
    }

    Write-Host ''

    if ($found.Count -gt 0)
    {
        Write-Host '  A user-scoped endpoint responded. Per-user limits may be programmable here:' -ForegroundColor Green
        $found | ForEach-Object { Write-Host "    $_" -ForegroundColor Green }
        Write-Host '  Share this output so the write path can be implemented.' -ForegroundColor Green
    }
    else
    {
        Write-Host '  No user-scoped threshold endpoint responded in this tenant.' -ForegroundColor Yellow
        Write-Host '  Per-user limits remain configurable through the admin centers only:' -ForegroundColor Yellow
        Write-Host '    - Power Platform admin center (MC1451872 preview, per environment), or' -ForegroundColor Yellow
        Write-Host '    - Microsoft 365 admin center > Copilot > Cost Management spending policies.' -ForegroundColor Yellow
    }

    return $found.ToArray()
}

function Invoke-UserScope
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

    Write-Section 'Users'
    Write-Host 'There are two separate Copilot Credit limit systems. This matters:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  1. Copilot Studio / Power Platform credits (entitlement MCSMessages) - what this script writes.' -ForegroundColor Yellow
    Write-Host '     Limits here apply to AGENTS ONLY. There is no per-user limit in this system, and the' -ForegroundColor Yellow
    Write-Host '     licensing API exposes /users endpoints as GET only (verified against the official' -ForegroundColor Yellow
    Write-Host '     licensing OpenAPI spec: the single PUT in the namespace is the resource threshold).' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  2. Microsoft 365 Cost Management spending policies - where per-user monthly caps DO exist' -ForegroundColor Yellow
    Write-Host '     (M365 admin center > Copilot > Cost Management > Spending policies). These are HARD' -ForegroundColor Yellow
    Write-Host '     limits: the service stops for that user when the cap is reached. Policies are scoped to' -ForegroundColor Yellow
    Write-Host '     an ENTRA GROUP (or the tenant); scoping a policy to a single user is not yet supported,' -ForegroundColor Yellow
    Write-Host '     so the pattern is: Entra group -> spending policy -> per-user monthly limit.' -ForegroundColor Yellow
    Write-Host '     Minimum per-user limit is 2,000 credits/user/month (7,000 recommended). Enforcement' -ForegroundColor Yellow
    Write-Host '     reconciles periodically, so a user can briefly exceed the cap before access is cut;' -ForegroundColor Yellow
    Write-Host '     that overage is not billed. This surface has no public REST API today.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '     Note: the older PAYG billing-policy "budget" only sends alerts. It does NOT stop usage.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Run this pass with -UserGroup to get the per-group user list to build those policies against.' -ForegroundColor Yellow

    $groupMembers = $null
    $groupLabel = ''

    if ($UserGroup)
    {
        try
        {
            $groupResult = Get-EntraGroupMembers -GroupIdOrName $UserGroup
            $groupLabel = $UserGroup
            $groupMembers = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

            foreach ($member in $groupResult.Members)
            {
                $groupMembers.Add("$($member.id)") | Out-Null

                if ($member.userPrincipalName)
                {
                    $groupMembers.Add("$($member.userPrincipalName)") | Out-Null
                }
            }

            Write-Host "User group '$UserGroup' ($($groupResult.GroupId)): $($groupResult.Members.Count) member(s)." -ForegroundColor Green
        }
        catch
        {
            Write-Warning "Could not expand the Entra group '$UserGroup'. $($_.Exception.Message)"
        }
    }

    $consumption = Get-LicensingUserConsumption

    foreach ($environmentObject in $Environments)
    {
        $environmentId = Get-EnvironmentId -EnvironmentObject $environmentObject
        $environmentName = Get-EnvironmentDisplayName -EnvironmentObject $environmentObject
        $instanceUrl = Get-EnvironmentInstanceApiUrl -EnvironmentObject $environmentObject

        Write-Host ''
        Write-Host "Environment: $environmentName ($environmentId)" -ForegroundColor Cyan

        if (-not $instanceUrl)
        {
            Write-Warning 'No Dataverse instance URL; skipping the user inventory for this environment.'
            continue
        }

        $users = @(Get-EnvironmentUsers -InstanceApiUrl $instanceUrl)
        $makerCount = @($users | Where-Object { $_.IsMaker }).Count
        Write-Host ("Users: {0} (makers: {1})" -f $users.Count, $makerCount) -ForegroundColor Green

        if ($ProbeUserThresholdApi)
        {
            # Probe with a user the LICENSING API actually knows. A Dataverse user with no Copilot
            # Credit consumption may be unknown to licensing and return 404 for reasons unrelated to
            # whether the route exists, which would produce a false negative.
            # NOTE: PowerShell variable names are case-insensitive, so a local named $probeUserId
            # would be the SAME variable as the $ProbeUserId parameter and would silently overwrite
            # the caller's value. The local is therefore named differently on purpose.
            $resolvedProbeUser = $null
            $probeUserSource = $null

            if ($ProbeUserId)
            {
                $resolvedProbeUser = $ProbeUserId
                $probeUserSource = 'the -ProbeUserId parameter'
            }
            elseif ($consumption.Keys.Count -gt 0)
            {
                # Prefer a user that actually consumed credits, and never probe with an empty GUID.
                $bestKey = $null

                foreach ($key in $consumption.Keys)
                {
                    if ("$key" -eq '00000000-0000-0000-0000-000000000000' -or -not "$key".Trim())
                    {
                        continue
                    }

                    $consumedValue = Get-PropertyValue -InputObject $consumption[$key] -Paths @('consumed', 'Consumed')

                    if ($null -ne $consumedValue -and [double] $consumedValue -gt 0)
                    {
                        $bestKey = $key
                        break
                    }

                    if (-not $bestKey)
                    {
                        $bestKey = $key
                    }
                }

                if ($bestKey)
                {
                    $resolvedProbeUser = $bestKey
                    $probeUserSource = 'the licensing user consumption snapshot'
                }
            }

            if (-not $resolvedProbeUser -and $users.Count -gt 0)
            {
                $consumingUser = $users | Where-Object { $_.ObjectId } | Select-Object -First 1

                if ($consumingUser)
                {
                    $resolvedProbeUser = $consumingUser.ObjectId
                    $probeUserSource = 'Dataverse (no licensing user found - result may be a false negative)'
                }
            }

            if ($resolvedProbeUser)
            {
                Write-Host ''
                Write-Host "  Probe user: $resolvedProbeUser (from $probeUserSource)" -ForegroundColor Cyan
                Test-UserThresholdApi -TargetEnvironmentId $environmentId -UserId $resolvedProbeUser | Out-Null
            }
            else
            {
                Write-Warning 'No user was available to probe with. Pass -ProbeUserId explicitly.'
            }
        }

        foreach ($user in $users)
        {
            $identifiers = @($user.ObjectId, $user.SystemUserId, $user.Upn) | Where-Object { $_ }

            if ($groupMembers)
            {
                $inGroup = @($identifiers | Where-Object { $groupMembers.Contains($_) }).Count -gt 0

                if (-not $inGroup)
                {
                    continue
                }
            }

            $action = 'ReportOnly'
            $message = 'User-level limits require Microsoft 365 Copilot Cost Management spending policies.'

            if (@($identifiers | Where-Object { $Exclusions.Contains($_) }).Count -gt 0)
            {
                $action = 'Skipped-Excluded'
                $message = 'Listed in the user exclusions.'
            }

            $consumed = $null

            foreach ($identifier in $identifiers)
            {
                if ($consumption.ContainsKey($identifier))
                {
                    $consumed = Get-PropertyValue -InputObject $consumption[$identifier] -Paths @('consumed', 'Consumed')
                    break
                }
            }

            $scopeLabel = if ($groupLabel) { "User ($groupLabel)" } else { 'User' }
            $targetId = if ($user.ObjectId) { $user.ObjectId } else { $user.SystemUserId }
            $targetName = if ($user.Upn) { "$($user.DisplayName) <$($user.Upn)>" } else { $user.DisplayName }
            $sourceLabel = if ($user.IsMaker) { 'Dataverse (maker)' } else { 'Dataverse (user)' }
            $plannedLimit = if ($PSBoundParameters.ContainsKey('UserCreditLimit')) { $UserCreditLimit } else { '' }

            $Report.Add([PSCustomObject]@{
                Timestamp             = (Get-Date).ToString('s')
                Scope                 = $scopeLabel
                EnvironmentName       = $environmentName
                EnvironmentId         = $environmentId
                TargetId              = $targetId
                TargetName            = $targetName
                Source                = $sourceLabel
                Action                = $action
                PreviousLimit         = ''
                NewLimit              = $plannedLimit
                NotificationThresholdPct = ''
                StopIfOverCapacity    = ''
                Consumed              = $consumed
                Message               = $message
            }) | Out-Null
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
        $bots = @(Get-DataverseAgents -InstanceApiUrl $instanceUrl)
        Write-Host "Agents: $($bots.Count)"
        $bots | Select-Object -First 10 -Property botid, name, schemaname, statecode | Format-Table -AutoSize | Out-Host
    }

    Write-Section 'Tenant user consumption (read-only)'
    $userConsumption = Get-LicensingUserConsumption
    Write-Host "Users with consumption: $($userConsumption.Count)"
    $userConsumption.Values | Select-Object -First 5 | ConvertTo-Json -Depth 8 | Write-Host
}

#endregion

#region Main

try
{
    $agentScopeRequested = ($Scope -eq 'Agents' -or $Scope -eq 'Both')

    if ($agentScopeRequested -and -not $Discover -and -not $ReportOnly -and -not $PSBoundParameters.ContainsKey('AgentCreditLimit'))
    {
        throw 'Specify -AgentCreditLimit, or use -ReportOnly / -Discover for a read-only pass.'
    }

    # notificationThreshold is a percentage (1-100) per the Power Platform licensing API contract.
    $script:EffectiveNotificationThreshold = $AgentNotificationThreshold

    # Parse exclusion files BEFORE any network call so a bad file fails immediately, rather than
    # after sign-in and environment enumeration.
    $agentExclusions = Get-ExclusionSet -InlineValues $ExcludeAgentId -CsvPath $ExcludeAgentIdCsv -ColumnCandidates @('AgentId', 'ResourceId', 'BotId', 'Id') -Label 'agent'
    $userExclusions = Get-ExclusionSet -InlineValues $ExcludeUserId -CsvPath $ExcludeUserIdCsv -ColumnCandidates @('UserId', 'ObjectId', 'Upn', 'Email', 'Id') -Label 'user'

    $environments = @(Resolve-TargetEnvironments)
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

    if ($userExclusions.Count -gt 0)
    {
        Write-Host "User exclusions: $($userExclusions.Count)" -ForegroundColor Yellow
    }

    if ($agentScopeRequested -and -not $ReportOnly -and -not $Force -and -not $WhatIfPreference)
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

    if ($agentScopeRequested)
    {
        Invoke-AgentScope -Environments $environments -Exclusions $agentExclusions -Report $report
    }

    if ($Scope -eq 'Users' -or $Scope -eq 'Both')
    {
        Invoke-UserScope -Environments $environments -Exclusions $userExclusions -Report $report
    }

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

