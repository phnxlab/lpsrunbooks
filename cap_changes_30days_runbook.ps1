<#
.SYNOPSIS
    Reports Conditional Access policies created or modified in the last 30 days.

.DESCRIPTION
    Runbook-oriented Conditional Access change report. The script uses sequential
    execution, captures Graph results explicitly, exports comma-delimited UTF-8
    CSV data, and refuses to upload an empty file.
#>
[CmdletBinding()]
param (
    [Parameter(Position = 1)]
    [string]$TenantID
)

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.SignIns'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Applications'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Users'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Groups'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Reports'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Az.Storage'; ModuleVersion = '5.0.0' }

$ErrorActionPreference = 'Stop'

function Resolve-DirectoryApp {
    param ([Parameter(Mandatory)][string]$AppId)

    switch ($AppId) {
        'All' { return 'All' }
        'Office365' { return 'Office365' }
        'MicrosoftAdminPortals' { return 'MicrosoftAdminPortals' }
        default {
            if ($servicePrincipalCache.ContainsKey($AppId)) { return $servicePrincipalCache[$AppId] }
            return "LookingUpError-$AppId"
        }
    }
}

function Resolve-NamedLocation {
    param ([Parameter(Mandatory)][string]$Id)

    switch ($Id) {
        '00000000-0000-0000-0000-000000000000' { return 'Unknown Site' }
        'All' { return 'All' }
        'AllTrusted' { return 'AllTrusted' }
        default {
            $locationName = ($namedLocations | Where-Object Id -eq $Id | Select-Object -First 1).DisplayName
            if ($locationName) { return $locationName }
            return "LookingUpError-$Id"
        }
    }
}

function Resolve-User {
    param ([Parameter(Mandatory)][string]$Id)

    switch ($Id) {
        'None' { return 'None' }
        'GuestsOrExternalUsers' { return 'GuestsOrExternalUsers' }
        'All' { return 'All' }
        default {
            if ($userCache.ContainsKey($Id)) { return $userCache[$Id] }
            return "LookingUpError-$Id"
        }
    }
}

function Resolve-DirectoryRole {
    param ([Parameter(Mandatory)][string]$Id)

    if ($directoryRoleCache.ContainsKey($Id)) { return $directoryRoleCache[$Id] }
    return "LookingUpError-$Id"
}

function Resolve-Group {
    param ([Parameter(Mandatory)][string]$Id)

    switch ($Id) {
        'None' { return 'None' }
        'GuestsOrExternalUsers' { return 'GuestsOrExternalUsers' }
        'All' { return 'All' }
        default {
            if ($groupCache.ContainsKey($Id)) { return $groupCache[$Id] }
            return "LookingUpError-$Id"
        }
    }
}

function Join-ReportValues {
    param ([object]$Value, [string]$EmptyValue = 'Not Configured')

    if ($null -eq $Value -or @($Value).Count -eq 0) { return $EmptyValue }
    return (@($Value) -join ',')
}

function Get-AuditActor {
    param ([Parameter(Mandatory)]$AuditEntry)

    if ($AuditEntry.InitiatedBy.User) {
        $user = $AuditEntry.InitiatedBy.User
        $name = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { $user.DisplayName }
        return [pscustomobject]@{ Name = $name; Type = 'User' }
    }

    if ($AuditEntry.InitiatedBy.App) {
        $app = $AuditEntry.InitiatedBy.App
        $name = if ($app.DisplayName) { $app.DisplayName } else { $app.AppId }
        return [pscustomobject]@{ Name = $name; Type = 'Application' }
    }

    return [pscustomobject]@{ Name = 'Unknown'; Type = 'Unknown' }
}

function Get-ModifiedPropertySummary {
    param ([Parameter(Mandatory)]$AuditEntries)

    $summaries = @(
        foreach ($entry in @($AuditEntries)) {
            foreach ($target in @($entry.TargetResources)) {
                foreach ($property in @($target.ModifiedProperties)) {
                    $oldValue = if ($property.OldValue) { $property.OldValue } else { '<empty>' }
                    $newValue = if ($property.NewValue) { $property.NewValue } else { '<empty>' }
                    "$($property.DisplayName): $oldValue -> $newValue"
                }
            }
        }
    )

    if ($summaries.Count -eq 0) { return 'Audit record contains no modified-property details' }
    return ($summaries -join ' | ')
}

Write-Output 'Connecting to Microsoft Graph using managed identity.'
if ([string]::IsNullOrWhiteSpace($TenantID)) {
    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
} else {
    Connect-MgGraph -Identity -TenantId $TenantID -NoWelcome -ErrorAction Stop
}

$graphContext = Get-MgContext
Write-Output "Graph tenant: $($graphContext.TenantId)"

Write-Output 'Connecting to Azure using managed identity.'
if ([string]::IsNullOrWhiteSpace($TenantID)) {
    Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
} else {
    Connect-AzAccount -Identity -Tenant $TenantID -ErrorAction Stop | Out-Null
}

$storageResourceGroupName = 'rg-idmgmt-poc'
$storageAccountName = 'stidmgmtpocdata'
$containerName = 'exports'
$storageContext = (Get-AzStorageAccount -ResourceGroupName $storageResourceGroupName -Name $storageAccountName -ErrorAction Stop).Context

$reportDate = Get-Date -Format 'dd-MMMM-yyyy'
$cutoffDate = [datetime]::UtcNow.AddDays(-30)
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$filePrefix = 'Conditional_Access_Policy_Changes_Last30Days'
$htmlBlobName = "${filePrefix}_${timestamp}.html"
$csvBlobName = "${filePrefix}_${timestamp}.csv"
$htmlTempPath = Join-Path $env:TEMP $htmlBlobName
$csvTempPath = Join-Path $env:TEMP $csvBlobName

Write-Output 'Collecting named locations.'
$namedLocations = @(
    Get-MgIdentityConditionalAccessNamedLocation -All -PageSize 999 -ErrorAction Stop |
        Select-Object DisplayName, Id
)

Write-Output 'Collecting only Conditional Access policies created or modified in the last 30 days.'
$cutoffLiteral = $cutoffDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
$modifiedPolicies = @(
    Get-MgIdentityConditionalAccessPolicy -Filter "modifiedDateTime ge $cutoffLiteral" -All -PageSize 999 -ErrorAction Stop
)
$createdPolicies = @(
    Get-MgIdentityConditionalAccessPolicy -Filter "createdDateTime ge $cutoffLiteral" -All -PageSize 999 -ErrorAction Stop
)

$policyById = @{}
foreach ($policy in @($modifiedPolicies) + @($createdPolicies)) {
    if ($policy.Id -and -not $policyById.ContainsKey($policy.Id)) {
        $policyById[$policy.Id] = $policy
    }
}
$policies = @($policyById.Values)
$modifiedPolicies = $null
$createdPolicies = $null
Write-Output "Filtered policies returned: $($policies.Count)"

if ($policies.Count -eq 0) {
    throw 'Microsoft Graph returned zero Conditional Access policies in the last 30 days.'
}

Write-Output 'Building bounded lookup caches.'
$servicePrincipalCache = @{}
Get-MgServicePrincipal -All -ErrorAction Stop | ForEach-Object {
    if ($_.AppId) { $servicePrincipalCache[$_.AppId] = $_.DisplayName }
}
$userCache = @{}
Get-MgUser -All -Property Id,UserPrincipalName -ErrorAction Stop | ForEach-Object {
    if ($_.Id) { $userCache[$_.Id] = $_.UserPrincipalName }
}
$groupCache = @{}
Get-MgGroup -All -Property Id,DisplayName -ErrorAction Stop | ForEach-Object {
    if ($_.Id) { $groupCache[$_.Id] = $_.DisplayName }
}
$directoryRoleCache = @{}
Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | ForEach-Object {
    if ($_.Id) { $directoryRoleCache[$_.Id] = $_.DisplayName }
}
Write-Output "Lookup cache sizes: users=$($userCache.Count), groups=$($groupCache.Count), service principals=$($servicePrincipalCache.Count), roles=$($directoryRoleCache.Count)"

$changedPolicies = @(
    foreach ($policy in $policies) {
        $createdDate = $null
        $modifiedDate = $null

        if ($policy.CreatedDateTime) { $createdDate = ([datetimeoffset]$policy.CreatedDateTime).UtcDateTime }
        if ($policy.ModifiedDateTime) { $modifiedDate = ([datetimeoffset]$policy.ModifiedDateTime).UtcDateTime }

        $lastChangedDate = $null
        $changeType = $null
        if ($modifiedDate -and $modifiedDate -ge $cutoffDate) {
            $lastChangedDate = $modifiedDate
            $changeType = 'Modified'
        } elseif ($createdDate -and $createdDate -ge $cutoffDate) {
            $lastChangedDate = $createdDate
            $changeType = 'Created'
        }

        if ($lastChangedDate) {
            [pscustomobject]@{
                Policy = $policy
                ChangeType = $changeType
                LastChangedDateTime = $lastChangedDate
            }
        }
    }
)

Write-Output "Policies changed in the last 30 days: $($changedPolicies.Count)"
$policyById = $null
$policies = $null

# Retain only lightweight audit entries for policies that actually changed.
$auditByPolicy = @{}
$changedPolicyIds = @{}
foreach ($change in $changedPolicies) {
    $changedPolicyIds[$change.Policy.Id] = $true
}

if ($changedPolicyIds.Count -gt 0) {
    Write-Output 'Streaming Conditional Access audit events.'
    $auditFilter = "category eq 'Policy' and activityDateTime ge $cutoffLiteral and (activityDisplayName eq 'Add Conditional Access policy' or activityDisplayName eq 'Update Conditional Access policy' or activityDisplayName eq 'Delete Conditional Access policy')"
    Get-MgAuditLogDirectoryAudit -All -PageSize 999 -Property activityDateTime,activityDisplayName,initiatedBy,targetResources -Filter $auditFilter -ErrorAction Stop | ForEach-Object {
        $auditEntry = $_
        foreach ($target in @($auditEntry.TargetResources)) {
            if (-not $target.Id -or -not $changedPolicyIds.ContainsKey($target.Id)) {
                continue
            }

            if (-not $auditByPolicy.ContainsKey($target.Id)) {
                $auditByPolicy[$target.Id] = New-Object System.Collections.ArrayList
            }

            $lightweightEntry = [pscustomobject]@{
                ActivityDateTime = $auditEntry.ActivityDateTime
                ActivityDisplayName = $auditEntry.ActivityDisplayName
                InitiatedBy = $auditEntry.InitiatedBy
                TargetResources = @([pscustomobject]@{
                    ModifiedProperties = @($target.ModifiedProperties)
                })
            }
            [void]$auditByPolicy[$target.Id].Add($lightweightEntry)
        }
    }
}

$policies = $null
Write-Output "Policies with matching audit events: $($auditByPolicy.Count)"

$report = @(
    foreach ($change in $changedPolicies) {
        $policy = $change.Policy
        $auditEntries = @($auditByPolicy[$policy.Id] | Sort-Object ActivityDateTime)
        $auditActors = @($auditEntries | ForEach-Object { Get-AuditActor $_ } | Sort-Object Name -Unique)
        $auditActivities = @($auditEntries | ForEach-Object ActivityDisplayName | Where-Object { $_ } | Sort-Object -Unique)
        $changedBy = if ($auditActors.Count -gt 0) { ($auditActors.Name -join ', ') } else { 'Audit event not found' }
        $changedByType = if ($auditActors.Count -gt 0) { ($auditActors.Type -join ', ') } else { 'Unknown' }
        $auditActivity = if ($auditActivities.Count -gt 0) { ($auditActivities -join ', ') } else { 'Audit event not found' }
        $auditDate = if ($auditEntries.Count -gt 0) { ($auditEntries | Select-Object -Last 1).ActivityDateTime } else { $null }
        $modifiedProperties = if ($auditEntries.Count -gt 0) { Get-ModifiedPropertySummary $auditEntries } else { 'Audit event not found' }
        [pscustomobject]@{
            DisplayName = $policy.DisplayName
            Description = $policy.Description
            ChangeType = $change.ChangeType
            LastChangedDateTime = $change.LastChangedDateTime
            AuditActivityDateTime = $auditDate
            ChangedBy = $changedBy
            ChangedByType = $changedByType
            AuditActivity = $auditActivity
            ModifiedProperties = $modifiedProperties
            State = $policy.State
            ID = $policy.Id
            CreatedDateTime = if ($policy.CreatedDateTime) { $policy.CreatedDateTime } else { 'Null' }
            ModifiedDateTime = if ($policy.ModifiedDateTime) { $policy.ModifiedDateTime } else { 'Null' }
            UserIncludeUsers = if ($policy.Conditions.Users.IncludeUsers) { (@($policy.Conditions.Users.IncludeUsers | ForEach-Object { Resolve-User $_ }) -join ',') } else { 'Not Configured' }
            DirectoryRolesInclude = if ($policy.Conditions.Users.IncludeRoles) { (@($policy.Conditions.Users.IncludeRoles | ForEach-Object { Resolve-DirectoryRole $_ }) -join ',') } else { 'Not Configured' }
            UserExcludeUsers = if ($policy.Conditions.Users.ExcludeUsers) { (@($policy.Conditions.Users.ExcludeUsers | ForEach-Object { Resolve-User $_ }) -join ',') } else { 'Not Configured' }
            DirectoryRolesExclude = if ($policy.Conditions.Users.ExcludeRoles) { (@($policy.Conditions.Users.ExcludeRoles | ForEach-Object { Resolve-DirectoryRole $_ }) -join ',') } else { 'Not Configured' }
            UserIncludeGroups = if ($policy.Conditions.Users.IncludeGroups) { (@($policy.Conditions.Users.IncludeGroups | ForEach-Object { Resolve-Group $_ }) -join ',') } else { 'Not Configured' }
            UserExcludeGroups = if ($policy.Conditions.Users.ExcludeGroups) { (@($policy.Conditions.Users.ExcludeGroups | ForEach-Object { Resolve-Group $_ }) -join ',') } else { 'Not Configured' }
            ConditionSignInRiskLevels = Join-ReportValues $policy.Conditions.SignInRiskLevels
            ConditionClientAppTypes = Join-ReportValues $policy.Conditions.ClientAppTypes
            PlatformIncludePlatforms = Join-ReportValues $policy.Conditions.Platforms.IncludePlatforms
            PlatformExcludePlatforms = Join-ReportValues $policy.Conditions.Platforms.ExcludePlatforms
            DevicesFilterStatesMode = if ($policy.Conditions.Devices.DeviceFilter.Mode) { $policy.Conditions.Devices.DeviceFilter.Mode } else { 'Not Configured' }
            DevicesFilterStatesRule = if ($policy.Conditions.Devices.DeviceFilter.Rule) { $policy.Conditions.Devices.DeviceFilter.Rule } else { 'Not Configured' }
            ApplicationIncludeApplications = if ($policy.Conditions.Applications.IncludeApplications) { (@($policy.Conditions.Applications.IncludeApplications | ForEach-Object { Resolve-DirectoryApp $_ }) -join ',') } else { 'Not Configured' }
            ApplicationExcludeApplications = if ($policy.Conditions.Applications.ExcludeApplications) { (@($policy.Conditions.Applications.ExcludeApplications | ForEach-Object { Resolve-DirectoryApp $_ }) -join ',') } else { 'Not Configured' }
            ApplicationIncludeUserActions = Join-ReportValues $policy.Conditions.Applications.IncludeUserActions
            LocationIncludeLocations = if ($policy.Conditions.Locations.IncludeLocations) { (@($policy.Conditions.Locations.IncludeLocations | ForEach-Object { Resolve-NamedLocation $_ }) -join ',') } else { 'Not Configured' }
            LocationExcludeLocations = if ($policy.Conditions.Locations.ExcludeLocations) { (@($policy.Conditions.Locations.ExcludeLocations | ForEach-Object { Resolve-NamedLocation $_ }) -join ',') } else { 'Not Configured' }
            GrantControlBuiltInControls = Join-ReportValues $policy.GrantControls.BuiltInControls
            GrantControlTermsOfUse = Join-ReportValues $policy.GrantControls.TermsOfUse
            GrantControlOperator = if ($policy.GrantControls.Operator) { $policy.GrantControls.Operator } else { 'Not Configured' }
            GrantControlCustomAuthenticationFactors = Join-ReportValues $policy.GrantControls.CustomAuthenticationFactors
            CloudAppSecurityCloudAppSecurityType = if ($policy.SessionControls.CloudAppSecurity.CloudAppSecurityType) { $policy.SessionControls.CloudAppSecurity.CloudAppSecurityType } else { 'Not Configured' }
            ApplicationEnforcedRestrictions = if ($null -ne $policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) { $policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled } else { 'Not Configured' }
            CloudAppSecurityIsEnabled = if ($null -ne $policy.SessionControls.CloudAppSecurity.IsEnabled) { $policy.SessionControls.CloudAppSecurity.IsEnabled } else { 'Not Configured' }
            PersistentBrowserIsEnabled = if ($null -ne $policy.SessionControls.PersistentBrowser.IsEnabled) { $policy.SessionControls.PersistentBrowser.IsEnabled } else { 'Not Configured' }
            PersistentBrowserMode = if ($policy.SessionControls.PersistentBrowser.Mode) { $policy.SessionControls.PersistentBrowser.Mode } else { 'Not Configured' }
            SignInFrequencyIsEnabled = if ($null -ne $policy.SessionControls.SignInFrequency.IsEnabled) { $policy.SessionControls.SignInFrequency.IsEnabled } else { 'Not Configured' }
            SignInFrequencyType = if ($policy.SessionControls.SignInFrequency.Type) { $policy.SessionControls.SignInFrequency.Type } else { 'Not Configured' }
            SignInFrequencyValue = if ($null -ne $policy.SessionControls.SignInFrequency.Value) { $policy.SessionControls.SignInFrequency.Value } else { 'Not Configured' }
        }
    }
)

if ($report.Count -eq 0) {
    $report = @([pscustomobject]@{
        DisplayName = 'No policies changed in the last 30 days.'
        Description = ''
        ChangeType = ''
        LastChangedDateTime = ''
        State = ''
        ID = ''
        CreatedDateTime = ''
        ModifiedDateTime = ''
        AuditActivityDateTime = ''
        ChangedBy = ''
        ChangedByType = ''
        AuditActivity = ''
        ModifiedProperties = ''
    })
}

$reportData = @($report | Sort-Object LastChangedDateTime -Descending)
Write-Output "Report rows created: $($reportData.Count)"

$htmlHead = @'
<style>
body { font-family: Arial; font-size: 10pt; color: #4C607B; }
table, th, td { border-collapse: collapse; border: 1.5px solid black; padding: 3px; }
th { font-size: 1.2em; text-align: center; background-color: #003366; color: white; }
td { color: black; }
tr:nth-child(even) { background-color: #d6d6d6; }
</style>
'@
$htmlBody = "<h1>Conditional Access Policy Changes - Last 30 Days ($reportDate)</h1><p>Cutoff: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss')) UTC</p>"

Write-Output "Generating HTML report: $htmlBlobName"
$htmlTableData = $reportData | ConvertTo-Html -Head $htmlHead -Body $htmlBody -PostContent "<p>Creation Date: $reportDate</p>"
$htmlTableData.Replace('<table>', '<table id="myCATable">') | Out-File -FilePath $htmlTempPath -Encoding UTF8 -ErrorAction Stop

Write-Output "Generating CSV report: $csvBlobName"
$reportData | Export-Csv -Path $csvTempPath -NoTypeInformation -Delimiter ',' -Encoding UTF8 -ErrorAction Stop

foreach ($filePath in @($htmlTempPath, $csvTempPath)) {
    if (-not (Test-Path -LiteralPath $filePath)) { throw "Expected report file was not created: $filePath" }
    $fileInfo = Get-Item -LiteralPath $filePath -ErrorAction Stop
    Write-Output "Generated $($fileInfo.Name): $($fileInfo.Length) bytes"
    if ($fileInfo.Length -eq 0) { throw "Generated report file is empty: $filePath" }
}

Write-Output 'Uploading reports to Azure Blob Storage.'
Set-AzStorageBlobContent -File $htmlTempPath -Container $containerName -Blob $htmlBlobName -Context $storageContext -Force -ErrorAction Stop | Out-Null
Set-AzStorageBlobContent -File $csvTempPath -Container $containerName -Blob $csvBlobName -Context $storageContext -Force -ErrorAction Stop | Out-Null

Write-Output 'Conditional Access change report completed successfully.'
