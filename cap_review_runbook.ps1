<#
.SYNOPSIS
    Generates a Conditional Access policy report and uploads the reports to Azure Blob Storage.

.DESCRIPTION
    Runbook-oriented version of New_CAP_Review.ps1. The script intentionally uses
    sequential execution rather than top-level begin/process/end blocks and refuses
    to upload empty report files.
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
#Requires -Modules @{ ModuleName = 'Az.Accounts'; ModuleVersion = '2.15.0' }
#Requires -Modules @{ ModuleName = 'Az.Storage'; ModuleVersion = '5.0.0' }

$ErrorActionPreference = 'Stop'

function Resolve-DirectoryApp {
    param (
        [Parameter(Mandatory)]
        [string]$AppId
    )

    switch ($AppId) {
        'All' { return 'All' }
        'Office365' { return 'Office365' }
        'MicrosoftAdminPortals' { return 'MicrosoftAdminPortals' }
        default {
            $appName = ($servicePrincipals | Where-Object AppId -eq $AppId | Select-Object -First 1).DisplayName
            if ($appName) {
                return $appName
            }

            return "LookingUpError-$AppId"
        }
    }
}

function Resolve-NamedLocation {
    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    switch ($Id) {
        '00000000-0000-0000-0000-000000000000' { return 'Unknown Site' }
        'All' { return 'All' }
        'AllTrusted' { return 'AllTrusted' }
        default {
            $locationName = ($namedLocations | Where-Object Id -eq $Id | Select-Object -First 1).DisplayName
            if ($locationName) {
                return $locationName
            }

            return "LookingUpError-$Id"
        }
    }
}

function Resolve-User {
    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    switch ($Id) {
        'None' { return 'None' }
        'GuestsOrExternalUsers' { return 'GuestsOrExternalUsers' }
        'All' { return 'All' }
        default {
            $user = Get-MgUser -UserId $Id -ErrorAction SilentlyContinue
            if ($user.UserPrincipalName) {
                return $user.UserPrincipalName
            }

            return "LookingUpError-$Id"
        }
    }
}

function Resolve-DirectoryRole {
    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    $roleName = ($directoryRoleTemplates | Where-Object Id -eq $Id | Select-Object -First 1).DisplayName
    if ($roleName) {
        return $roleName
    }

    return "LookingUpError-$Id"
}

function Resolve-Group {
    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    switch ($Id) {
        'None' { return 'None' }
        'GuestsOrExternalUsers' { return 'GuestsOrExternalUsers' }
        'All' { return 'All' }
        default {
            $group = Get-MgGroup -GroupId $Id -ErrorAction SilentlyContinue
            if ($group.DisplayName) {
                return $group.DisplayName
            }

            return "LookingUpError-$Id"
        }
    }
}

function Get-NamedLocationType {
    param (
        [Parameter(Mandatory)]
        [string]$Type
    )

    switch ($Type) {
        '#microsoft.graph.ipNamedLocation' { return 'ipNamedLocation' }
        '#microsoft.graph.countryNamedLocation' { return 'countryNamedLocation' }
        default { return 'UnknownType' }
    }
}

function Join-ReportValues {
    param (
        [object]$Value,
        [string]$EmptyValue = 'Not Configured'
    )

    if ($null -eq $Value -or @($Value).Count -eq 0) {
        return $EmptyValue
    }

    return (@($Value) -join ',')
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
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$policyFilePrefix = 'Conditional_Access_Policies_Report'
$namedLocationFilePrefix = 'NamedLocations_Report'
$htmlBlobName = "${policyFilePrefix}_${timestamp}.html"
$csvBlobName = "${policyFilePrefix}_${timestamp}.csv"
$namedLocationsBlobName = "${namedLocationFilePrefix}_${timestamp}.csv"
$htmlTempPath = Join-Path $env:TEMP $htmlBlobName
$csvTempPath = Join-Path $env:TEMP $csvBlobName
$namedLocationsTempPath = Join-Path $env:TEMP $namedLocationsBlobName

Write-Output 'Collecting named locations.'
$namedLocations = @(
    Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop |
        Select-Object DisplayName, Id,
            @{Name = 'Type'; Expression = { Get-NamedLocationType -Type $_.AdditionalProperties.'@odata.type' }},
            @{Name = 'IsTrusted'; Expression = { $_.AdditionalProperties.isTrusted }},
            @{Name = 'IpRanges'; Expression = { $_.AdditionalProperties.ipRanges.cidrAddress -join ',' }},
            @{Name = 'Country'; Expression = { $_.AdditionalProperties.countriesAndRegions -join ',' }},
            @{Name = 'IncludeUnknownCountriesAndRegions'; Expression = { $_.AdditionalProperties.includeUnknownCountriesAndRegions }},
            @{Name = 'CountryLookupMethod'; Expression = { $_.AdditionalProperties.countryLookupMethod }}
)
Write-Output "Named locations returned: $($namedLocations.Count)"

Write-Output 'Collecting service principals.'
$servicePrincipals = @(
    Get-MgServicePrincipal -All -ErrorAction Stop | Select-Object DisplayName, AppId
)

Write-Output 'Collecting directory role templates.'
$directoryRoleTemplates = @(
    Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Select-Object DisplayName, Id
)

Write-Output 'Collecting Conditional Access policies.'
$policies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
Write-Output "Policies returned: $($policies.Count)"

if ($policies.Count -eq 0) {
    throw 'Microsoft Graph returned zero Conditional Access policies. Refusing to create or upload an empty report.'
}

$report = @(
    foreach ($policy in $policies) {
        [pscustomobject]@{
            DisplayName = $policy.DisplayName
            Description = $policy.Description
            State = $policy.State
            ID = $policy.Id
            CreatedDateTime = if ($policy.CreatedDateTime) { $policy.CreatedDateTime } else { 'Null' }
            ModifiedDateTime = if ($policy.ModifiedDateTime) { $policy.ModifiedDateTime } else { 'Null' }
            UserIncludeUsers = if ($policy.Conditions.Users.IncludeUsers) { (@($policy.Conditions.Users.IncludeUsers | ForEach-Object { Resolve-User -Id $_ }) -join ',') } else { 'Not Configured' }
            DirectoryRolesInclude = if ($policy.Conditions.Users.IncludeRoles) { (@($policy.Conditions.Users.IncludeRoles | ForEach-Object { Resolve-DirectoryRole -Id $_ }) -join ',') } else { 'Not Configured' }
            UserExcludeUsers = if ($policy.Conditions.Users.ExcludeUsers) { (@($policy.Conditions.Users.ExcludeUsers | ForEach-Object { Resolve-User -Id $_ }) -join ',') } else { 'Not Configured' }
            DirectoryRolesExclude = if ($policy.Conditions.Users.ExcludeRoles) { (@($policy.Conditions.Users.ExcludeRoles | ForEach-Object { Resolve-DirectoryRole -Id $_ }) -join ',') } else { 'Not Configured' }
            UserIncludeGroups = if ($policy.Conditions.Users.IncludeGroups) { (@($policy.Conditions.Users.IncludeGroups | ForEach-Object { Resolve-Group -Id $_ }) -join ',') } else { 'Not Configured' }
            UserExcludeGroups = if ($policy.Conditions.Users.ExcludeGroups) { (@($policy.Conditions.Users.ExcludeGroups | ForEach-Object { Resolve-Group -Id $_ }) -join ',') } else { 'Not Configured' }
            ConditionSignInRiskLevels = Join-ReportValues $policy.Conditions.SignInRiskLevels
            ConditionClientAppTypes = Join-ReportValues $policy.Conditions.ClientAppTypes
            PlatformIncludePlatforms = Join-ReportValues $policy.Conditions.Platforms.IncludePlatforms
            PlatformExcludePlatforms = Join-ReportValues $policy.Conditions.Platforms.ExcludePlatforms
            DevicesFilterStatesMode = if ($policy.Conditions.Devices.DeviceFilter.Mode) { $policy.Conditions.Devices.DeviceFilter.Mode } else { 'Not Configured' }
            DevicesFilterStatesRule = if ($policy.Conditions.Devices.DeviceFilter.Rule) { $policy.Conditions.Devices.DeviceFilter.Rule } else { 'Not Configured' }
            ApplicationIncludeApplications = if ($policy.Conditions.Applications.IncludeApplications) { (@($policy.Conditions.Applications.IncludeApplications | ForEach-Object { Resolve-DirectoryApp -AppId $_ }) -join ',') } else { 'Not Configured' }
            ApplicationExcludeApplications = if ($policy.Conditions.Applications.ExcludeApplications) { (@($policy.Conditions.Applications.ExcludeApplications | ForEach-Object { Resolve-DirectoryApp -AppId $_ }) -join ',') } else { 'Not Configured' }
            ApplicationIncludeUserActions = Join-ReportValues $policy.Conditions.Applications.IncludeUserActions
            LocationIncludeLocations = if ($policy.Conditions.Locations.IncludeLocations) { (@($policy.Conditions.Locations.IncludeLocations | ForEach-Object { Resolve-NamedLocation -Id $_ }) -join ',') } else { 'Not Configured' }
            LocationExcludeLocations = if ($policy.Conditions.Locations.ExcludeLocations) { (@($policy.Conditions.Locations.ExcludeLocations | ForEach-Object { Resolve-NamedLocation -Id $_ }) -join ',') } else { 'Not Configured' }
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

Write-Output "Report rows created: $($report.Count)"
if ($report.Count -eq 0) {
    throw 'Report construction produced zero rows. Refusing to export or upload.'
}

$reportData = @(
    $report | Select-Object -Property DisplayName, Description, State, ID, CreatedDateTime, ModifiedDateTime,
        UserIncludeUsers, UserExcludeUsers, DirectoryRolesInclude, DirectoryRolesExclude, UserIncludeGroups, UserExcludeGroups,
        ConditionSignInRiskLevels, ConditionClientAppTypes, PlatformIncludePlatforms, PlatformExcludePlatforms,
        DevicesFilterStatesMode, DevicesFilterStatesRule, ApplicationIncludeApplications, ApplicationExcludeApplications,
        ApplicationIncludeUserActions, LocationIncludeLocations, LocationExcludeLocations, GrantControlBuiltInControls,
        GrantControlTermsOfUse, GrantControlOperator, GrantControlCustomAuthenticationFactors,
        ApplicationEnforcedRestrictions, CloudAppSecurityCloudAppSecurityType, CloudAppSecurityIsEnabled,
        PersistentBrowserIsEnabled, PersistentBrowserMode, SignInFrequencyIsEnabled, SignInFrequencyType,
        SignInFrequencyValue | Sort-Object DisplayName
)

if ($reportData.Count -eq 0) {
    throw 'ReportData is empty. Refusing to export or upload.'
}

$htmlHead = @'
<style>
body { font-family: Arial; font-size: 10pt; color: #4C607B; }
table, th, td { border-collapse: collapse; border: 1.5px solid black; padding: 3px; }
th { font-size: 1.2em; text-align: center; background-color: #003366; color: white; }
td { color: black; }
tr:nth-child(even) { background-color: #d6d6d6; }
</style>
'@

$htmlBody = @"
<h1>Conditional Access Policies Report - $reportDate</h1>
"@

Write-Output "Generating HTML report: $htmlBlobName"
$htmlTableData = $reportData | ConvertTo-Html -Head $htmlHead -Body $htmlBody -PostContent "<p>Creation Date: $reportDate</p>"
$htmlTableData.Replace('<table>', '<table id="myCATable">') | Out-File -FilePath $htmlTempPath -Encoding UTF8 -ErrorAction Stop

Write-Output "Generating CSV reports: $csvBlobName and $namedLocationsBlobName"
$reportData | Export-Csv -Path $csvTempPath -NoTypeInformation -Delimiter ',' -Encoding UTF8 -ErrorAction Stop
$namedLocations | Export-Csv -Path $namedLocationsTempPath -NoTypeInformation -Delimiter ',' -Encoding UTF8 -ErrorAction Stop

$generatedFiles = @($htmlTempPath, $csvTempPath, $namedLocationsTempPath)
foreach ($filePath in $generatedFiles) {
    if (-not (Test-Path -LiteralPath $filePath)) {
        throw "Expected report file was not created: $filePath"
    }

    $fileInfo = Get-Item -LiteralPath $filePath -ErrorAction Stop
    Write-Output "Generated $($fileInfo.Name): $($fileInfo.Length) bytes"
    if ($fileInfo.Length -eq 0) {
        throw "Generated report file is empty: $filePath"
    }
}

Write-Output 'Uploading reports to Azure Blob Storage.'
Set-AzStorageBlobContent -File $htmlTempPath -Container $containerName -Blob $htmlBlobName -Context $storageContext -Force -ErrorAction Stop | Out-Null
Set-AzStorageBlobContent -File $csvTempPath -Container $containerName -Blob $csvBlobName -Context $storageContext -Force -ErrorAction Stop | Out-Null
Set-AzStorageBlobContent -File $namedLocationsTempPath -Container $containerName -Blob $namedLocationsBlobName -Context $storageContext -Force -ErrorAction Stop | Out-Null

Write-Output 'Report generation and upload completed successfully.'
