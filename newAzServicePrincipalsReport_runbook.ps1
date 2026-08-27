<#

# Description: This script generates a report of new Azure Service Principals created in the last 30 days and saves the results to a storage blob.
The report includes:
- ID, DisplayName, AppId, CreatedDateTime, and LastSignInDateTime of the service principals.
- Any API permissions granted to the service principals.
- Any owners of the service principals.
- Any secrets or certificates associated with the service principals.
- Any roles assigned to the service principals.
- The report is saved in CSV format and uploaded to a specified Azure Storage Blob container.
#>

[CmdletBinding()]
param()

function Get-GraphDateTimeValue {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $propertyNames = @($PropertyName)
    if ($PropertyName.Length -gt 1) {
        $propertyNames += ($PropertyName.Substring(0, 1).ToLowerInvariant() + $PropertyName.Substring(1))
    }

    foreach ($name in $propertyNames | Select-Object -Unique) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    $additionalProperties = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($additionalProperties -and $additionalProperties.Value) {
        $additional = $additionalProperties.Value
        foreach ($name in $propertyNames | Select-Object -Unique) {
            $containsKeyMethod = $additional.PSObject.Methods['ContainsKey']
            if ($containsKeyMethod) {
                if ($additional.ContainsKey($name)) {
                    return $additional[$name]
                }
            }

            if ($additional -is [System.Collections.IDictionary]) {
                if ($additional.Contains($name)) {
                    return $additional[$name]
                }
            }

            $additionalProperty = $additional.PSObject.Properties[$name]
            if ($additionalProperty -and $null -ne $additionalProperty.Value) {
                return $additionalProperty.Value
            }
        }
    }

    return $null
}

function Get-ServicePrincipalLastSignInDateTime {
    param(
        [Parameter()]
        [string]$AppId,

        [Parameter()]
        [string]$ServicePrincipalId
    )

    if ([string]::IsNullOrWhiteSpace($AppId) -and [string]::IsNullOrWhiteSpace($ServicePrincipalId)) {
        return $null
    }

    try {
        $signIns = @()

        # Managed identities are more reliably found via servicePrincipalId in sign-in logs.
        if (-not [string]::IsNullOrWhiteSpace($ServicePrincipalId)) {
            $signIns = @(Get-MgAuditLogSignIn -All -Filter "servicePrincipalId eq '$ServicePrincipalId'" -ErrorAction Stop)
        }

        if (($null -eq $signIns -or $signIns.Count -eq 0) -and -not [string]::IsNullOrWhiteSpace($AppId)) {
            $signIns = @(Get-MgAuditLogSignIn -All -Filter "appId eq '$AppId'" -ErrorAction Stop)
        }

        if ($null -eq $signIns -or $signIns.Count -eq 0) {
            return $null
        }

        $latestSignIn = $signIns | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        return (Get-GraphDateTimeValue -InputObject $latestSignIn -PropertyName 'CreatedDateTime')
    }
    catch {
        $lookupId = if (-not [string]::IsNullOrWhiteSpace($ServicePrincipalId)) { $ServicePrincipalId } else { $AppId }
        Write-Warning "Unable to retrieve sign-in history for service principal $lookupId. $($_.Exception.Message)"
        return $null
    }
}

function Format-ListValue {
    param(
        [Parameter()]
        [object[]]$Values
    )

    if ($null -eq $Values) {
        return ""
    }

    return ($Values | Where-Object { $_ -and $_.ToString().Trim() } | ForEach-Object { $_.ToString().Trim() } | Sort-Object -Unique) -join '; '
}

$Date = Get-Date -Format 'dd-MMMM-yyyy'
$htmlHead = @'
<style>
header {
        text-align: center;
}
body {
        font-family: "Arial";
        font-size: 10pt;
        color: #4C607B;
}
table, th, td {
        width: 450px;
        border-collapse: collapse;
        border: solid;
        border: 1.5px solid black;
        padding: 3px;
}
th {
        font-size: 1.2em;
        text-align: center;
        background-color: #003366;
        color: #ffffff;
}
td {
        color: #000000;
}
tr:nth-child(even) {background-color: #d6d6d6;}
#myDisplayNameFilterID {
        width: 75%;
        font-size: 18px;
        padding: 10px 20px 10px 20px;
        border: 1.5px solid #ddd;
        margin-bottom: 15px;
}
.dropbtn {
    background-color: #003366;
    color: white;
    padding: 16px;
    font-size: 20px;
    border: none;
    cursor: pointer;
    font-weight: bold;
}
.dropbtn:hover, .dropbtn:focus {
    background-color: #b8d5e9;
}
.dropdown {
    position: relative;
    display: inline-block;
}
.dropdown-content {
    display: none;
    position: absolute;
    background-color: #f1f1f1;
    min-width: 160px;
    overflow: auto;
    box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
    z-index: 1;
}
.dropdown-content a {
    color: black;
    padding: 12px 16px;
    text-decoration: none;
    display: block;
}
.dropdown a:hover {background-color: #ddd;}
.show {display: block;}
</style>
'@

$htmlBody = @"
<font color="Black"><h1><center>New Service Principals Report - $($Date)</center></h1></font>
<div class="dropdown">
    <button onclick="myDropdownFunction()" class="dropbtn">Quick Filter</button>
    <div id="myDropdown" class="dropdown-content">
        <a href="#All Service Principals" onclick="myAccountEnabledFilter('all')"> Clear filters</a>
        <a href="#Enabled" onclick="myAccountEnabledFilter('True')">Enabled</a>
        <a href="#Disabled" onclick="myAccountEnabledFilter('False')">Disabled</a>
    </div>
</div>
<input type="text" id="myDisplayNameFilterID" onkeyup="myDisplayNameFilter()" placeholder="Search for Display Names..">
<br>
<script>
function myDropdownFunction() {
    document.getElementById("myDropdown").classList.toggle("show");
}

window.onclick = function(event) {
    if (!event.target.matches('.dropbtn')) {
        var dropdowns = document.getElementsByClassName("dropdown-content");
        var i;
        for (i = 0; i < dropdowns.length; i++) {
            var openDropdown = dropdowns[i];
            if (openDropdown.classList.contains('show')) {
                openDropdown.classList.remove('show');
            }
        }
    }
}

function myAccountEnabledFilter(a)
{
    var filter, table, tr, td, i, txtValue;
    filter = a.toUpperCase();
    table = document.getElementById("mySPTable");
    tr = table.getElementsByTagName("tr");

    if (a == "all")
    {
        for (i = 0; i < tr.length; i++)
        {
            td = tr[i].getElementsByTagName("td")[5];
            if (td)
            {
                tr[i].style.display = "";
            }
        }
    }
    else
    {
        for (i = 0; i < tr.length; i++)
        {
            td = tr[i].getElementsByTagName("td")[5];
            if (td)
            {
                txtValue = td.textContent || td.innerText;
                if (txtValue.toUpperCase().indexOf(filter) > -1)
                {
                    tr[i].style.display = "";
                } else
                {
                    tr[i].style.display = "none";
                }
            }
        }
    }
}

function myDisplayNameFilter()
{
    var input, filter, table, tr, td, i, txtValue;
    input = document.getElementById("myDisplayNameFilterID");
    filter = input.value.toUpperCase();
    table = document.getElementById("mySPTable");
    tr = table.getElementsByTagName("tr");

    for (i = 0; i < tr.length; i++)
    {
        td = tr[i].getElementsByTagName("td")[1];
        if (td)
        {
            txtValue = td.textContent || td.innerText;
            if (txtValue.toUpperCase().indexOf(filter) > -1)
            {
                tr[i].style.display = "";
            } else
            {
                tr[i].style.display = "none";
            }
        }
    }
}
</script>
"@

Connect-AzAccount -Identity -ErrorAction Stop
Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop

# Set Storage parameters
$stResourceGroupName = "rg-idmgmt-poc"
$stAccountName = "stidmgmtpocdata"
$containerName = "exports"

# Get Storage context required to upload results to blob
$stContext = (Get-AzStorageAccount -ResourceGroupName $stResourceGroupName -Name $stAccountName).Context

### Define blob name and temp path
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvBlobName = "NewServiceAccountReport_$timestamp.csv"
$htmlBlobName = "NewServiceAccountReport_$timestamp.html"
$csvTempPath = Join-Path $env:TEMP $csvBlobName
$htmlTempPath = Join-Path $env:TEMP $htmlBlobName

# Get all service principals created after the cutoff
$cutoff = (Get-Date).AddMonths(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
$allNewSp = Get-MgServicePrincipal -All -Filter "createdDateTime ge $cutoff" -Property "id,displayName,appId,createdDateTime,accountEnabled,servicePrincipalType,signInAudience,tags"
$newSpResults = @()

foreach ($sp in $allNewSp) {
    $createdDateTime = Get-GraphDateTimeValue -InputObject $sp -PropertyName 'CreatedDateTime'

    $owners = @()
    try {
        $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $sp.Id -All -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to retrieve owners for service principal $($sp.DisplayName). $($_.Exception.Message)"
    }

    $ownerSummary = @($owners | ForEach-Object {
        if ($_.DisplayName) {
            $_.DisplayName
        }
        elseif ($_.UserPrincipalName) {
            $_.UserPrincipalName
        }
        elseif ($_.Mail) {
            $_.Mail
        }
        else {
            $_.Id
        }
    })

    $appRoleAssignments = @()
    try {
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -All -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to retrieve app role assignments for service principal $($sp.DisplayName). $($_.Exception.Message)"
    }

    $delegatedPermissionGrants = @()
    try {
        $delegatedPermissionGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id -All -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to retrieve delegated permission grants for service principal $($sp.DisplayName). $($_.Exception.Message)"
    }

    $servicePrincipalDetails = $null
    try {
        $servicePrincipalDetails = Get-MgServicePrincipal -ServicePrincipalId $sp.Id -Property "passwordCredentials,keyCredentials" -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to retrieve credential details for service principal $($sp.DisplayName). $($_.Exception.Message)"
    }

    $passwordCredentials = @()
    if ($servicePrincipalDetails -and $servicePrincipalDetails.PasswordCredentials) {
        $passwordCredentials = @($servicePrincipalDetails.PasswordCredentials)
    }

    $keyCredentials = @()
    if ($servicePrincipalDetails -and $servicePrincipalDetails.KeyCredentials) {
        $keyCredentials = @($servicePrincipalDetails.KeyCredentials)
    }

    $roleAssignments = @()
    try {
        $roleAssignments = Get-AzRoleAssignment -ObjectId $sp.Id -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Unable to retrieve Azure role assignments for service principal $($sp.DisplayName). $($_.Exception.Message)"
    }

    $apiPermissionSummary = @()
    foreach ($assignment in $appRoleAssignments) {
        $resourceName = if ($assignment.ResourceDisplayName) { $assignment.ResourceDisplayName } else { $assignment.ResourceAppId }
        if ($assignment.AppRoleDisplayName) {
            $apiPermissionSummary += "$($assignment.AppRoleDisplayName) [$resourceName]"
        }
        else {
            $apiPermissionSummary += $assignment.Id
        }
    }

    foreach ($grant in $delegatedPermissionGrants) {
        $scopeSummary = if ($grant.Scope) { $grant.Scope } else { "<none>" }
        $apiPermissionSummary += "Delegated: $scopeSummary [$($grant.ResourceAppId)]"
    }

    $roleSummary = @($roleAssignments | ForEach-Object {
        if ($_.RoleDefinitionName) {
            "$($_.RoleDefinitionName) [$($_.Scope)]"
        }
        else {
            $_.Id
        }
    })

    $secretSummary = @($passwordCredentials | ForEach-Object {
        if ($_.DisplayName) {
            $_.DisplayName
        }
        else {
            "PasswordCredential"
        }
    })

    $certificateSummary = @($keyCredentials | ForEach-Object {
        if ($_.DisplayName) {
            $_.DisplayName
        }
        else {
            "KeyCredential"
        }
    })

    $newSpResults += [PSCustomObject]@{
        Id = $sp.Id
        DisplayName = $sp.DisplayName
        AppId = $sp.AppId
        CreatedDateTime = $createdDateTime
        LastSignInDateTime = Get-ServicePrincipalLastSignInDateTime -AppId $sp.AppId -ServicePrincipalId $sp.Id
        AccountEnabled = $sp.AccountEnabled
        ServicePrincipalType = $sp.ServicePrincipalType
        SignInAudience = $sp.SignInAudience
        ApiPermissions = Format-ListValue -Values $apiPermissionSummary
        Owners = Format-ListValue -Values $ownerSummary
        Secrets = Format-ListValue -Values $secretSummary
        Certificates = Format-ListValue -Values $certificateSummary
        Roles = Format-ListValue -Values $roleSummary
        Tags = Format-ListValue -Values $sp.Tags
    }
}

# Export to csv/html
if ($newSpResults.Count -gt 0) {
    $reportData = $newSpResults | Sort-Object -Property DisplayName
}
else {
    $reportData = @([PSCustomObject]@{
        Id = $null
        DisplayName = $null
        AppId = $null
        CreatedDateTime = $null
        LastSignInDateTime = $null
        AccountEnabled = $null
        ServicePrincipalType = $null
        SignInAudience = $null
        ApiPermissions = $null
        Owners = $null
        Secrets = $null
        Certificates = $null
        Roles = $null
        Tags = $null
    })
}

$reportData | Export-Csv -Path $csvTempPath -NoTypeInformation

$htmlTableData = $reportData | ConvertTo-Html -Head $htmlHead -Body $htmlBody -PostContent "<p>Creation Date: $($Date)</p>"
($htmlTableData.Replace("<table>", "<table id=`"mySPTable`">")) | Out-File $htmlTempPath

# Upload results to blob
Set-AzStorageBlobContent -File $csvTempPath -Container $containerName -Blob $csvBlobName -Context $stContext -Force
Set-AzStorageBlobContent -File $htmlTempPath -Container $containerName -Blob $htmlBlobName -Context $stContext -Force

