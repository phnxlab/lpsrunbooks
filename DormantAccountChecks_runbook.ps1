

Connect-azaccount -Identity
Connect-mggraph -Identity -NoWelcome

$dateOutputFormat = "yyyy-MM-ddTHH:mm:ssZ"

function Format-DateTimeOutput {
	param(
		[Parameter(ValueFromPipeline = $true)]
		$Value
	)

	if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) {
		return $null
	}

	try {
		return ([datetimeoffset]$Value).ToUniversalTime().ToString($dateOutputFormat)
	}
	catch {
		return ([datetime]$Value).ToUniversalTime().ToString($dateOutputFormat)
	}
}

# Set Storage parameters
$stResourceGroupName = "rg-idmgmt-poc"
$stAccountName = "stidmgmtpocdata"
$containerName = "exports"

# Get Storage context required to upload results to blob
$stcontext = (Get-AzStorageAccount -ResourceGroupName $stResourceGroupName -Name $stAccountName).context

### Dormant Guest check
$guestblobName = "DormantGuests_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$tempPath = Join-Path $env:TEMP $guestblobName
$createdcutoffdate = (Get-Date).AddMonths(-6)
$cutoffdate = (Get-Date).AddDays(-42)
$activeguests = Get-mguser -all -Property Id, DisplayName, UserPrincipalName, SigninActivity, AccountEnabled, UserType, CreatedDateTime | where-object {($_.AccountEnabled -eq 'True') -and ($_.UserType -eq 'Guest') -and ($_.CreatedDateTime -lt $createdcutoffdate)} 
$dormantguests = $activeguests | Where-object {($_.SignInActivity.LastNonInteractiveSignInDateTime -lt $cutoffdate ) -and ($_.SignInActivity.LastSignInDateTime -lt $cutoffdate)} 

# Format results output
$results = $dormantguests | Select-object Id, DisplayName, UserPrincipalName, AccountEnabled, @{Name="CreatedDateTime";Expression={Format-DateTimeOutput $_.CreatedDateTime}}, @{Name="LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSignInDateTime}}, @{Name="NonInt-LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastNonInteractiveSignInDateTime}}, @{Name="LastSuccessfullSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSuccessfulSignInDateTime}} 
# Export results to temp storage
$results | Export-csv -Path $tempPath -NoTypeInformation

# Upload results to new blob
Set-AzStorageBlobContent -File $tempPath -Container $containerName -Blob $guestblobName -Context $stcontext -Force

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

### Dormant user check
$doruserblobName = "DormantUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$dorusertempPath = Join-Path $env:TEMP $doruserblobName
$dorlicuserblobName = "DormantLicencedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$dorlicusertempPath = Join-Path $env:TEMP $dorlicuserblobName
$createdcutoffdate = (Get-Date).AddMonths(-6)
$cutoffperiod = (Get-Date).AddDays(-42)
$activeusers = Get-mguser -all -Property Id, DisplayName, UserPrincipalName, SigninActivity, AccountEnabled, UserType, createdDateTime, AssignedLicenses | where-object {($_.AccountEnabled -eq 'True') -and ($_.UserType -eq 'Member') -and ($_.DisplayName -ne 'svc*') -and ($_.CreatedDateTime -lt $createdcutoffdate)}

$dormantusers = $activeusers | Where-object {($_.SignInActivity.LastNonInteractiveSignInDateTime -lt $cutoffperiod ) -and ($_.SignInActivity.LastSignInDateTime -lt $cutoffperiod)} 
$alldormantusers = $dormantusers | Select-object Id, DisplayName, UserPrincipalName, AccountEnabled, @{Name="CreatedDateTime";Expression={Format-DateTimeOutput $_.CreatedDateTime}}, @{Name="LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSignInDateTime}}, @{Name="NonInt-LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastNonInteractiveSignInDateTime}}, @{Name="LastSuccessfullSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSuccessfulSignInDateTime}}
$licDormantUsers = $dormantusers | where-object {$_.AssignedLicenses -ne $null}
$alllicdormantusers = $licDormantUsers | select-object ID, DisplayName, UserPrincipalName, AccountEnabled, @{Name="AssignedLicenses";Expression={$_.AssignedLicenses}}, @{Name="LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSignInDateTime}}, @{Name="NonInt-LastSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastNonInteractiveSignInDateTime}}, @{Name="LastSuccessfullSignInDateTime";Expression={Format-DateTimeOutput $_.SignInActivity.LastSuccessfulSignInDateTime}}

$alldormantusers | Export-csv -Path $dorusertempPath -NoTypeInformation
$alllicdormantusers | Export-csv -Path $dorlicusertempPath -NoTypeInformation

# Upload results to new blob
Set-AzStorageBlobContent -File $dorusertempPath -Container $containerName -Blob $doruserblobName -Context $stcontext -Force
Set-AzStorageBlobContent -File $dorlicusertempPath -Container $containerName -Blob $dorlicuserblobName -Context $stcontext -Force

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



