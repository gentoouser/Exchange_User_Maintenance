<# Exchange User Maintenance Script
 Operations:
	*Enables Mailboxes for user in select OUs
	*Enables User and Contact to show up in the GAL and be part of Distribution lists from select OUs. (Does Not Create Mailboxes) 
	*Remove groups and distribution lists from $DisabledOUDN OU.
	*Disabled Mail Users for $DisabledOUDN OU
	*Disabled Mail Box Users for $DisabledOUDN OU
	*Enable "No Longer With" Users for $DisabledOUWithEmailRule 
	*Disabled Mail Box Users for $DisabledOUWithEmailRule Over $PSTExportTime Days
Dependencies for this script:
	*Active Directory PowerShell Tools Installed
	*Active Directory Administrative Rights
	*Exchange impersonation Rights
	*Exchange Remote Management Shell
	*Exchange Administrative Rights
	*Exchange Trusted Subsystem needs to have Modify rights to user Home Drive
 Code snippits from Sources:
	http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
	http://poshcode.org/624
Changes:
	*Updated Enable-MailUser and Get-MailboxExportRequest for Exchange 2013
	*Updated to see Failed status when exporting to PST
	*Updated exporting to PST to be cleaner
	*Updated to allow groups to have names with spaces
	*Added Enable Mailboxes - Version 1.4.0
	*Updated Auto-Reply text - Version 1.4.3
	*Updated Mail Export to stop infinite loop - Version 1.4.7
	*Updated Remove EWS managed API. Exchange 2016 can not use server side rule. - Version 1.4.9
	*Fixed Issue where Auto-Reply was trying to be set on MailUsers. - Version 1.4.9
	*Fixed Issue with AD description parsing. - Version 1.4.10
	*Added keywords "No Forwarding" in the description in AD to Not setup e-mail forwarding to manager. - Version 1.5.0
	*Added keywords "No Auto-Clean" in the description in AD to Not export mailbox to pst . - Version 1.5.0
	*Added Transcript Logging - Version 1.5.1
	*Updated so old e-mail address is put back on AD acount after exchange account is removed - Version 1.5.2	
	*Update issue where e-mail forwarding was not being set. - Version 1.5.3
	*Fixed issue with export to PST status display not working - Version 1.5.4
	*Fixed output issue - Version 1.5.5
	*Fixed issue to allow unlimited results 1.5.6
	*Missed some calls for in allow unlimited results 1.5.6 - 1.5.7
#>
$ScriptVersion = "1.5.7"
#############################################################################
# User Variables
#############################################################################

#User Home Drive Share
$HomeDriveShare = "\\File Server FQDN\Share"
$PSTFolder = "Outlook"
$PSTExportTime = 120
$ExchangeServer = "Exchange Server"
$Company = "Company Name"
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + `
		$MyInvocation.MyCommand.Name + "_" + `
		(Get-Date -format yyyyMMdd-hhmm) + ".log")
$Database = "Mail DataBase"

#Organizational Units need to be in DistinguishedName format
$EnableMailboxUserOUs = "OU Name to Create Mailbox","2nd OU Name to Create Mailbox"
$EnableEmailUsersOUs = "OU Name to Mail Enable","2nd OU Name to Mail Enable"
$ExchangeGroupsOU = "Exchange E-Mail Groups OU","2nd Exchange E-Mail Groups OU"
$ADContactOU = "AD Contacts OU Name"
$DisabledOUDN = "OU for Disabled user with no Exchange Attribute"
$DisabledOUWithEmailRule = "OU for no longer with $Company User"

#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Start-Transcript -Path $LogFile -Append
	Write-Host ("Script: " + $MyInvocation.MyCommand.Name)
	Write-Host ("Version: " + $ScriptVersion)
	Write-Host (" ")
}
##Load Active Directory Module
# Load AD PSSnapins
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}

# Load All Exchange PSSnapins 
If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" }).Count -eq 0 ) {
	Write-Host ("Loading Exchange Plugins") -foregroundcolor "Green"
	If ($([System.Net.Dns]::GetHostByName(($env:computerName))).hostname -eq $([System.Net.Dns]::GetHostByName(($ExchangeServer))).hostname) {
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
		. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer -auto -AllowClobber
	} else {
		$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
		Import-PSSession $ERPSession -AllowClobber
	}
} Else {
	Write-Host ("Exchange Plug-ins Already Loaded") -foregroundcolor "Green"
}

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  

## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
	public class TrustAll : System.Net.ICertificatePolicy {
	  public TrustAll() { 
	  }
	  public bool CheckValidationResult(System.Net.ServicePoint sp,
		System.Security.Cryptography.X509Certificates.X509Certificate cert, 
		System.Net.WebRequest req, int problem) {
		return true;
	  }
	}
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624


#Get Defaults Domain
$PrimaryEmailDomain = ((get-emailaddresspolicy | Where-Object { $_.Priority -Match "Lowest" } ).EnabledPrimarySMTPAddressTemplate).split('@')[-1]
#############################################################################
#Starting Main 

#MailUsers
ForEach ($EMOU in $EnableEmailUsersOUs) {
	Write-Host ("")
	Write-Host ("Searching for Users to Mail Enable in OU: `t $EMOU"  )
	#Mail Enable All user that have E-Mail Address in an AD OU "UIC Campus Users"
	$enablemailusers = get-user -organizationalUnit $EMOU -ResultSize Unlimited | where-object{$_.RecipientType -eq "User" -and $_.WindowsEmailAddress -ne $null}
	$enablemailusers | ForEach-Object { 
		$data = $_.WindowsEmailAddress -split("@")
		if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
			Write-Host ("`tEnable Mail Name: " + $_.Name + " Alias: " + $_.SamAccountName + " Email: " + $_.WindowsEmailAddress) -foregroundcolor "Gray"
			#Remove any Exchange Attributes to reduce errors
			set-aduser -Identity $_.SamAccountName -clear msExchMailboxGuid,msexchhomeservername,legacyexchangedn,mailnickname,msexchmailboxsecuritydescriptor,msexchpoliciesincluded,msexchrecipientdisplaytype,msexchrecipienttypedetails,msexchumdtmfmap,msexchuseraccountcontrol,msexchversion | out-null
			#Write-Host ("`tEnable-MailUser -Identity " + $_.Name + " -ExternalEmailAddress " + $_.WindowsEmailAddress + " -Alias " + $_.SamAccountName)
			Enable-MailUser -Identity $_.Name -ExternalEmailAddress $_.WindowsEmailAddress.tostring() -Alias $_.SamAccountName.tostring() | out-null
		}
	}
}
#MailBoxes
ForEach ($EMOU in $EnableMailboxUserOUs) {
	Write-Host ("")
	Write-Host ("Searching for Users to Create Mailboxes for in OU: `t $EMOU"  )
	#Mail Enable All user that have E-Mail Address in an AD OU "UIC Campus Users"
	$enablemailusers = get-user -organizationalUnit $EMOU -ResultSize Unlimited | where-object{$_.RecipientType -eq "User" -and $_.WindowsEmailAddress -ne $null}
	$enablemailusers | ForEach-Object { 
		$data = $_.WindowsEmailAddress -split("@")
		if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
			Write-Host ("`tEnable Mail Name: " + $_.Name + " Alias: " + $_.SamAccountName + " Email: " + $_.WindowsEmailAddress + " Database: " + $Database) -foregroundcolor "Gray"
			#Remove any Exchange Attributes to reduce errors
			set-aduser -Identity $_.SamAccountName -clear msExchMailboxGuid,msexchhomeservername,legacyexchangedn,mailnickname,msexchmailboxsecuritydescriptor,msexchpoliciesincluded,msexchrecipientdisplaytype,msexchrecipienttypedetails,msexchumdtmfmap,msexchuseraccountcontrol,msexchversion	
			#Write-Host ("`tEnable-Mailbox  -Identity " + $_.Name + " -ExternalEmailAddress " + $_.WindowsEmailAddress + " -Alias " + $_.SamAccountName)
			Enable-Mailbox  -Identity $_.Name -PrimarySmtpAddress $_.WindowsEmailAddress.tostring() -Alias $_.SamAccountName.tostring() -Database "$Database"
		}
	}
}

Write-Host ("Searching for Contacts to Mail Enable on OU: `t $ADContactOU")
#Mail Enable All contact that have E-Mail Address in an AD OU "Contacts"
$enablemailusers = Get-Contact -organizationalUnit $ADContactOU -ResultSize Unlimited| where-object { $_.RecipientType -NotLike "*Mail*" -and $_.WindowsEmailAddress -ne $null }
$enablemailusers | ForEach-Object { 
	$data = $_.WindowsEmailAddress -split("@")
	if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
		
		Write-Host ("`tEnable Contact Name: " + $_.Name + " Alias: " + $($data[0]) + " Email: " + $_.WindowsEmailAddress) -foregroundcolor "Gray"

		Enable-MailContact -Identity $_.Name -ExternalEmailAddress $($data[0] + "@" + $data[1]) -Alias $($data[0]) 
	}
}

Write-Host ("Searching for Users to Mail Disable in OU: `t $DisabledOUDN")
#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users"
get-aduser  -SearchBase $DisabledOUDN  -Filter * | ForEach-Object { 
	$UserDN = $_.DistinguishedName
	$userSAM = $_.SamAccountName
	Get-ADGroup -LDAPFilter "(member=$UserDN)" | foreach-object {
		if ($_.name -ne "Domain Users") {
			Write-Host ("`t Removing $userSAM from group $_.name") -foregroundcolor "magenta"
			if ($_.DistinguishedName.tostring().contains($ExchangeGroupsOU)) {
				Remove-DistributionGroupMember -identity $_.DistinguishedName -member $UserDN -Confirm:$False
			} else {
				remove-adgroupmember -identity $_.DistinguishedName -member $UserDN -Confirm:$False
			}
		} 
	}
}


Write-Host ("Searching for Users to Disable in Exchange in OU: `t $DisabledOUDN")
#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users"
$enablemailusers = get-user -organizationalUnit $DisabledOUDN -ResultSize Unlimited| where-object {$_.RecipientType -ne "User" -and $_.WindowsEmailAddress -ne $null}
ForEach ($EEUser in $enablemailusers) {
	If ($EEUser.WindowsEmailAddress -ne "") {
		If ($EEUser.RecipientType -eq "MailUser" ) {
			Write-Host ("`tDisable Mail Name: " + $EEUser.Name + " Alias: " + $EEUser.SamAccountName + " Email: " + $EEUser.WindowsEmailAddress) -foregroundcolor "magenta"
			$tempemail = $EEUser.WindowsEmailAddress
			Disable-MailUser -Identity $EEUser.SamAccountName -Confirm:$False
			Set-ADUser -identity $EEUser.SamAccountName -EmailAddress $tempemail
		}
		If ($EEUser.RecipientType -eq "UserMailbox" ) {
			$CurrentMailBox = $EEUser | Get-Mailbox
			#Testing to see if is in queue
			If ((Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity}).count -eq 0) {
				Write-Host ("`tExport Mail Name: " + $EEUser.Name + " Alias: " + $EEUser.SamAccountName + " Email: " + $EEUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
				#Create New Home Drive
				if (-Not (Test-Path $($HomeDriveShare + "\" + $EEUser.SamAccountName))) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $EEUser.SamAccountName)}
				if (-Not (Test-Path $($HomeDriveShare + "\" + $EEUser.SamAccountName + "\" + $PSTFolder + "\"))) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $EEUser.SamAccountName + "\" + $PSTFolder + "\")}
				#Export Mailbox to PST
				New-MailboxExportRequest -Mailbox $EEUser.SamAccountName -FilePath $($HomeDriveShare + "\" + $EEUser.SamAccountName  + "\" + $PSTFolder + "\" + $($EEUser.SamAccountName) + ".pst")
				$ExportJobName = $null

				
				$ExportJobStatusName = $null
				$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
				If ($ExportJobStatusName -ne $null) {
					Write-Host ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
					while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
						#View Status of Mailbox Export
						$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
						 Write-Progress -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete:" + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
						 #$ExportJobStatusName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
						Start-Sleep -Seconds 10
					}
				}
				$ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And $_.status -ne "Completed" -And $_.status -ne "Failed"} 
				$ExportMailBoxListCompleted = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity} 
				If ($ExportMailBoxList.count -eq $ExportMailBoxListCompleted.Count) {
					#Remove mailbox from Exchange
					$tempemail = $EEUser.WindowsEmailAddress	
					Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false
					get-user $EEUser.SamAccountName -ResultSize Unlimited | Set-ADUser -EmailAddress $tempemail
				}
			} else {
				Write-Host ("`t`tUser " + $EEUser.Name + " already submitted. " + $DisabledOUDN)
				$ExportJobStatusName = $null
				$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
				If ($ExportJobStatusName -ne $null -And $ExportJobStatusName.status -ne 10) {
					while  ($ExportJobStatusName.status -ne 10) {
						$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
						# Write-Host ("`t`t`t`t Job Status already submitted loop: " + $ExportJobStatusName.status)
						If ($ExportJobStatusName.status -eq "Completed") {break}
						If ($ExportJobStatusName.status -eq "Failed") {break}
						#View Status of Mailbox Export
						Write-Progress -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete:" + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
						#$ExportJobStatusName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
						Start-Sleep -Seconds 10
					}
				}
				#$ExportJobStatusName.status = 10 = Complete
				$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
				If ($ExportJobStatusName.status -eq 10) {
					#Remove mailbox from Exchange
					Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
					$tempemail = $EEUser.WindowsEmailAddress
					Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false
					Set-ADUser -identity $EEUser.SamAccountName -EmailAddress $tempemail
				}
			}
		}
	}
}

Write-Host ("Searching for Disable Users in OU: `t $DisabledOUWithEmailRule")

#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users under 6 months"
$enablemailusers = get-user -organizationalUnit $DisabledOUWithEmailRule -ResultSize Unlimited | where-object {$_.RecipientType -ne "User" -and $_.WindowsEmailAddress -ne $null}
ForEach ($CurrentAccount In $enablemailusers) { 
	If ( $CurrentAccount.WindowsEmailAddress -ne "") {
	$UserDN = $CurrentAccount.DistinguishedName
	$userSAM = $CurrentAccount.SamAccountName
		Get-ADGroup -LDAPFilter "(member=$UserDN)" | foreach-object {
		if ($_.name -ne "Domain Users") {
			Write-Host ("`t Removing $userSAM from group $($_.name)") -foregroundcolor "magenta"
			if ($_.DistinguishedName.tostring().contains($ExchangeGroupsOU)) {
				Remove-DistributionGroupMember -identity $_.DistinguishedName -member $UserDN -Confirm:$False
			} else {
				remove-adgroupmember -identity $_.DistinguishedName -member $UserDN -Confirm:$False
			}
		} 
	}
		If ($CurrentAccount.RecipientType -eq "MailUser" ) {
				Write-Host ("`tDisable Mail Name: " + $CurrentAccount.Name + " Alias: " + $CurrentAccount.SamAccountName + " Email: " + $CurrentAccount.WindowsEmailAddress) -foregroundcolor "magenta"
				$tempemail = $CurrentAccount.WindowsEmailAddress
				Disable-MailUser -Identity $CurrentAccount.SamAccountName -Confirm:$False
				get-user $CurrentAccount.SamAccountName  -ResultSize Unlimited | Set-ADUser -EmailAddress $tempemail				
			}	
		If ( $CurrentAccount.RecipientType -eq "UserMailbox" ) {
			$CurrentMailBox = $CurrentAccount | Get-Mailbox
			#Need to parse out description to get date and then see if it is over 6 months.
			$ADUser = Get-adUser $CurrentAccount.SamAccountName -Properties Description,Manager
			#converts string to date. Also do not display errors
			If ($ADUser.description.substring(0,8) -match "[0-9]" ) {
				$StrTestDate = [datetime]::ParseExact($ADUser.description.substring(0,8),"yyyyMMdd",$null) 
			} else {
				Write-Host ("`tPlease update AD description for " + $CurrentAccount.Name + " with deactivation date. Current description: " +  $ADUser.description) -foregroundcolor "red"
				continue 
			}
			#Find out how old
			$currentdate= GET-DATE
			$TimeSpan = [DateTime]$currentdate - [DateTime]$StrTestDate
			$UsersManager= get-user $CurrentAccount.Manager -ResultSize Unlimited
			#Look to see if OOA E-Mail is set
			
			#Enable Mail forwarding to manager.
			If ([string]::IsNullOrWhiteSpace($CurrentAccount.ForwardingAddress) -and !($ADUser.description.Contains("No Forwarding"))) {
					If (-Not [string]::IsNullOrEmpty($CurrentAccount.Manager)) {
						Write-Host ("`tForwarding e-mail for $($CurrentAccount.SamAccountName) to $($UsersManager.WindowsEmailAddress.ToString())") -foregroundcolor "Cyan"
						Set-Mailbox -Identity $CurrentAccount.SamAccountName -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $null -verbose
						Set-Mailbox -Identity $CurrentAccount.SamAccountName -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $UsersManager.WindowsEmailAddress.ToString().ToLower() -verbose
						Write-Host ("`t Enabled Auto-Reply for: $($CurrentAccount.SamAccountName)") -foregroundcolor "yellow"
						Set-MailboxAutoReplyConfiguration -Identity $CurrentAccount.SamAccountName -AutoReplyState enabled -ExternalAudience "all" -InternalMessage "$($CurrentAccount.FirstName) is no longer with $Company For any business related needs please e-mail $($UsersManager.FirstName) at $($UsersManager.WindowsEmailAddress). " -ExternalMessage "$($CurrentAccount.FirstName) is no longer with $Company For any business related needs please e-mail $($UsersManager.FirstName) at $($UsersManager.WindowsEmailAddress). "
					}	
			}

			If ($TimeSpan.TotalDays -ge $PSTExportTime -and !($ADUser.description.Contains("No Auto-Clean"))) {
				#Testing to see if is in queue
				If ((Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity }).count -eq 0) {
					Write-Host ("`tExport Mail Name: " + $CurrentAccount.Name + " Alias: " + $CurrentAccount.SamAccountName + " Email: " + $CurrentAccount.WindowsEmailAddress)  -foregroundcolor "Cyan"
					#Create New Home Drive
					if (-Not (Test-Path $($HomeDriveShare + "\" + $CurrentAccount.SamAccountName))) 
						{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $CurrentAccount.SamAccountName)}
					if (-Not (Test-Path $($HomeDriveShare + "\" + $CurrentAccount.SamAccountName + "\" + $PSTFolder + "\"))) 
						{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $CurrentAccount.SamAccountName + "\" + $PSTFolder + "\")}
					#Export Mailbox to PST
					New-MailboxExportRequest -Mailbox $CurrentAccount.SamAccountName -FilePath $($HomeDriveShare + "\" + $CurrentAccount.SamAccountName  + "\" + $PSTFolder + "\" + $($CurrentAccount.SamAccountName) + '.pst')
					$ExportJobName = $null

					$ExportJobName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 

					If ($ExportJobName -ne $null) {
						while ($ExportJobName.status -ne 10 ) {
							$ExportJobName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
							# Write-Host ("`t`t`t`t Job Status loop: " + $ExportJobName.status)
							If ($ExportJobName.status -eq "Completed") {break}
							If ($ExportJobName.status -eq "Failed") {break}
							#View Status of Mailbox Export
							Write-Progress -Activity $("Exporting user: " + $ExportJobName.SourceAlias ) -status $("Export Percent Complete:" + $ExportJobName.PercentComplete + " Copied " + $ExportJobName.BytesTransferred + " out of " + $ExportJobName.EstimatedTransferSize ) -percentComplete $ExportJobName.PercentComplete

							Start-Sleep -Seconds 10
						}
					}
					#$ExportJobStatusName.status = 10 = Complete
					$ExportJobName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
					If ($ExportJobName.status -eq 10) {
						Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
						#Remove mailbox from Exchange
						$tempemail = $CurrentAccount.WindowsEmailAddress
						Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
						Set-ADUser -Identity $CurrentAccount.SamAccountName -EmailAddress $tempemail
						Write-Host ("`t`t`t`t Moving User " + $ADUser.name + " to " + $DisabledOUDN)
						#Move User to Disabled Outlook
						Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
					}
				} else {
					Write-Host ("`t`tUser " + $CurrentAccount.Name + " already submitted. " + $DisabledOUWithEmailRule)
					$ExportJobStatusName = $null
					$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
					
					If ($ExportJobStatusName -ne $null) {
						while  ($ExportJobStatusName.status -ne 10) {
							$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
							# Write-Host ("`t`t`t`t Job Status already submitted loop: " + $ExportJobStatusName.status)
							If ($ExportJobStatusName.status -eq "Completed") {break}
							If ($ExportJobStatusName.status -eq "Failed") {break}
							#View Status of Mailbox Export
							$ExportJobStatusName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
							Start-Sleep -Seconds 10
						}
					}
					#$ExportJobStatusName.status = 10 = Complete
					$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
					If ($ExportJobStatusName.status -eq 10) {
						#Remove mailbox from Exchange
						Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
						$tempemail = $CurrentAccount.WindowsEmailAddress
						Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
						get-user $CurrentAccount.SamAccountName -ResultSize Unlimited | Set-ADUser -EmailAddress $tempemail							
						
						#Move User to Disabled Outlook
						Write-Host ("`t`t`t`t Moving User " + $ADUser.name + " to " + $DisabledOUDN)
						Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
					}
				}
			}
		}
	}
}

If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
