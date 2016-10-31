# Exchange User Maintenance Script
# Version 1.4.7
# Operations:
#	*Enables Mailboxes for user in select OUs
#	*Enables User and Contact to show up in the GAL and be part of Distribution lists from select OUs. (Does Not Create Mailboxes) 
#	*Remove groups and distribution lists from $DisabledOUDN OU.
#	*Disabled Mail Users for $DisabledOUDN OU
#	*Disabled Mail Box Users for $DisabledOUDN OU
#	*Enable "No Longer With" Users for $DisabledOUWithEmailRule 
#	*Disabled Mail Box Users for $DisabledOUWithEmailRule Over $PSTExportTime Days
#Dependencies for this script:
#	*Active Directory PowerShell Tools Installed
#	*Active Directory Administrative Rights
#	*Exchange impersonation Rights
#	*Exchange Remote Management Shell
#	*Exchange Administrative Rights
#	*Exchange Trusted Subsystem needs to have Modify rights to user Home Drive
#	*EWS Managed API needs to be installed
# Code snippits from Sources:
#	http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
#	http://poshcode.org/624
#Changes:
#	*Updated Enable-MailUser and Get-MailboxExportRequest for Exchange 2013
#	*Updated to see Failed status when exporting to PST
#	*Updated exporting to PST to be cleaner
#	*Updated to allow groups to have names with spaces
#	*Added Enable Mailboxes - Version 1.4.0
#	*Updated Auto-Reply text - Version 1.4.3
#	*Updated Mail Export to stop infinite loop - Version 1.4.7
#############################################################################
# User Variables
#############################################################################

#User Home Drive Share
$HomeDriveShare = "\\File Server FQDN\Share"
$PSTFolder = "Outlook"
$PSTExportTime = 120
$ExchangeServer = "Exchange Server"
$Company = "Company Name"
$Database = "Mail DataBase"

#Organizational Units need to be in DistinguishedName format
$EnableMailboxUserOUs = "OU Name to Create Mailbox","2nd OU Name to Create Mailbox"
$EnableEmailUsersOUs = "OU Name to Mail Enable","2nd OU Name to Mail Enable"
$ExchangeGroupsOU = "Exchange E-Mail Groups OU","2nd Exchange E-Mail Groups OU"
$ADContactOU = "AD Contacts OU Name"
$DisabledOUDN = "OU for Disabled user with no Exchange Attribute"
$DisabledOUWithEmailRule = "OU for no longer with $Company User"

#############################################################################

##Load Active Directory Module
# Load AD PSSnapins
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}
## Load Exchange WebServices API dll  
If ((Get-Module | Where-Object {$_.Name -Match "Microsoft.Exchange.WebServices"}).Count -eq 0 ) {
	Write-Host ("Loading Exchange WebServices Plugins") -foregroundcolor "Green"
	###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
	## Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
	$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
	if (Test-Path $EWSDLL) {Import-Module $EWSDLL -ErrorAction SilentlyContinue}
	#Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -ErrorAction SilentlyContinue
	
	## Set Exchange Version
	$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
	## Create Exchange Web Service Object  
	$EWSservice = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
	$EWSservice.UseDefaultCredentials = $true
	## End Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
} Else {
	Write-Host ("Exchange WebServices Plug-ins Already Loaded") -foregroundcolor "Green"
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
	$enablemailusers = get-user -organizationalUnit $EMOU  | where-object{$_.RecipientType -eq "User" -and $_.WindowsEmailAddress -ne $null}
	$enablemailusers | ForEach-Object { 
		$data = $_.WindowsEmailAddress -split("@")
		if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
			Write-Host ("`tEnable Mail Name: " + $_.Name + " Alias: " + $_.SamAccountName + " Email: " + $_.WindowsEmailAddress) -foregroundcolor "Gray"
			#Remove any Exchange Attributes to reduce errors
			set-aduser -Identity $_.SamAccountName -clear msExchMailboxGuid,msexchhomeservername,legacyexchangedn,mailnickname,msexchmailboxsecuritydescriptor,msexchpoliciesincluded,msexchrecipientdisplaytype,msexchrecipienttypedetails,msexchumdtmfmap,msexchuseraccountcontrol,msexchversion	
			#Write-Host ("`tEnable-MailUser -Identity " + $_.Name + " -ExternalEmailAddress " + $_.WindowsEmailAddress + " -Alias " + $_.SamAccountName)
			Enable-MailUser -Identity $_.Name -ExternalEmailAddress $_.WindowsEmailAddress.tostring() -Alias $_.SamAccountName.tostring() 
		}
	}
}
#MailBoxes
ForEach ($EMOU in $EnableMailboxUserOUs) {
	Write-Host ("")
	Write-Host ("Searching for Users to Create Mailboxes for in OU: `t $EMOU"  )
	#Mail Enable All user that have E-Mail Address in an AD OU "UIC Campus Users"
	$enablemailusers = get-user -organizationalUnit $EMOU  | where-object{$_.RecipientType -eq "User" -and $_.WindowsEmailAddress -ne $null}
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
$enablemailusers = Get-Contact -organizationalUnit $ADContactOU| where-object { $_.RecipientType -NotLike "*Mail*" -and $_.WindowsEmailAddress -ne $null }
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
$enablemailusers = get-user -organizationalUnit $DisabledOUDN | where-object {$_.RecipientType -ne "User" -and $_.WindowsEmailAddress -ne $null}
ForEach ($EEUser in $enablemailusers) {
	$CurrentMailBox = $EEUser | Get-Mailbox
	If ($EEUser.WindowsEmailAddress -ne "") {
		If ($EEUser.RecipientType -eq "MailUser" ) {
			Write-Host ("`tDisable Mail Name: " + $EEUser.Name + " Alias: " + $EEUser.SamAccountName + " Email: " + $EEUser.WindowsEmailAddress) -foregroundcolor "magenta"
			Disable-MailUser -Identity $EEUser.SamAccountName -Confirm:$False
		}
		If ($EEUser.RecipientType -eq "UserMailbox" ) {
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
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
				If ($ExportJobStatusName -ne $null) {
					Write-Host ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
					while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
						#View Status of Mailbox Export
						Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
						 Write-Progress -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete:" + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
						 #$ExportJobStatusName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
						Start-Sleep -Seconds 10
					}
				}
				$ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And $_.status -ne "Completed" -And $_.status -ne "Failed"} 
				$ExportMailBoxListCompleted = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity} 
				If ($ExportMailBoxList.count -eq $ExportMailBoxListCompleted.Count) {
					#Remove mailbox from Exchange
					Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false
				}
			} else {
				Write-Host ("`t`tUser " + $EEUser.Name + " already submitted. " + $DisabledOUDN)
				$ExportJobStatusName = $null
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
				If ($ExportJobStatusName -ne $null -And $ExportJobStatusName.status -ne 10) {
					while  ($ExportJobStatusName.status -ne 10) {
						Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
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
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
				If ($ExportJobStatusName.status -eq 10) {
					#Remove mailbox from Exchange
					Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
					Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false
				}
			}
		}
	}
}

Write-Host ("Searching for Disable Users in OU: `t $DisabledOUWithEmailRule")

#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users under 6 months"
$enablemailusers = get-user -organizationalUnit $DisabledOUWithEmailRule | where-object {$_.RecipientType -ne "User" -and $_.WindowsEmailAddress -ne $null}
ForEach ($CurrentAccount In $enablemailusers) { 
	$CurrentMailBox = $CurrentAccount | Get-Mailbox
	If ( $($CurrentAccount.WindowsEmailAddress) -ne "" ) {
		#Need to parse out description to get date and then see if it is over 6 months.
		$ADUser = Get-adUser $CurrentAccount.SamAccountName -Properties Description,Manager
		#converts string to date
		$StrTestDate = [datetime]::ParseExact($ADUser.description.substring(0,8),"yyyyMMdd",$null)
		#Find out how old
		$currentdate= GET-DATE
		$TimeSpan = [DateTime]$currentdate - [DateTime]$StrTestDate
		$UsersManager= get-user $CurrentAccount.Manager
		#Look to see if OOA E-Mail is set
		
		$AllRules = Get-InboxRule -Mailbox $CurrentAccount.SamAccountName
		if ($AllRules | where-object{ $_.name -eq "Termination Auto Reply"})
		{
			#OOA Set
		} else {
			#Disable all other rules
			
			ForEach ($Rule in $AllRules) {
				Disable-InboxRule -Identity $Rule.RuleIdentity -Mailbox $CurrentAccount.WindowsEmailAddress -Confirm:$false -Force
			}
			#Write-Host ("`tCreating Email Rule for " + $($CurrentAccount.SamAccountName)) -foregroundcolor "Cyan"
			# $EWSservice.AutodiscoverUrl($CurrentAccount.WindowsEmailAddress,{$true})
			$EWSservice.URL = [system.URI] "https://$ExchangeServer/ews/exchange.asmx"
			$EWSservice.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $CurrentAccount.WindowsEmailAddress) 
			Write-Host ("`t Using CAS Server : " + $EWSservice.url)
			## Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
			#Create Message to reply with
			$templateEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $EWSservice
			$templateEmail.ItemClass = "IPM.Note.Rules.ReplyTemplate.Microsoft";
			$templateEmail.IsAssociated = $true;
			$templateEmail.Subject = "$($CurrentAccount.FirstName) is no longer with $Company";
			$htmlBodyString = " $($CurrentAccount.FirstName) is no longer with $Company For any business related needs please e-mail $($UsersManager.FirstName) at $($UsersManager.WindowsEmailAddress). ";
			$templateEmail.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody($htmlBodyString);
			$PidTagReplyTemplateId = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65C2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
			$templateEmail.SetExtendedProperty($PidTagReplyTemplateId, [System.Guid]::NewGuid().ToByteArray());
			$templateEmail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox);
		 
			#Create Inbox Rule
			$inboxRule = New-Object Microsoft.Exchange.WebServices.Data.Rule
			$inboxRule.DisplayName = "Termination Auto Reply";
			$inboxRule.Actions.ServerReplyWithMessage = $templateEmail.Id;
			$inboxRule.Exceptions.ContainsSubjectStrings.Add("RE:");
			$inboxRule.Exceptions.ContainsSubjectStrings.Add("FW:");			
			$createRule = New-Object Microsoft.Exchange.WebServices.Data.CreateRuleOperation[] 1
			$createRule[0] = $inboxRule
			$EWSservice.UpdateInboxRules($createRule,$true);
			## End Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
			
			if ($LastExitCode -eq 0) { 
				#Enable Mail forwarding to manager.
				If ($CurrentAccount.ForwardingAddress -eq $null ) {
						If (-Not [string]::IsNullOrEmpty($UsersManager.WindowsEmailAddress.ToString())) {
							Write-Host ("`tForwarding e-mail for $($CurrentAccount.SamAccountName) to $($UsersManager.Name)") -foregroundcolor "Cyan"
							$CurrentAccount | Set-Mailbox -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $($UsersManager.WindowsEmailAddress.ToString())
						}	
				}
			} else {
				#Enable Mail forwarding to manager.
				If ($CurrentAccount.ForwardingAddress -eq $null ) {
						If (-Not [string]::IsNullOrEmpty($UsersManager.WindowsEmailAddress.ToString())) {
							Write-Host ("`tForwarding e-mail for $($CurrentAccount.SamAccountName) to $($UsersManager.Name)") -foregroundcolor "Cyan"
							$CurrentAccount | Set-Mailbox -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $($UsersManager.WindowsEmailAddress.ToString())
							Write-Host ("`t Enabled Auto-Reply for: $($CurrentAccount.SamAccountName)") -foregroundcolor "yellow"
							$CurrentAccount | Set-MailboxAutoReplyConfiguration -AutoReplyState enabled -ExternalAudience "all" -InternalMessage "$($CurrentAccount.FirstName) is no longer with $Company For any business related needs please e-mail $($UsersManager.FirstName) at $($UsersManager.WindowsEmailAddress). " -ExternalMessage "$($CurrentAccount.FirstName) is no longer with $Company For any business related needs please e-mail $($UsersManager.FirstName) at $($UsersManager.WindowsEmailAddress). "
						}	
				}
			}
		}

		#Write-Host ("Testing Mail Name: " + $_.Name + " Alias: " + $($data[0]) + "Disable Date: " + $StrTestDate + " Date Age: " + $TimeSpan.TotalDays)
		
		If ($TimeSpan.TotalDays -ge $PSTExportTime) {
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

				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobName = $_}}

				If ($ExportJobName -ne $null) {
					while ($ExportJobName.status -ne 10 ) {
						Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobName = $_.name}}
						# Write-Host ("`t`t`t`t Job Status loop: " + $ExportJobName.status)
						If ($ExportJobName.status -eq "Completed") {break}
						If ($ExportJobName.status -eq "Failed") {break}
						#View Status of Mailbox Export
						Write-Progress -Activity $("Exporting user: " + $ExportJobName.SourceAlias ) -status $("Export Percent Complete:" + $ExportJobName.PercentComplete + " Copied " + $ExportJobName.BytesTransferred + " out of " + $ExportJobName.EstimatedTransferSize ) -percentComplete $ExportJobName.PercentComplete
						#$ExportJobName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
						Start-Sleep -Seconds 10
					}
				}
				#$ExportJobStatusName.status = 10 = Complete
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobName = $_}}
				If ($ExportJobName.status -eq 10) {
					Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
					#Remove mailbox from Exchange
					Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
					Write-Host ("`t`t`t`t Moving User " + $ADUser.name + " to " + $DisabledOUDN)
					#Move User to Disabled Outlook
					Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
				}
			} else {
				Write-Host ("`t`tUser " + $CurrentAccount.Name + " already submitted. " + $DisabledOUWithEmailRule)
				$ExportJobStatusName = $null
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
				If ($ExportJobStatusName -ne $null) {
					while  ($ExportJobStatusName.status -ne 10) {
						Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
						# Write-Host ("`t`t`t`t Job Status already submitted loop: " + $ExportJobStatusName.status)
						If ($ExportJobStatusName.status -eq "Completed") {break}
						If ($ExportJobStatusName.status -eq "Failed") {break}
						#View Status of Mailbox Export
						$ExportJobStatusName | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
						Start-Sleep -Seconds 10
					}
				}
				#$ExportJobStatusName.status = 10 = Complete
				Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity } | Get-MailboxExportRequestStatistics | ForEach-Object {If ($_.identity -ne $null) {$ExportJobStatusName = $_}}
				If ($ExportJobStatusName.status -eq 10) {
					#Remove mailbox from Exchange
					Write-Host ("`t`t`t`t Removing Mailbox from Exchange")
					Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
					
					#Move User to Disabled Outlook
					Write-Host ("`t`t`t`t Moving User " + $ADUser.name + " to " + $DisabledOUDN)
					Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
				}
			}
		}
    }
}
