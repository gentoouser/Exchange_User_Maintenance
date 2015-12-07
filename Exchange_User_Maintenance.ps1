# Exchange User Maintenance Script
# Version 1.3.0
# Operations:
#	*Enables User and Contact to show up in the GAL and be part of Distribution lists. (Does Not Create Mailboxes)
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

##Load Active Directory Module
Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue

## Load Exchange WebServices API dll  
## Set Exchange Version  
Write-Host ("Loading Exchange WebServices Plugins") -foregroundcolor "Green"

###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
## Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL) {Import-Module $EWSDLL -ErrorAction SilentlyContinue}
#Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -ErrorAction SilentlyContinue
## Create Exchange Web Service Object  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$EWSservice = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
$EWSservice.UseDefaultCredentials = $true
## End Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
# Load All Exchange PSSnapins 
Write-Host ("Loading Exchange Plugins") -foregroundcolor "Green"
If ($([System.Net.Dns]::GetHostByName(($env:computerName))).hostname -eq $([System.Net.Dns]::GetHostByName(($ExchangeServer))).hostname) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
} else {
	$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
	Import-PSSession $ERPSession -AllowClobber
}
#############################################################################
# User Varibles
#############################################################################

#User Home Drive Share
$HomeDriveShare = "\\File Server FQDN\Share"
$PSTFolder = "Outlook"
$PSTExportTime = 120
$ExchangeServer = "Exchange Server"
$Company = "Company Name"
$DisabledOUDN = "Disabled user Distinguished Name"
$DisabledOU = (Get-ADOrganizationalUnit $DisabledOUDN).Name
$DisabledOUWithEmailRule = "Disabled Users under 6 months"
$EnableEmailUsersOUs = "OU Name to Mail Enable","2nd OU Name to Mail Enable"
$ExchangeGroupsOU = "Exchange E-Mail Groups"
$ADContactOU = "AD Contacts OU Name"

#Set Defaults
$PrimaryEmailDomain = ((get-emailaddresspolicy | Where-Object { $_.Priority -Match "1" } ).EnabledPrimarySMTPAddressTemplate).split('@')[-1]


#############################################################################
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


ForEach ($EMOU in $EnableEmailUsersOUs) {
	Write-Host ("")
	Write-Host ("Searching for Users to Mail Enable in OU: $EMOU"  )
	#Mail Enable All user that have E-Mail Address in an AD OU "UIC Campus Users"
	$enablemailusers = get-user -organizationalUnit $EMOU  | where-object{$_.RecipientType -eq "User" -and $_.WindowsEmailAddress -ne $null}
	$enablemailusers | ForEach-Object { 
		$data = $_.WindowsEmailAddress -split("@")
		if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
			Write-Host ("`tEnable Mail Name: " + $_.Name + " Alias: " + $_.SamAccountName + " Email: " + $_.WindowsEmailAddress) -foregroundcolor "Gray"
			#Remove any Exchange Attributes to reduce errors
			set-aduser -Identity $_.SamAccountName -clear msExchMailboxGuid,msexchhomeservername,legacyexchangedn,mailnickname,msexchmailboxsecuritydescriptor,msexchpoliciesincluded,msexchrecipientdisplaytype,msexchrecipienttypedetails,msexchumdtmfmap,msexchuseraccountcontrol,msexchversion	
			Enable-MailUser -Identity $_.Name -ExternalEmailAddress $_.WindowsEmailAddress -Alias $_.SamAccountName 
		}
	}
}

Write-Host ("Searching for Contacts to Mail Enable on OU: $ADContactOU")
#Mail Enable All contact that have E-Mail Address in an AD OU "Contacts"
$enablemailusers = Get-Contact -organizationalUnit $ADContactOU| where-object { $_.RecipientType -NotLike "*Mail*" -and $_.WindowsEmailAddress -ne $null }
$enablemailusers | ForEach-Object { 
	$data = $_.WindowsEmailAddress -split("@")
	if (($data[0] -ne "") -and ($data[1] -ne $PrimaryEmailDomain)) {
		
		Write-Host ("`tEnable Contact Name: " + $_.Name + " Alias: " + $($data[0]) + " Email: " + $_.WindowsEmailAddress) -foregroundcolor "Gray"

		Enable-MailContact -Identity $_.Name -ExternalEmailAddress $($data[0] + "@" + $data[1]) -Alias $($data[0]) 
	}
}

Write-Host ("Searching for Users to Mail Disable in DN: $DisabledOUDN")
#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users"
get-aduser  -SearchBase $DisabledOUDN  -Filter * | ForEach-Object { 
	$UserDN = $_.DistinguishedName
	$userSAM = $_.SamAccountName
	Get-ADGroup -LDAPFilter "(member=$UserDN)" | foreach-object {
		if ($_.name -ne "Domain Users") {
			Write-Host ("`t Removing $userSAM from group $_.name") -foregroundcolor "magenta"
			if ($_.DistinguishedName.tostring().contains("OU=" + $ExchangeGroupsOU)) {
				Remove-DistributionGroupMember -identity $_.name -Member $UserDN -Confirm:$False
			} else {
				remove-adgroupmember -identity $_.name -member $UserDN -Confirm:$False
			}
		} 
	}
}


Write-Host ("Searching for Users to Disable in Exchange in OU: $DisabledOU")
#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users"
$enablemailusers = get-user -organizationalUnit $DisabledOU | where-object {$_.RecipientType -ne "User" -and $_.WindowsEmailAddress -ne $null}
ForEach ($EEUser in $enablemailusers) {

	if ($EEUser.WindowsEmailAddress -ne "") {
		If ($EEUser.RecipientType -eq "MailUser" ) {
			Write-Host ("`tDisable Mail Name: " + $EEUser.Name + " Alias: " + $EEUser.SamAccountName + " Email: " + $EEUser.WindowsEmailAddress) -foregroundcolor "magenta"
			Disable-MailUser -Identity $EEUser.SamAccountName -Confirm:$False
		}
		If ($EEUser.RecipientType -eq "UserMailbox" ) {
			#Testing to see if is in queue
			If ((Get-MailboxExportRequest | Where-Object { $_.Identity  -contains $EEUser.Identity -And $_.Status -ne "Completed"}) -eq $null) {
				Write-Host ("`tExport Mail Name: " + $EEUser.Name + " Alias: " + $EEUser.SamAccountName + " Email: " + $EEUser.WindowsEmailAddress) -foregroundcolor "Blue"
				#Create New Home Drive
				if (-Not (Test-Path $($HomeDriveShare + "\" + $EEUser.SamAccountName))) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $EEUser.SamAccountName)}
				if (-Not (Test-Path $($HomeDriveShare + "\" + $EEUser.SamAccountName + "\" + $PSTFolder + "\"))) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $EEUser.SamAccountName + "\" + $PSTFolder + "\")}
				#Export Mailbox to PST
				New-MailboxExportRequest -Mailbox $EEUser.SamAccountName -FilePath $($HomeDriveShare + "\" + $EEUser.SamAccountName  + "\" + $PSTFolder + "\" + $EEUser.SamAccountName + ".pst")

				while ( (Get-MailboxExportRequestStatistics -Identity $($EEUser.SamAccountName + "\MailboxExport")).status -ne "Completed" ) {
					#View Status of Mailbox Export
					Get-MailboxExportRequestStatistics -Identity $($EEUser.SamAccountName + "\MailboxExport") | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
					Start-Sleep -Seconds 10
				}

				#Remove mailbox from Exchange
				Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false			

			} else {
				Write-Host ("`t`tUser " + $EEUser.Name + " already submitted.")
				while ((Get-MailboxExportRequestStatistics -Identity ($EEUser.SamAccountName + "\MailboxExport")).status -ne $("Completed")) {
					#View Status of Mailbox Export
					Get-MailboxExportRequestStatistics -Identity ($EEUser.SamAccountName + "\MailboxExport") | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
					Start-Sleep -Seconds 10
				}

				#Remove mailbox from Exchange
				Disable-Mailbox -Identity $EEUser.SamAccountName -confirm:$false
								
			}
		}
	}
}

Write-Host ("Searching for Disable Users in OU: $DisabledOUWithEmailRule")

#Mail Disable All user that have E-Mail Address in an AD OU "Disabled Users"
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
			## Code From http://gsexdev.blogspot.in/2012/11/creating-sender-domain-auto-reply-rule.html
			ForEach ($Rule in $AllRules) {
				Disable-InboxRule -Identity $Rule.RuleIdentity -Mailbox $CurrentAccount.WindowsEmailAddress
			}
			Write-Host ("`tCreating Email Rule for $CurrentAccount.SamAccountName") -foregroundcolor "Blue"
			$EWSservice.AutodiscoverUrl($CurrentAccount.WindowsEmailAddress,{$true})
			$EWSservice.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $CurrentAccount.WindowsEmailAddress) 
			Write-Host ("`t Using CAS Server : " + $EWSservice.url)
			
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
			
			#Enable Mail forwarding to manager.
			Write-Host ("`tForwarding e-mail for $CurrentAccount.SamAccountName to $($UsersManager.Name)") -foregroundcolor "Blue"
			If ($CurrentAccount.ForwardingAddress -eq $null ) {
					If (-Not [string]::IsNullOrEmpty($UsersManager.WindowsEmailAddress.ToString())) {
						$CurrentAccount | Set-Mailbox -DeliverToMailboxAndForward $true -ForwardingSMTPAddress "$($UsersManager.WindowsEmailAddress.ToString())"
						#$CurrentAccount | Set-MailboxAutoReplyConfiguration -AutoReplyState enabled -ExternalAudience "all" -InternalMessage "$CurrentAccount.FirstName is no longer with $Company For any business related needs please e-mail $UsersManager.FirstName at $UsersManager.WindowsEmailAddress." -ExternalMessage "$CurrentAccount.FirstName is no longer with $Company For any business related needs please e-mail $UsersManager.FirstName at $UsersManager.WindowsEmailAddress."
					}	
			}
		}

		#Write-Host ("Testing Mail Name: " + $_.Name + " Alias: " + $($data[0]) + "Disable Date: " + $StrTestDate + " Date Age: " + $TimeSpan.TotalDays)
		
		If ($TimeSpan.TotalDays -ge $PSTExportTime) {
			#Testing to see if is in queue
			If ((Get-MailboxExportRequest | Where-Object { $_.Identity  -contains $($CurrentAccount.Identity)}) -eq $null) {
				Write-Host ("`tExport Mail Name: " + $CurrentAccount.Name + " Alias: " + $CurrentAccount.SamAccountName + " Email: " + $CurrentAccount.WindowsEmailAddress)  -foregroundcolor "Blue"
				#Create New Home Drive
				if (-Not (Test-Path $HomeDriveShare + "\" + $CurrentAccount.SamAccountName)) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $CurrentAccount.SamAccountName)}
				if (-Not (Test-Path $HomeDriveShare + "\" + $CurrentAccount.SamAccountName + "\" + $PSTFolder + "\")) 
					{New-Item -ItemType directory -Path ($HomeDriveShare + "\" + $CurrentAccount.SamAccountName + "\" + $PSTFolder + "\")}
				#Export Mailbox to PST
				New-MailboxExportRequest -Mailbox $_.SamAccountName -FilePath $($HomeDriveShare + "\" + $CurrentAccount.SamAccountName  + "\" + $PSTFolder + "\" + $CurrentAccount.SamAccountName + ".pst")

				while ( (Get-MailboxExportRequestStatistics -Identity $($CurrentAccount.SamAccountName + "\MailboxExport")).status -ne "Completed" ) {
					#View Status of Mailbox Export
					Get-MailboxExportRequestStatistics -Identity $($CurrentAccount.SamAccountName + "\MailboxExport") | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
					Start-Sleep -Seconds 10
				}

				#Remove mailbox from Exchange
				Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
				
				#Move User to Disabled Outlook
				Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
			} else {
				Write-Host ("`t`tUser " + $CurrentAccount.Name + " already submitted.")
				while ((Get-MailboxExportRequestStatistics -Identity ($CurrentAccount.SamAccountName + "\MailboxExport")).status -ne $("Completed")) {
					#View Status of Mailbox Export
					Get-MailboxExportRequestStatistics -Identity ($CurrentAccount.SamAccountName + "\MailboxExport") | ft SourceAlias,Status,PercentComplete,EstimatedTransferSize,BytesTransferred
					Start-Sleep -Seconds 10
				}

				#Remove mailbox from Exchange
				Disable-Mailbox -Identity $CurrentAccount.SamAccountName -confirm:$false
				
				#Move User to Disabled Outlook
				Move-ADObject -Identity $ADUser -TargetPath $DisabledOUDN
				
			}
		}
    }
}
