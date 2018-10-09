<# 
.SYNOPSIS
    Name: Exchange_Export_Disabled_Mailboxes.ps1
    Disabled User MailBox Export

.DESCRIPTION
	*Find All Disabled Users in Active Directory
	*Filters for Mailboxes 
	*See if other users have full mailbox rights
    *Export mailbox to PST

    .DEPENDENCIES
    *Active Directory Module
    *Exchange Module

.PARAMETER Archive
    UNC path for where .pst will be exported to.
.PARAMETER Server
    FQDN of Exchange server.
.PARAMETER Wait
    Time to wait for PST export.
.EXAMPLE
   & Exchange_Export_Disabled_Mailboxes.ps1 -Archive \\remoteserver\share

.NOTES
 Author: Paul Fuller
 Changes:
  1.0.1 - Fixing Display issues. 
	1.0.2 - Fixed Dispalay issue with Mailbox Permissions.
	1.0.3 - Updated Progress Display both Progressbars and Create sub-folder called Logs for log files
#>
PARAM (
    [Parameter(Mandatory=$true)][string]$Archive = $null,
    [Parameter(Mandatory=$true)][string]$Server = $null,
    [Parameter(Mandatory=$false)][int]$Wait = 120
)
$ScriptVersion = "1.0.3"
#############################################################################
#region User Variables
#############################################################################
$ExcludeUsers=@(
"<Domian>\Domain Admins",
"<Domian>\Enterprise Admins",
"<Domian>\Organization Management",
"<Domian>\Exchange Servers",
"<Domian>\Exchange Domain Servers",
"NT AUTHORITY\SYSTEM",
"NT AUTHORITY\SELF"
)
$LogFile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\Logs\" + `
		$MyInvocation.MyCommand.Name + "_" + `
		(Get-Date -format yyyyMMdd-hhmm) + ".log")
$sw = [Diagnostics.Stopwatch]::StartNew()
$ID = 1
#############################################################################
#endregion User Variables
#############################################################################

#############################################################################
#region Setup Sessions
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	If (-Not( Test-Path (Split-Path -Path $LogFile -Parent))) {
		New-Item -ItemType directory -Path (Split-Path -Path $LogFile -Parent)
	}
	try { 
	Start-Transcript -Path $LogFile -Append
	} catch { 
		Stop-transcript
		Start-Transcript -Path $LogFile -Append
	} 
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
##Load Exchange Module
# Load All Exchange PSSnapins 
If ((Get-PSSession | Where-Object { $_.ConfigurationName -Match "Microsoft.Exchange" }).Count -eq 0 ) {
	Write-Host ("Loading Exchange Plugins") -foregroundcolor "Green"
	If ($([System.Net.Dns]::GetHostByName(($env:computerName))).hostname -eq $([System.Net.Dns]::GetHostByName(($Server))).hostname) {
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
		. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer -auto -AllowClobber
	} else {
		$ERPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Server/PowerShell/ -Authentication Kerberos
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

#Get All Disabled accounts in AD
Write-Host ("Getting Disabled Accounts. Please wait . . .")
#$DisabledAccounts = Search-ADAccount -AccountDisabled -Usersonly |  Get-Aduser -Properties name,msExchMailboxGuid,sAMAccountName |  Where-Object {$_.msExchMailboxGuid -ne $null}
$DisabledAccounts = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited | where {$_.UseraccountControl -like “*accountdisabled*”}
#############################################################################
#endregion Setup Sessions
#############################################################################

#############################################################################
#region Functions
#############################################################################
Function FormatElapsedTime($ts) 
{
    #https://stackoverflow.com/questions/3513650/timing-a-commands-execution-in-powershell
	$elapsedTime = ""

    if ( $ts.Hours -gt 0 )
    {
        $elapsedTime = [string]::Format( "{0:00} hours {1:00} min. {2:00}.{3:00} sec.", $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
    }else {
        if ( $ts.Minutes -gt 0 )
        {
            $elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
        }
        else
        {
            $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
        }

        if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
        }

        if ($ts.Milliseconds -eq 0)
        {
            $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
        }
    }
    return $elapsedTime
}
Function Export-Mail($User,$Archive) {
    #Get User Mailbox object
    $ObjUser = Get-User $User
		If ($ObjUser.RecipientType -eq "UserMailbox" ) {
			$CurrentMailBox = $ObjUser | Get-Mailbox
			#Testing to see if is in queue
			If ((Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity}).count -eq 0) {
				Write-Host ("`tExport Mail Name: " + $ObjUser.Name + " Alias: " + $ObjUser.SamAccountName + " Email: " + $ObjUser.WindowsEmailAddress)  -foregroundcolor "Cyan"
				#Create Archive path if not created
				if (-Not (Test-Path $Archive )) {
                    New-Item -ItemType directory -Path $Archive | Out-Null
                    If (-Not $?) {
                         Write-Warning ("Path not valid: $Archive")
                        Return
                    }
                }
                #test to see if User has been exported
                if (Test-Path ($Archive + "\" + $($ObjUser.SamAccountName) + ".pst") ) {
                        Write-Warning ("User: " + $ObjUser.SamAccountName + " already has been exported to: " + $($ObjUser.SamAccountName) + ".pst")
                        Return
                }
				#Export Mailbox to PST
				$MER = New-MailboxExportRequest -Mailbox $ObjUser.SamAccountName -FilePath $($Archive + "\" + $($ObjUser.SamAccountName) + ".pst")
				Start-Sleep -Seconds 15
			} else {
				Write-Host ("`t`tUser " + $ObjUser.Name + " already submitted. ")
			}
			#Monitor Export	
            $ExportJobStatusName = $null
			$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
			If ($ExportJobStatusName -ne $null) {
				#Write-Host ("`t`t`t Job Status loop: " + $ExportJobStatusName.status)
				while (($ExportJobStatusName.status -ne "Completed") -And ($ExportJobStatusName.status -ne "Failed")) {
					#View Status of Mailbox Export
					$ExportJobStatusName = Get-MailboxExportRequest | Where-Object { $_.mailbox -eq $CurrentMailBox.Identity -and $_.status -ne 10 } | select -first 1 | Get-MailboxExportRequestStatistics 
					Write-Progress -Id ($Id+1) -Activity $("Exporting user: " + $ExportJobStatusName.SourceAlias ) -status $("Export Percent Complete: " + $ExportJobStatusName.PercentComplete + " Copied " + $ExportJobStatusName.BytesTransferred + " out of " + $ExportJobStatusName.EstimatedTransferSize ) -percentComplete $ExportJobStatusName.PercentComplete
					Start-Sleep -Seconds 15
				}
			}

            #Check for Completion status
			$ExportMailBoxList = Get-MailboxExportRequest | Where-Object { $_.Mailbox -eq $CurrentMailBox.Identity -And ($_.status -ne "Completed" -Or $_.status -ne "Failed")}
				
            If ($ExportMailBoxList.status -eq "Completed") {
				#Remove Exchange account of PST was successful. 
                Write-Host ("`t`t Removing Mailbox from Exchange: " + $CurrentMailBox.Identity)
                Disable-Mailbox -Identity $ObjUser.SamAccountName -confirm:$false
                $ExportMailBoxList | Remove-MailboxExportRequest -Confirm:$false
			}
            #Stop if PST Export failed.
			If ($ExportMailBoxList.status -eq "Failed") {
                throw ("PST Export failed: " + $error[0].Exception)
                Break
            }
		}
}
#############################################################################
#endregion Functions
#############################################################################

#############################################################################
#region Main
#############################################################################
$AtE = 0
$NAtE = 0

ForEach ($DA in $DisabledAccounts) {
    $FixedDAMP = @{}
    Write-Progress -Id 0 -Activity $("Processing User: " + $DA.Name ) -status $("User: " + ($AtE + $NAtE) + " out of " + $DisabledAccounts.count ) -percentComplete ((($AtE + $NAtE)/$DisabledAccounts.count)*100) 
    Write-Host ("Processing User: " + $DA.Name) -ForegroundColor DarkGray
    $DAMP = $DA | Get-MailboxPermission | where {($_.AccessRights -eq "FullAccess") -and ($_.User -notin $ExcludeUsers) -and ($_.User -notmatch "S-1-5-*" )}
    #Mailboxes that other have permission to. 
	If ($DAMP.count -gt 0 ) {
        ForEach( $ACE in $DAMP) {
            If ((Get-User $ACE.User).UseraccountControl  -notlike “*accountdisabled*”) {
                $FixedDAMP.add($ACE.User,$ACE.AccessRights)
            }
        }

       If ($FixedDAMP.count -eq 0 ) {
			#Continue Export

			#Create Archive Folder
			if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
				New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
			}
			#Start Mail Export
			Export-Mail $DA.SamAccountName ($Archive  + "\" + $DA.SamAccountName)
       }else {
			Write-Host ("`tMailBox Permissions Count: " + $DAMP.count) -ForegroundColor Red
			Write-Host ("`tMailBox Permissions Fixed Count: " + $FixedDAMP.count) -ForegroundColor yellow
			$FixedDAMP | ft
            $NAtE ++
       }
        
    }else {
        #Continue Export

        #Create Archive Folder
		if (-Not (Test-Path ($Archive + "\" + $DA.SamAccountName))) {
			New-Item -ItemType directory -Path ($Archive  + "\" + $DA.SamAccountName) | Out-Null
		}
        #Start Mail Export
        Export-Mail $DA.SamAccountName ($Archive  + "\" + $DA.SamAccountName)
        $AtE ++
    }
}

$sw.Stop()
Write-Host ("Script took: " + (FormatElapsedTime($sw.Elapsed)) + " to run.")
#############################################################################
#endregion Main
#############################################################################
If (-Not [string]::IsNullOrEmpty($LogFile)) {
	Stop-Transcript
}
