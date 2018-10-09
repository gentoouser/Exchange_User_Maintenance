#Region Move all users one at a time to new MailStore
function MoveMailbox(
	[string]$SourceDatabase="",
	[string]$DestinationDatabase="",
	[string]$LogPath="D:\",
	[Int32]$Limit = 0,
	[Int32]$Sleep = 120,
	[Int32]$LowSpaceStop = 50,
	[Switch]$Descending = $false
	)
	{
		#Start Logging
		try { 
			Start-Transcript -append -path $($LogPath + "Move_From_" + $SourceDatabase + "_to_" + $DestinationDatabase + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".log")

		} catch { 
			Stop-transcript
			Start-Transcript -append -path $($LogPath + "Move_From_" + $SourceDatabase + "_to_" + $DestinationDatabase + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".log")
		} 
		Write-Host ("="*[console]::BufferWidth)
		Write-Host "Source Database: $SourceDatabase"
		Write-Host "Destination Database: $DestinationDatabase"
		Write-Host ("="*[console]::BufferWidth)
		#Get Database drive info
		$DDObject = Get-MailboxDatabase $DestinationDatabase
		
		#Get Mailboxes in Database
		If ( $Limit -gt 0) {
			If($Descending){
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | select -first $Limit | Sort-Object -Descending)
			}else{
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | select -first $Limit)
			}
		}else{
			If($Descending){
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | Sort-Object -Descending)
			}else{
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase)
			}
		}
		
		$mailboxmove | foreach {
			Write-Host ("-"*[console]::BufferWidth)
			#Start Move
			New-MoveRequest -Identity $_ -TargetDatabase $DestinationDatabase
			Write-Host ("-"*[console]::BufferWidth)
			# Monitor Move
			While ($moveRequests = Get-MoveRequest -Identity $_ | Where-Object {$_.Status -eq "InProgress" -or $_.Status -eq "Queued"}) {
				foreach($moveRequest in $moveRequests) {
					$CurrentStatus = Get-MoveRequestStatistics -Identity $_  | select DisplayName, PercentComplete, BadItemsEncountered, TotalMailboxSize, TotalMailboxItemCount, TotalInProgressDuration,TotalSuspendedDuration,TotalQueuedDuration
					Write-Progress -Id 0 -Activity ("Source: " +  $SourceDatabase + "`t Destination: " + $DestinationDatabase) -status ("Processing User: " + $CurrentStatus.DisplayName + "Mail Size: " + $CurrentStatus.TotalMailboxSize + " Progress Duration: " + $CurrentStatus.TotalInProgressDuration ) -percentComplete $CurrentStatus.PercentComplete 
					$host.ui.RawUI.WindowTitle = ( "Source: " +  $SourceDatabase + " Destination: " + $DestinationDatabase + " User: " + $CurrentStatus.DisplayName + " Progress: " + $CurrentStatus.PercentComplete + "%")
				}
			#Write-Host "`t===Waiting for $Sleep seconds==="
			
			
			Start-Sleep $Sleep;	
			} 
			#Get Drive Space
			$ht = @{} 
			$objDrives = Get-WmiObject Win32_Volume -Filter "DriveType='3'"
			foreach ($ObjDisk in $objDrives) 
			{ 
				$size = ([Math]::Round($ObjDisk.Capacity /1GB,2))
				$freespace = ([Math]::Round($ObjDisk.FreeSpace /1GB,2))
				$IntUsed = ([Math]::Round($($size – $freeSpace),2))
				$IntPercentFree = ([Math]::Round(($freeSpace/$size)*100,2))
				$ht.Add($ObjDisk.Name,@{Lable = $objDisk.Label; PercentFree = $IntPercentFree; Size = $size; FreeSpace = $freespace; Used = $IntUsed }) 
			}
			#Test for freespace
			(Split-Path -Parent -Path $DDObject.EdbFilePath.PathName)
			# $DatabaseUsedPercent = ($ht.GetEnumerator() | ?{(Split-Path -Parent -Path $DDObject.EdbFilePath.PathName) -like ($_.key + "*") -and $_.key.tostring().length -gt 3}).value.PercentFree
			$LogUsedPercent = ($ht.GetEnumerator() | ?{$DDObject.LogFolderPath.PathName -like ($_.key + "*") -and $_.key.tostring().length -gt 3}).value.PercentFree
			If ($LogUsedPercent -le $LowSpaceStop){
				Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
				Write-Host "!!! Stopping Due to Low Space on $DDObject.LogFolderPath.PathName !!!" -ForegroundColor Red
				Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
				exit; 
			}
		}
		Stop-transcript
	}
