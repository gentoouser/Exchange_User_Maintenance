#region Move all users one at a time to new MailStore
function MoveMailbox(
	[string]$SourceDatabase="",
	[string]$DestinationDatabase="",
	[string]$LogPath="D:\",
	[Int32]$Limit = 0,
	[Int32]$Sleep = 120,
	[Int32]$LowSpaceStopEDB = 20,
	[Int32]$LowSpaceStopLog = 50,
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
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | Select-Object -first $Limit | Sort-Object -Descending)
			}else{
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | Select-Object -first $Limit)
			}
		}else{
			If($Descending){
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase | Sort-Object -Descending)
			}else{
				$mailboxmove = (Get-mailbox -resultsize unlimited -database $SourceDatabase)
			}
		}
		#Get count of mailboxes to move
		$count = 1
		[Int32]$mailboxcount = $mailboxmove.count
		If ( $mailboxcount  -gt 0) {
			$mailboxmove | foreach {
				#Get Drive Space
				$ht = @{} 
				$objDrives = Get-WmiObject Win32_Volume -Filter "DriveType='3'"
				foreach ($ObjDisk in $objDrives) 
				{ 
					$size = ([Math]::Round($ObjDisk.Capacity /1GB,2))
					$freespace = ([Math]::Round($ObjDisk.FreeSpace /1GB,2))
					$IntUsed = ([Math]::Round($($size – $freeSpace),2))
					$IntPercentFree = ([Math]::Round(($freeSpace/$size)*100,2))
					$ht.Add($ObjDisk.Name,@{Label = $objDisk.Label; PercentFree = $IntPercentFree; Size = $size; FreeSpace = $freespace; Used = $IntUsed }) 
				}
				#Test for FreeSpace
				#(Split-Path -Parent -Path $DDObject.EdbFilePath.PathName)
				$EbdUsedPercent = ($ht.GetEnumerator() | Where-Object {$DDObject.EdbFilePath.PathName -like ($_.key + "*") -and $_.key.tostring().length -gt 3}).value.PercentFree
				If ($EbdUsedPercent -le $LowSpaceStopEDB){
					Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
					Write-Host ("!!! Stopping Due to Low Space on " + $DDObject.EdbFilePath.PathName + "!!!") -ForegroundColor Red
					Write-Host ("Drive is at: " + $EbdUsedPercent + "% Free" ) -ForegroundColor Red
					Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
					break; 
				}				
				$LogUsedPercent = ($ht.GetEnumerator() | Where-Object {$DDObject.LogFolderPath.PathName -like ($_.key + "*") -and $_.key.tostring().length -gt 3}).value.PercentFree
				If ($LogUsedPercent -le $LowSpaceStopLog){
					Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
					Write-Host ("!!! Stopping Due to Low Space on " + $DDObject.LogFolderPath.PathName + "!!!") -ForegroundColor Red
					Write-Host ("Drive is at: " + $LogUsedPercent + "% Free" ) -ForegroundColor Red
					Write-Host ("="*[console]::BufferWidth) -ForegroundColor Red
					break; 
				}
				#Update status
				Write-Progress -Id 0 -Activity ("Source: " +  $SourceDatabase + "      Destination: " + $DestinationDatabase)  -status ("Moving: " + $count + " of " + $mailboxcount + " [Database Drive Free: " + $EbdUsedPercent + "%] [Log Drive Free: " + $LogUsedPercent + "%]") -percentComplete ($count/$mailboxcount)
				Write-Host ("-"*[console]::BufferWidth)
				#Start Move
				New-MoveRequest -Identity $_ -TargetDatabase $DestinationDatabase
				Write-Host ("-"*[console]::BufferWidth)
				# Monitor Move
				While ($moveRequests = Get-MoveRequest -Identity $_ | Where-Object {$_.Status -eq "InProgress" -or $_.Status -eq "Queued"}) {
					foreach($moveRequest in $moveRequests) {
						$CurrentStatus = Get-MoveRequestStatistics -Identity $_  | Select-Object DisplayName, PercentComplete, BadItemsEncountered, TotalMailboxSize, TotalMailboxItemCount, TotalInProgressDuration,TotalSuspendedDuration,TotalQueuedDuration
						Write-Progress -Id 1 -Activity ("Processing User: " + $CurrentStatus.DisplayName)  -status ("Mail Size: " + $CurrentStatus.TotalMailboxSize + "      Progress Duration: " + $CurrentStatus.TotalInProgressDuration ) -percentComplete $CurrentStatus.PercentComplete 
						
						$host.ui.RawUI.WindowTitle = ( "Source: " +  $SourceDatabase + "      Destination: " + $DestinationDatabase + "      User: " + $CurrentStatus.DisplayName + "      Progress: " + $CurrentStatus.PercentComplete + "%      Total Progress: " + ([Math]::Round(($count/$mailboxcount)*100)) + "%")
					}
				#Write-Host "`t===Waiting for $Sleep seconds==="
				
				
				Start-Sleep $Sleep;	
				} 
				
				$count++
			}
		} else {
			Write-Warning ("No More Mailboxes to Move on MailStore" )
		}
		Stop-transcript
	}

# MoveMailbox -SourceDatabase "SourceMailStore" -DestinationDatabase "DestinationMailStore" -Limit 10 -Descending
#endregion Move all users one at a time to new MailStore
