# Quarantine =======================================================================================================================================================

function Export-QuarantineEML {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="The ID of the message to be exported (example: 20230617142935.F5B74194B266E458@contoso.com)")]
    [string]$messageID,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export the EML file (e.g. C:\Temp)")]
    [string] $folder
  )
  
  if ( -not($messageID.StartsWith('<')) ) { $messageID = '<' + $messageID }
  if ( -not($messageID.EndsWith('>')) ) { $messageID += '>' }
  
  $exportFolder = priv_CheckFolder($folder)
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    if ( (Get-QuarantineMessage -MessageId $($messageID)).QuarantineTypes -eq "Malware" ) {
      $e = Get-QuarantineMessage -MessageId $($messageID) | Export-QuarantineMessage
      $bytes = [Convert]::FromBase64String($e.eml)
      [IO.File]::WriteAllBytes("$($exportFolder)\QuarantineEML.eml", $bytes)
    } else {
      $e = Get-QuarantineMessage -MessageId $($messageID) | Export-QuarantineMessage
      $txt = [System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml))
      [IO.File]::WriteAllText("$($exportFolder)\QuarantineEML.eml", $txt)
    }    

    Invoke-Item "$($exportFolder)\QuarantineEML.eml"
    Start-Sleep -s 3
    Remove-Item "$($exportFolder)\QuarantineEML.eml"
    
    $options_result = priv_TakeDecisionOptions "Should I release the message to all recipients?" "&Yes" "&No" "Release the message." "Do not release the message." 1
    if ($options_result -eq 0) {
      $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1
      if ( $reportFalsePositive -eq 0 ) {
        Get-QuarantineMessage -MessageId $($messageID) | Release-QuarantineMessage -ReleaseToAll -ReportFalsePositive -Confirm:$false
      } else { 
        Get-QuarantineMessage -MessageId $($messageID) | Release-QuarantineMessage -ReleaseToAll
      }
    } else {
      Write-Warning "Message not released (aborted by user)."
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Get-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      ForEach ( $CurrentSender in $SenderAddress ) {
        try {
          Write-Host "Find e-mail(s) from known senders quarantined: e-mail(s) from $($SenderAddress) ..."
          Get-QuarantineMessage -SenderAddress $SenderAddress | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }
  }
}

function Get-QuarantineFromDomain {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail domain in quarantine (e.g. contoso.com)")]
    [string[]]$SenderDomain
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      ForEach ( $CurrentSender in $SenderDomain ) {
        try {
          Write-Host "Find e-mail(s) from known domains quarantined: e-mail(s) from $($SenderDomain) ..."
          Get-QuarantineMessage | Where-Object { $_.SenderAddress -like "*@$($SenderDomain)" } | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }
  }
}

function Get-QuarantineToRelease {
  [CmdletBinding(DefaultParameterSetName='CalendarSetOptions')]
  param (
      [Parameter(Mandatory=$false, ParameterSetName='CalendarSetOptions', ValueFromPipeline=$True, HelpMessage="Choose a single day from calendar")]
      [switch]$ChooseDayFromCalendar,
      [Parameter(Mandatory=$true, ParameterSetName='NoCalendarSetOptions', ValueFromPipeline=$True, HelpMessage="Number of days to be analyzed from today (maximum 30)")]
      [ValidateNotNullOrEmpty()]
      [int]$Interval,
      [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
      [switch] $GridView,
      [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a HTML file")]
      [switch] $HTML,
      [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
      [switch] $CSV
  )

  if ($PSCmdlet.ParameterSetName -eq 'CalendarSetOptions' -and !$ChooseDayFromCalendar) {
      Write-Error "Number of days to be analyzed from today (maximum 30) is a mandatory value (or you must use -ChooseDayFromCalendar)."
      return
  }

  priv_SetPreferences -Verbose

  if ($ChooseDayFromCalendar) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $objForm = New-Object Windows.Forms.Form
    $objForm.Size = New-Object Drawing.Size @(200,190)
    $objForm.StartPosition = "CenterScreen"
    $objForm.KeyPreview = $True

    $objForm.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
          $dtmDate = $objCalendar.SelectionStart
          $objForm.Close()
        }
    })

    $objForm.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") {
          $objForm.Close()
        }
    })

    $objCalendar = New-Object System.Windows.Forms.MonthCalendar
    $objCalendar.ShowTodayCircle = $True
    $objCalendar.MaxSelectionCount = 1
    $objForm.Controls.Add($objCalendar)
    $objForm.Topmost = $True

    New-Variable dtmDate -Option AllScope

    Write-Host "Select the day to be analyzed in the popup and press enter" -f "Yellow"
    $objForm.Text = "Select the day to be analyzed"
    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    if ($dtmDate) {
      $startDate = $dtmDate.AddDays(-1)
      $endDate = ($dtmDate)
    } else {
      Write-Error "You must select at least one day from the calendar."
      return
    }
  }

  if ($Interval) { 
    if ( $Interval -gt 30 ) { 
      $Interval = 30
    } else { 
      $Interval = $($Interval)
    }

    $startDate = (Get-Date).AddDays(-$Interval)
    $endDate = Get-Date
  }

  $arr_QuarantineToRelease = @()
  $ReleaseQuarantinePreview = @()
  $ReleaseQuarantineReleased = @()
  $ReleaseQuarantineDeleted = @()
  $Page = 1
  
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    Write-InformationColored "Quarantine report from $($startDate.Date) to $($endDate)" -ForegroundColor "Yellow"
    
    do {
      # Credits: https://community.spiceworks.com/topic/2343368-merge-eop-quarantine-pages#entry-9354845
      $QuarantinedMessages = Get-QuarantineMessage -StartReceivedDate $startDate.Date -EndReceivedDate $endDate -PageSize 1000 -ReleaseStatus NotReleased -Page $Page
      $Page++
      $QuarantinedMessagesAll += $QuarantinedMessages
    } until ( $QuarantinedMessages -eq $null )

    Write-InformationColored "Total items: $($QuarantinedMessagesAll.Count)" -ForegroundColor "Yellow"

    $QuarantinedMessagesAll | ForEach {
      $Message = $_
      $arr_QuarantineToRelease += New-Object -TypeName PSObject -Property $([ordered]@{
        SenderAddress = $Message.SenderAddress
        RecipientAddress = $Message.RecipientAddress
        Subject = $Message.Subject
        ReceivedTime = $Message.ReceivedTime
        QuarantineTypes = $Message.QuarantineTypes
        Released = $Message.Released
        MessageId = $Message.MessageId
        Identity = $Message.Identity
      })
    }

    if ( $GridView ) {
      # Credits: https://stackoverflow.com/a/51033908
      $ReleaseQuarantine = $arr_QuarantineToRelease | Sort-Object -Descending ReceivedTime | Out-GridView -Title "$($startDate.Date) to $($endDate) • $($Interval) days • $($QuarantinedMessagesAll.Count) items" -PassThru

      $ProcessedCount = 0
      
      if ( $ReleaseQuarantine -ne $null ) {
        if ( $ReleaseQuarantine.Count -eq 1 ) {
          $single_menu = priv_TakeDecisionOptions "What do you want to do with $($ReleaseQuarantine.Subject)?" "&Release" "&Analyze" "Release the message." "Export a copy of the EML file for analysis."

          Switch ( $single_menu ) {
            0 { 
                $decision = priv_TakeDecisionOptions "Do you really want to release $($ReleaseQuarantine.Subject)?" "&Yes" "&No" "Release the message." "Do not release message."
                if ($decision -eq 0) {
                  
                  $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1
                  if ( $reportFalsePositive -eq 0 ) {
                    Release-QuarantineMessage -Identity $ReleaseQuarantine.Identity -ReleaseToAll -ReportFalsePositive -Confirm:$false
                  } else { 
                    Release-QuarantineMessage -Identity $ReleaseQuarantine.Identity -ReleaseToAll -Confirm:$false
                  }
                  
                  $released = Get-QuarantineMessage -Identity $ReleaseQuarantine.Identity
                  
                  $releasedResults = @()
                  $releasedResults += New-Object -TypeName PSObject -Property $([ordered]@{
                    Subject = priv_MaxLenghtSubString $released.Subject 40
                    SenderAddress = priv_MaxLenghtSubString $released.SenderAddress $MaxFieldLength
                    Released = $released.Released
                    ReleasedUser = $released.ReleasedUser
                  })
                  $releasedResults | Sort-Object Subject | Out-Host
                }
              }
            1 { Export-QuarantineEML -messageID "$($ReleaseQuarantine.MessageId)" }
          }
        } else {
          $ReleaseQuarantine | ForEach {
            $QuarantinedMessage = $_
            $ReleaseQuarantinePreview += New-Object -TypeName PSObject -Property $([ordered]@{
              Subject = priv_MaxLenghtSubString $QuarantinedMessage.Subject 50
              SenderAddress = priv_MaxLenghtSubString $QuarantinedMessage.SenderAddress $MaxFieldLength
              Released = $QuarantinedMessage.Released
            })
          }
          
          Write-Host "`n$($ReleaseQuarantine.Count) items selected, take a look at the preview below:" -f "Cyan"
          $ReleaseQuarantinePreview | Sort-Object Subject | Select-Object Subject,SenderAddress,Released | Out-Host

          $release_or_delete = priv_TakeDecisionOptions "Do you want to release or delete $($ReleaseQuarantine.Count) selected items?" "&Release" "&Delete" "Release messages" "Delete messages"
          
          if ( $release_or_delete -eq 1 ) {
            # DELETE QUARANTINED EMAILS SELECTED
            $decision = priv_TakeDecisionOptions "Do you really want to permanently delete $($ReleaseQuarantine.Count) selected items?" "&Yes" "&No" "Delete message(s)." "Do not delete message(s)."

            if ($decision -eq 0) {
              $ReleaseQuarantine | ForEach {
                $QuarantinedMessageToDelete = $_
                # $QuarantinedMessageToDelete = Get-QuarantineMessage -Identity $_.Identity
                
                $ProcessedCount++
                $PercentComplete = ( ($ProcessedCount / $ReleaseQuarantine.Count) * 100 )
                Write-Progress -Activity "Deleting $(priv_MaxLenghtSubString $QuarantinedMessageToDelete.Subject $MaxFieldLength)" -Status "$ProcessedCount out of $($ReleaseQuarantine.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

                # $QuarantinedMessageToDelete | Delete-QuarantineMessage -Confirm:$false
                Delete-QuarantineMessage -Identity $QuarantinedMessageToDelete.Identity -Confirm:$false
              }
              Write-Host "`nDone.`n" -f "Green"
            }


          } else {
            # RELEASE QUARANTINED EMAILS SELECTED
            $decision = priv_TakeDecisionOptions "Do you really want to release $($ReleaseQuarantine.Count) selected items?" "&Yes" "&No" "Release messages." "Do not release messages."
            $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1

            if ( $decision -eq 0 ) {
              $ReleaseQuarantine | ForEach {                
                $QuarantinedMessageToRelease = $_
                # $QuarantinedMessageToRelease = Get-QuarantineMessage -Identity $_.Identity

                $ProcessedCount++
                $PercentComplete = (($ProcessedCount / $ReleaseQuarantine.Count) * 100)
                Write-Progress -Activity "Processing $(priv_MaxLenghtSubString $QuarantinedMessageToRelease.Subject $MaxFieldLength)" -Status "$ProcessedCount out of $($ReleaseQuarantine.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

                if ( $reportFalsePositive -eq 0 ) {
                  Release-QuarantineMessage -Identity $QuarantinedMessageToRelease.Identity -ReleaseToAll -ReportFalsePositive -Confirm:$false
                } else { 
                  Release-QuarantineMessage -Identity $QuarantinedMessageToRelease.Identity -ReleaseToAll -Confirm:$false
                }
                
                $ReleaseQuarantineReleased += New-Object -TypeName PSObject -Property $([ordered]@{
                  Subject = priv_MaxLenghtSubString $QuarantinedMessageToRelease.Subject $MaxFieldLength
                  SenderAddress = priv_MaxLenghtSubString $QuarantinedMessageToRelease.SenderAddress $MaxFieldLength
                  # Released = $QuarantinedMessageToRelease.Released
                  # ReleasedUser = $QuarantinedMessageToRelease.ReleasedUser
                })
              }

              Write-Host "Done, please take a look below." -f "Green"
              $ReleaseQuarantineReleased | Select-Object Subject,SenderAddress,Released,ReleasedUser | Sort-Object Subject | Out-Host
            }
          }
        }
      }
    } elseif ( $CSV ) {
      $folder = priv_CheckFolder($folderCSV)
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-QuarantineToRelease-Report.csv")
      $arr_QuarantineToRelease | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } elseif ( $HTML ) {
      if (-not(Get-Module -Name PSWriteHTML)) { 
        if (Get-Module -ListAvailable -Name PSWriteHTML) {
          Import-Module PSWriteHTML
        } else {
          Write-InformationColored "PSWriteHTML module is not available." -ForegroundColor "Yellow"
          $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No "
          if ( $Confirm -match "[yY]" ) {
            try {
              Write-InformationColored "Installing PSWriteHTML module ..." -ForegroundColor "Yellow"
              Install-Module PSWriteHTML -Scope CurrentUser -AllowClobber -Force
            } catch {
              Write-Error "`nCan't install PSWriteHTML module. `nPlease check logs."
              exit
            }
          } else {
            Write-Error "PSWriteHTML module is required to run this script. Please install it first using Install-Module PSWriteHTML cmdlet."
            exit
          }
        }
      }
      $arr_QuarantineToRelease | Out-GridHtml
    } else {
      $arr_QuarantineToRelease | Select-Object SenderAddress,RecipientAddress,Subject,QuarantineTypes,Released | Sort-Object Subject
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
}

function Release-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    if ( $eolConnectedCheck -eq $true ) {

      $releasedResults = @()
      $SenderAddress | ForEach {
        try {
          $CurrentSender = $_
          Write-Host "Release quarantine from known senders: release e-mail(s) from $($CurrentSender) ..."
          Get-QuarantineMessage -SenderAddress $CurrentSender | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Where-Object { $null -ne $_.QuarantinedUser -and $_.ReleaseStatus -ne "RELEASED" } | 
              Release-QuarantineMessage -ReleaseToAll
          Get-QuarantineMessage -SenderAddress $CurrentSender | 
            ForEach { 
              $released = Get-QuarantineMessage -Identity $_.Identity
              $row = "" | Select-Object Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
              $row.Subject = priv_MaxLenghtSubString $released.Subject 40
              $row.SenderAddress = $released.SenderAddress
              $row.ReceivedTime = $released.ReceivedTime
              $row.Released = $released.Released
              $row.ReleasedUser = $released.ReleasedUser
              $releasedResults += $row
            } 
          $releasedResults | Format-Table -AutoSize
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }
  }
}
# Credits: https://www.jonathanmedd.net/2014/03/using-powershell-aliases-in-a-module.html
Set-Alias -Name rqf -Value Release-QuarantineFrom -Description "Release Quarantine from (function)"

function Release-QuarantineMessageId {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="ID of the message locked in quarantine (e.g. CAH_w85uSio_cz4HsFxJAGQDd-kzxGijLaMagZU95m3A1G8hWBA@mail.contoso.com)")]
    [string[]]$MessageId
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    
    if ( $eolConnectedCheck -eq $true ) {
      $releasedResults = @()
      $MessageId | ForEach {
        try {
          $CurrentMessage = "<$($_)>"
          $ReleaseId = Get-QuarantineMessage -MessageId $CurrentMessage | 
              Where-Object { $null -ne $_.QuarantinedUser -and $_.ReleaseStatus -ne "RELEASED" }
          if ( $ReleaseId.Count -ge 1) {
            $ReleaseId | Release-QuarantineMessage -ReleaseToAll
            Write-Host "Release quarantine message with id $($CurrentMessage) ..."
            Get-QuarantineMessage -MessageId $CurrentMessage | 
              ForEach { 
                $released = Get-QuarantineMessage -Identity $_.Identity
                $row = "" | Select-Object Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
                $row.Subject = priv_MaxLenghtSubString $released.Subject 40
                $row.SenderAddress = $released.SenderAddress
                $row.ReceivedTime = $released.ReceivedTime
                $row.Released = $released.Released
                $row.ReleasedUser = $released.ReleasedUser
                $releasedResults += $row
              }
            $releasedResults | Format-Table -AutoSize
          } else {
            Write-Warning "No quarantined messages to release with id $($CurrentMessage) (already released or not found yet)."
          }
        } catch {
          Write-Error $_.Exception.Message
        }
      }
      
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }
  }
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Export-QuarantineEML"
Export-ModuleMember -Function "Get-QuarantineFrom"
Export-ModuleMember -Function "Get-QuarantineFromDomain"
Export-ModuleMember -Function "Get-QuarantineToRelease"
Export-ModuleMember -Function "Release-QuarantineFrom"
Export-ModuleMember -Function "Release-QuarantineMessageId"