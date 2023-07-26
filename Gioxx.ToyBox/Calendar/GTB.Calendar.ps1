# Calendar =========================================================================================================================================================

function Export-CalendarPermission {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Mailbox to analyze (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all mailboxes calendar permissions")]
    [switch] $All
  )

  begin {
    $mboxCounter = 0
    $Result = @()
    Set-Variable ProgressPreference Continue

    if ([string]::IsNullOrEmpty($SourceMailbox)) {
      Write-Host "WARNING: no mailbox(es) specified, I scan all the mailboxes, please be patient." -f "Yellow"
      $All = $True
    }

    if ($All) {
      $SourceMailbox = Get-Mailbox -ResultSize Unlimited
    }
  }

  process {
    $SourceMailbox | ForEach {
      $CurrentUser = $_
      $GetCM = Get-Mailbox $CurrentUser -ErrorAction SilentlyContinue

      $mboxCounter++
      $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
      Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      $calendarFolder = Get-MailboxFolderStatistics $GetCM.PrimarySmtpAddress -ErrorAction SilentlyContinue -FolderScope Calendar | 
          Where { $_.FolderType -eq "Calendar" }
      $folderPerms = Get-MailboxFolderPermission "$($GetCM.PrimarySmtpAddress):$($calendarFolder.FolderId)" -ErrorAction SilentlyContinue | 
          Where { $_.AccessRights -notlike "AvailabilityOnly" -and $_.AccessRights -notlike "None" }
      $folderPerms | ForEach {
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
              PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
              User = $_.User
              Permissions = $_.AccessRights
          })
      }
    }

    if (-not([string]::IsNullOrEmpty($folderCSV)) -Or $All) { 
      $folder = priv_CheckFolder($folderCSV)
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-CalendarPermissions-Report.csv")
      $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $Result
    }
  }
}

function Set-OoO {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox on which to activate OoO (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Disable OoO on specified mailbox")]
    [switch] $Disable
  )
  
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

  if ( $Disable ) {
    Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Disabled
    Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox
    break
  }

  $previousMessage = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox | Select -ExpandProperty ExternalMessage
  if ( [string]::IsNullOrEmpty($previousMessage) ) {
    $proposedText = "I'm out of office and will have limited access to my mailbox.<br />
      I will reply to your email as soon as possible.
      <br /><br />
      Have a nice day."
  } else { $proposedText = $previousMessage }
  
  $InternalReply = priv_GUI_TextBox "Out of Office message for internal addresses (same server)" $proposedText
  $ExternalReply = priv_GUI_TextBox "Out of Office message for external addresses (different server)" $InternalReply

  try {
    $AbsenceIntervalMsg = "Do you want to specify an absence interval?"
    $option_n = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Continue without specifying a period of absence"
    $option_y = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Specify a period of absence"
    $AbsenceIntervalOpt = [System.Management.Automation.Host.ChoiceDescription[]]($option_y, $option_n)
    $AbsenceIntervalReply = $host.ui.PromptForChoice("", $AbsenceIntervalMsg, $AbsenceIntervalOpt, 1)
    
    if ( $AbsenceIntervalReply -eq 0 ) {
      Write-Host "Now select the first day off in the popup and press enter" -f "Yellow"
      $objForm.Text = "Select OoO start date (first day of absence) or press ESC to ignore"
      $objForm.Add_Shown({$objForm.Activate()})
      [void] $objForm.ShowDialog()
      $StartDate = $dtmDate
      if ($dtmDate) {
        Write-Host "Start date selected: $dtmDate"
      }

      Write-Host "Now select in the popup the last day of absence and press enter" -f "Yellow"
      $objForm.Text = "Select OoO end date (last day of absence) or press ESC to ignore"
      $objForm.Add_Shown({$objForm.Activate()})
      [void] $objForm.ShowDialog()
      $EndDate = $dtmDate
      if ($dtmDate) {
        Write-Host "End date selected: $dtmDate"
      }
    }

    if ([string]::IsNullOrEmpty($Start) -eq $False -And [string]::IsNullOrEmpty($End) -eq $False ) {
      $Status = "Scheduled"
      #Write-Host "DEBUG: Scheduled" -f "Yellow"
    } else {
      $Status = "Enabled"
      #Write-Host "DEBUG: Enabled" -f "Yellow"
    }

    Switch ($Status) {
        "Enabled" {
          #Write-Host "DEBUG: Int- $InternalReply" -f "Yellow"
          #Write-Host "DEBUG: Ext- $ExternalReply" -f "Yellow"
          Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Enabled -InternalMessage $InternalReply -ExternalMessage $ExternalReply
        }
        "Scheduled" {
          #Write-Host "DEBUG: Int- $InternalReply" -f "Yellow"
          #Write-Host "DEBUG: Ext- $ExternalReply" -f "Yellow"
          Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Scheduled -StartTime $StartDate -EndTime $EndDate -InternalMessage $InternalReply -ExternalMessage $ExternalReply
        }
    }
    Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox
  } catch {
    Write-Error $_.Exception.Message
  }

}


# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Export-CalendarPermission"
Export-ModuleMember -Function "Set-OoO"