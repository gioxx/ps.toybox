# Calendar =========================================================================================================================================================

function Clone-OoOMessage {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox from which to copy OoO messages (e.g. source@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox on which to write copied OoO messages (e.g. destination@contoso.com)")]
    [string] $DestinationMailbox
  )

  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {
    $externalMessage = (Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox).ExternalMessage
    $internalMessage = (Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox).InternalMessage

    Set-MailboxAutoReplyConfiguration -Identity $DestinationMailbox -AutoReplyState "Enabled" -InternalMessage $internalMessage -ExternalMessage $externalMessage
    Get-MailboxAutoReplyConfiguration -Identity $DestinationMailbox | Select-Object -Property AutoReplyState,InternalMessage,ExternalMessage

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

}

function Export-CalendarPermission {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Mailbox to analyze (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Domain to analyze (e.g. contoso.com)")]
    [string] $SourceDomain,
    [Parameter(Mandatory=$False, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all mailboxes calendar permissions")]
    [switch] $All
  )

  begin {
    priv_SetPreferences -Verbose
    $eolConnectedCheck = priv_CheckEOLConnection
    $mboxCounter = 0
    $arr_CalPerm = @()

    if ( $eolConnectedCheck -eq $true ) {      
      if ([string]::IsNullOrEmpty($SourceMailbox) -And [string]::IsNullOrEmpty($SourceDomain)) {
        Write-Warning "WARNING: no mailbox(es) or domain(s) specified, I scan all the mailboxes, please be patient."
        $All = $True
      }

      if (-not([string]::IsNullOrEmpty($SourceDomain))) {
        Write-InformationColored "Analyzing mailboxes in $($SourceDomain) ..." -ForegroundColor Cyan
        $SourceMailbox = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue | Where { $_.RecipientTypeDetails -ne "GuestMailUser" -And $_.EmailAddresses -like "*@" + $SourceDomain }
      }

      if ($All) {
        $SourceMailbox = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue
      }
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
      Return
    }
  }

  process {
    $SourceMailbox | ForEach {
      $CurrentUser = $_
      $GetCM = Get-EXOMailbox $CurrentUser -ErrorAction SilentlyContinue

      $mboxCounter++
      $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
      Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      $calendarFolder = Get-MailboxFolderStatistics $GetCM.PrimarySmtpAddress -ErrorAction SilentlyContinue -FolderScope Calendar | 
          Where { $_.FolderType -eq "Calendar" }
      $folderPerms = Get-MailboxFolderPermission "$($GetCM.PrimarySmtpAddress):$($calendarFolder.FolderId)" -ErrorAction SilentlyContinue | 
          Where { $_.AccessRights -notlike "AvailabilityOnly" -and $_.AccessRights -notlike "None" }
      $folderPerms | ForEach {
          $arr_CalPerm += New-Object -TypeName PSObject -Property $([ordered]@{
              PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
              User = $_.User
              Permissions = $_.AccessRights
          })
      }
    }

    if (-not([string]::IsNullOrEmpty($folderCSV)) -Or $All) { 
      $folder = priv_CheckFolder($folderCSV)
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-CalendarPermissions-Report.csv")
      $arr_CalPerm | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_CalPerm | Out-Host
    }

    priv_RestorePreferences
  }
}

function Set-OoO {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox on which to activate OoO (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Disable OoO on specified mailbox")]
    [switch] $Disable
  )

  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {
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
    $Status = "Enabled"

    if ( $Disable ) {
      Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Disabled
      Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox
      break
    }

    $previousMessage = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox | Select -ExpandProperty ExternalMessage
    if ( [string]::IsNullOrEmpty($previousMessage) ) {
      $proposedText = "I'm out of office and will have limited access to my mailbox.<br />
        I will reply to your e-mail as soon as possible.
        <br /><br />
        Have a nice day."
    } else { $proposedText = $previousMessage }
    
    $InternalReply = priv_GUI_TextBox "Out of Office message for internal addresses (same server)" $proposedText
    $ExternalReply = priv_GUI_TextBox "Out of Office message for external addresses (different server)" $InternalReply
    $AbsenceIntervalReply = priv_TakeDecisionOptions "Do you want to specify an absence interval?" "&Yes" "&No" "Specify a period of absence" "Continue without specifying a period of absence"
    
    if ( $AbsenceIntervalReply -eq 0 ) {
      Write-Host "Now select the first day off in the popup and press enter" -f "Yellow"
      $objForm.Text = "Select OoO start date (first day of absence)"
      $objForm.Add_Shown({$objForm.Activate()})
      [void] $objForm.ShowDialog()
      $StartDate = $dtmDate

      if ($StartDate) {
        Write-Host "Start date selected: $StartDate"
        $Status = "Scheduled"
      } else {
        Write-Error "You must select at least one day from the calendar."
        return
      }

      Write-Host "Now select in the popup the last day of absence and press enter" -f "Yellow"
      $objForm.Text = "Select OoO end date (last day of absence)"
      $objForm.Add_Shown({$objForm.Activate()})
      [void] $objForm.ShowDialog()
      $EndDate = $dtmDate
    
      if ($EndDate) {
        Write-Host "Start date selected: $EndDate"
        $Status = "Scheduled"
      } else {
        Write-Error "You must select at least one day from the calendar."
        return
      }
    }

    Switch ($Status) {
        "Scheduled" { Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState "Scheduled" -StartTime $StartDate -EndTime $EndDate -InternalMessage $InternalReply -ExternalMessage $ExternalReply }
        Default { 
          Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState "Enabled" -InternalMessage $InternalReply -ExternalMessage $ExternalReply -ExternalAudience "All" }
    }
    
    Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

}


# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Clone-OoOMessage"
Export-ModuleMember -Function "Export-CalendarPermission"
Export-ModuleMember -Function "Set-OoO"