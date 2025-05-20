# Statistics =======================================================================================================================================================

function Export-MboxStatistics {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Single user to analyze (e.g. mario.rossi@contoso.com)")]
    [string] $user,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Round up the values of ArchiveWarningQuotaInGB and ArchiveQuotaInGB (by excess)")]
    [switch] $Round
  )

  priv_SetPreferences -Verbose
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    $folder = priv_CheckFolder($folderCSV)
    $arr_MbxStats = @()
    $ProcessedCount = 0

    if ( [string]::IsNullOrEmpty($user) ) { 
      $Mailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue
      $WriteToCSV = $True
    } else { 
      $Mailboxes = Get-Mailbox $user
      $WriteToCSV = $False
    }

    $TotalMailboxes = $Mailboxes.Count

    if ( $WriteToCSV ) { 
      $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format 'yyyyMMdd').ToString())_M365-MailboxStatistics.csv")
      Write-InformationColored "Saving report to $($CSV)" -ForegroundColor "Yellow"
      if (Test-Path $CSV) {
        $ProcessedUsers = Import-CSV $CSV | Select-Object -ExpandProperty PrimarySmtpAddress
      } else {
        $ProcessedUsers = @()
      }
    }
    
    $Mailboxes | ForEach-Object {
      $ProcessedCount++
      $PercentComplete = (($ProcessedCount / $TotalMailboxes) * 100)
      $Mbox = $_
      $Size = $null
      $ArchiveSize = $null
      Write-Progress -Activity "Processing $Mbox" -Status "$ProcessedCount out of $TotalMailboxes completed ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      # Skip mailboxes already processed
      if ($ProcessedUsers -contains $Mbox.PrimarySmtpAddress) {
        Write-Host "Skipping $($Mbox.PrimarySmtpAddress), already processed."
        continue
      }

      # Retry logic for Get-MailboxStatistics
      $maxAttempts = 3
      $attempt = 0
      $MailboxStats = $null
      do {
        try {
          $MailboxStats = Get-MailboxStatistics $Mbox.UserPrincipalName -ErrorAction Stop
          $success = $true
        } catch {
          $attempt++
          Write-Warning "Error retrieving mailbox statistics for $($Mbox.UserPrincipalName), retry $attempt of $maxAttempts"
          Start-Sleep -Seconds 5
          $success = $false
        }
      } while (-not $success -and $attempt -lt $maxAttempts)

      if (-not $success) {
        Write-Error "Failed to retrieve mailbox statistics for $($Mbox.UserPrincipalName) after $maxAttempts attempts."
        $MailboxSize = "Error"
      } else {
        $MailboxSize = if ($MailboxStats.TotalItemSize -ne $null) { 
          [math]::Round(($MailboxStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)
        } else { 0 }
      }

      # Archive Mailbox Size
      if ($Mbox.ArchiveDatabase -ne $null) {
        try {
          $MailboxArchiveSize = Get-MailboxStatistics $Mbox.UserPrincipalName -Archive -ErrorAction Stop
          $ArchiveSize = if ($MailboxArchiveSize.TotalItemSize -ne $null) {
            [math]::Round(($MailboxArchiveSize.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
          } else { 0 }
        } catch {
          Write-Warning "Error retrieving archive mailbox size for $($Mbox.UserPrincipalName): $_"
          $ArchiveSize = "Error"
        }
      }

      # Save mailbox data
      $arr_MbxStats += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserName = $Mbox.DisplayName
        ServerName = $Mbox.ServerName
        Database = $Mbox.Database
        RecipientTypeDetails = $Mbox.RecipientTypeDetails
        PrimarySmtpAddress = $Mbox.PrimarySmtpAddress
        "Mailbox Size (GB)" = $MailboxSize
        "Issue Warning Quota (GB)" = if ( $Round ) { [Math]::Ceiling($Mbox.IssueWarningQuota -Replace " GB.*") } else { $Mbox.IssueWarningQuota -Replace " GB.*" }
        "Prohibit Send Quota (GB)" = if ( $Round ) { [Math]::Ceiling($Mbox.ProhibitSendQuota -Replace " GB.*") } else { $Mbox.ProhibitSendQuota -Replace " GB.*" }
        "Archive Database" = if ($Mbox.ArchiveDatabase -ne $null) { $Mbox.ArchiveDatabase } else { $null }
        "Archive Name" = if ($Mbox.ArchiveDatabase -ne $null) { $Mbox.ArchiveName } else { $null }
        "Archive State" = if ($Mbox.ArchiveDatabase -ne $null) { $Mbox.ArchiveState } else { $null }
        "Archive MailboxSize (GB)" = $ArchiveSize
        "Archive Warning Quota (GB)" = if ($Mbox.ArchiveDatabase -ne $null) { if ($Round) { [Math]::Ceiling($Mbox.ArchiveWarningQuota -Replace " GB.*") } else { $Mbox.ArchiveWarningQuota -Replace " GB.*" } } else { $null }
        "Archive Quota (GB)" = if ($Mbox.ArchiveDatabase -ne $null) { if ($Round) { [Math]::Ceiling($Mbox.ArchiveQuota -Replace " GB.*") } else { $Mbox.ArchiveQuota -Replace " GB.*" } } else { $null }
        AutoExpandingArchiveEnabled = $Mbox.AutoExpandingArchiveEnabled
      })

      # Save partial results every 10 mailboxes
      if ($WriteToCSV -and ($ProcessedCount % 10 -eq 0)) {
        Write-Host "Processed $ProcessedCount out of $TotalMailboxes mailboxes, saving partial results to CSV ..."
        $arr_MbxStats | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Append
      }
    }

    # Save final results
    if ($WriteToCSV) {
      $arr_MbxStats | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_MbxStats | Out-Host
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
}

function Export-MsolAccountSku {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV
  )
  
  priv_SetPreferences -Verbose
  $mggConnectedCheck = priv_CheckMGGraphModule
  $folder = priv_CheckFolder($folderCSV)

  if ( $mggConnectedCheck -eq $true ) {
    $arr_MsolAccountSku = @()
    $ProcessedCount = 0
    $licenseFileURL = "https://raw.githubusercontent.com/$($GTBVars.RepoOwner)/$($GTBVars.RepoName)/main/$($GTBVars.LicenseFilePath)"
    
    # Check GitHub for last commit date
    $apiUrl = "https://api.github.com/repos/$($GTBVars.RepoOwner)/$($GTBVars.RepoName)/commits?path=$($GTBVars.LicenseFilePath)"
    $maxAttempts = 3
    $attempt = 0
    do {
      try {
        $response = Invoke-RestMethod -Uri $apiUrl -Headers @{'User-Agent' = 'Gioxx.ToyBox'} -ErrorAction Stop
        $lastCommitDate = $response[0].commit.committer.date
        $utcDateTime = [DateTime]::ParseExact($lastCommitDate, "MM/dd/yyyy HH:mm:ss", $null)
        Write-InformationColored "License file: $($licenseFileURL)`nLast license file update: $($utcDateTime.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss"))" -ForegroundColor "Cyan"
        $success = $true
      } catch {
        Write-InformationColored "Failed to retrieve last commit date, attempt $($attempt) of $($maxAttempts)." -ForegroundColor "Red"
        Start-Sleep -Seconds 5
        $success = $false
      }
      $attempt++
    } while (-not $success -and $attempt -lt $maxAttempts)

    # Download license file from GitHub (https://raw.githubusercontent.com/gioxx/ps.toybox/main/JSON/M365_licenses.json)
    $attempt = 0
    do {
      try {
        $licenseFile = Invoke-RestMethod -Method Get -Uri $licenseFileURL -ErrorAction Stop
        Write-InformationColored "License file downloaded correctly." -ForegroundColor "Green"
        $success = $true
      } catch {
        Write-InformationColored "Failure downloading license file, attempt $attempt of $maxAttempts" -ForegroundColor "Red"
        Start-Sleep -Seconds 5
        $success = $false
      }
      $attempt++
    } while (-not $success -and $attempt -lt $maxAttempts)

    if (-not $success) {
      Write-Error "Downloading license file failed after $maxAttempts attempts. Aborting."
      exit 1
    }

    $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format 'yyyyMMdd').ToString())_M365-User-License-Report.csv")
    if (Test-Path $CSV) {
      $ProcessedUsers = Import-CSV $CSV | Select-Object -ExpandProperty UserPrincipalName
    } else {
      $ProcessedUsers = @()
    }

    try {
      $Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All -ErrorAction Stop
    } catch {
      Write-Error "Failed to retrieve users with assigned licenses: $_"
      exit 1
    }

    $Users | ForEach {
      $ProcessedCount++
      $User = $_
      $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
      Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      if ($ProcessedUsers -contains $User.UserPrincipalName) {
        Write-Host "Skipping $($User.UserPrincipalName), already processed."
        continue
      }
      $attempt = 0
      $GraphLicense = $null
      do {
        try {
          $GraphLicense = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
          $success = $true
        } catch {
          Write-InformationColored "Failed to retrieve licenses for $($User.UserPrincipalName), attempt $attempt of $maxAttempts" -ForegroundColor "Red"
          Start-Sleep -Seconds 5
          $success = $false
        }
        $attempt++
      } while (-not $success -and $attempt -lt $maxAttempts)

      if (-not $success) {
        Write-Error "Failed to retrieve licenses for $($User.UserPrincipalName) after $maxAttempts attempts. Skipping."
        continue
      }
      if ($GraphLicense -ne $null) {
        ForEach ( $License in $($GraphLicense.SkuPartNumber) ) {
          ForEach ( $LicenseStringId in $licenseFile ) {
            if ( $LicenseStringId.String_Id -eq $License ) {
              $arr_MsolAccountSku += New-Object -TypeName PSObject -Property $([ordered]@{
                DisplayName = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                PrimarySmtpAddress = $User.Mail
                Licenses = $LicenseStringId.Product_Display_Name
              })
              break
            }
          }
        }
      }

      # Save partial results every 50 users
      if ($ProcessedCount % 50 -eq 0) {
        Write-InformationColored "Processed $ProcessedCount out of $totalUsers, saving partial results ..." -ForegroundColor "Yellow"
        $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Append
      }
    }

    # Save final results
    $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

  } else {
    Write-Error "`nCan't connect or use Microsoft Graph modules. `nPlease check logs."
  }
  
  priv_RestorePreferences
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Export-MboxStatistics"
Export-ModuleMember -Function "Export-MsolAccountSku"
