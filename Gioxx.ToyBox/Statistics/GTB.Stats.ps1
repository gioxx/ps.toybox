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
    
    $Mailboxes | ForEach {
      $ProcessedCount++
      $PercentComplete = (($ProcessedCount / $TotalMailboxes) * 100)
      $Mbox = $_
      $Size = $null
      $ArchiveSize = $null
      Write-Progress -Activity "Processing $Mbox" -Status "$ProcessedCount out of $TotalMailboxes completed ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
      
      if ( $Mbox.ArchiveDatabase -ne $null) {
        $MailboxArchiveSize = Get-MailboxStatistics $Mbox.UserPrincipalName -Archive
        if ( $MailboxArchiveSize.TotalItemSize -ne $null ) {
          $ArchiveSize = [math]::Round(($MailboxArchiveSize.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
        } else {
          $ArchiveSize = 0
        }
      }

      $MailboxSize = [math]::Round((((Get-MailboxStatistics $Mbox.UserPrincipalName).TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)

      $arr_MbxStats += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserName = $Mbox.DisplayName
        ServerName = $Mbox.ServerName
        Database = $Mbox.Database
        RecipientTypeDetails = $Mbox.RecipientTypeDetails
        PrimarySmtpAddress = $Mbox.PrimarySmtpAddress
        "Mailbox Size (GB)" = $MailboxSize
        "Issue Warning Quota (GB)" = if ( $Round ) { [Math]::Ceiling($Mbox.IssueWarningQuota -Replace " GB.*") } else { $Mbox.IssueWarningQuota -Replace " GB.*" }
        "Prohibit Send Quota (GB)" = if ( $Round ) { [Math]::Ceiling($Mbox.ProhibitSendQuota -Replace " GB.*") } else { $Mbox.ProhibitSendQuota -Replace " GB.*" }
        "Archive Database" = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveDatabase } else { $null }
        "Archive Name" = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveName } else { $null }
        "Archive State" = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveState } else { $null }
        "Archive MailboxSize (GB)" = if ( $Mbox.ArchiveDatabase -ne $null ) { $ArchiveSize } else { $null }
        "Archive Warning Quota (GB)" = if ( $Mbox.ArchiveDatabase -ne $null ) { if ( $Round ) { [Math]::Ceiling($Mbox.ArchiveWarningQuota -Replace " GB.*") } else { $Mbox.ArchiveWarningQuota -Replace " GB.*" } } else { $null }
        "Archive Quota (GB)" = if ( $Mbox.ArchiveDatabase -ne $null ) { if ( $Round ) { [Math]::Ceiling($Mbox.ArchiveQuota -Replace " GB.*") } else { $Mbox.ArchiveQuota -Replace " GB.*" } } else { $null }
        AutoExpandingArchiveEnabled = $Mbox.AutoExpandingArchiveEnabled
        })
    }

    if ( $WriteToCSV ) {
      $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MailboxStatistics.csv")
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
    $response = Invoke-RestMethod -Uri $apiUrl -Headers @{'User-Agent' = 'Gioxx.ToyBox'}
    $lastCommitDate = $response[0].commit.committer.date
    $utcDateTime = [DateTime]::ParseExact($lastCommitDate, "MM/dd/yyyy HH:mm:ss", $null)
    "License file: $($licenseFileURL)`nLast license file update: $($utcDateTime.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss"))" | Out-Host

    # Download license file from GitHub (https://raw.githubusercontent.com/gioxx/ps.toybox/main/JSON/M365_licenses.json)
    try {
      $licenseFile = Invoke-RestMethod -Method Get -Uri $licenseFileURL
      "License file downloaded correctly." | Out-Host
    }
    catch {
      "License file retrieval error: $_" | Out-Host
      exit 1
    }

    $Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All

    $Users | ForEach {
      $ProcessedCount++
      $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
      $User = $_
      Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
      $GraphLicense = Get-MgUserLicenseDetail -UserId $User.Id
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
    }
    
    $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-User-License-Report.csv")
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
