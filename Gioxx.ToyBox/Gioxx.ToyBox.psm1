# PSM1 Functions ===================================================================================================================================================

function SaveFileWithProgressiveNumber($path)
{
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $extension = [System.IO.Path]::GetExtension($path)
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $count = 1
    while (Test-Path $path)
    {
        $fileName = $baseName + "_$count" + $extension
        $path = Join-Path -Path $directory -ChildPath $fileName
        $count++
    }
    return $path
}

# Connections ======================================================================================================================================================

function ConnectEOL {
  param(
    [Parameter(Mandatory, HelpMessage="User to connect to Exchange Online with")][string] $UserPrincipalName
  )
  if ( (Get-Module -Name ExchangeOnlineManagement -ListAvailable).count -eq 0 ) {
    Write-Host "Install the ExchangeOnlineManagement module using this command (then relaunch this script): `nInstall-Module ExchangeOnlineManagement" -f "Yellow"
  } else {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
  }
}

function ConnectMSOnline {
  if ( (Get-Module -Name MSOnline -ListAvailable).count -eq 0 ) {
    Write-Host "Install the MSOnline module using this command (then relaunch this script): `nInstall-Module MSOnline" -f "Yellow"
  } else {
    Import-Module MSOnline -UseWindowsPowershell
    Connect-MsolService | Out-Null
  }
}

# Groups ===========================================================================================================================================================

function ExplodeDDG {
  param(
    [Parameter(Mandatory, HelpMessage="Dynamic Distribution Group e-mail address or display name")][string] $DDG,
    [Parameter(Mandatory=$false, HelpMessage="Show results in a grid view")][switch] $GridView
  )
  if ($GridView) {
    Write-Host "List $($DDG) members using GridView ..."
    Get-DynamicDistributionGroupMember $DDG | Select-Object DisplayName,FirstName,LastName,PrimarySmtpAddress,Company,City | Out-GridView
  } else {
    Write-Host "List $($DDG) members ..."
    Get-DynamicDistributionGroupMember $DDG | Select-Object DisplayName,FirstName,LastName,PrimarySmtpAddress,Company,City
  }
}

# Mailboxes ========================================================================================================================================================

function MboxAlias {
  param(
    [Parameter(Mandatory, HelpMessage="User to edit (e.g. mario.rossi)")][string] $SourceMailbox,
    [Parameter(Mandatory, HelpMessage="Alias to be added (e.g. mario.rossi.alias@contoso.com)")][string] $MailboxAlias,
    [Parameter(Mandatory=$false, HelpMessage="Remove alias from mailbox")][switch] $Remove
  )
  if ($Remove) {
    Set-Mailbox $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"}
    Get-Recipient $SourceMailbox | Select Name -Expand EmailAddresses | Where-Object {$_ -like 'smtp*'}
  } else {
    Set-Mailbox $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"}
    Get-Recipient $SourceMailbox | Select Name -Expand EmailAddresses | Where-Object {$_ -like 'smtp*'}
  }
}

function MboxPermission {
  param(
    [Parameter(Mandatory, HelpMessage="Mailbox e-mail address or display name (e.g. mario.rossi@contoso.com)")][string] $SourceMailbox
  )
  Get-MailboxPermission -Identity $SourceMailbox | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select-Object Identity,User,AccessRights | ft -Wrap
  Get-RecipientPermission $SourceMailbox -AccessRights SendAs | Where-Object {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select-Object Identity,Trustee,AccessRights | Out-String
  Get-Mailbox $SourceMailbox | Select-Object -Expand GrantSendOnBehalfTo
}

function MboxPermission-Add {
  param(
    [Parameter(Mandatory, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")][string] $SourceMailbox,
    [Parameter(Mandatory, HelpMessage="E-mail address of the mailbox to which to allow access (e.g. mario.rossi@contoso.com)")][string] $UserMailbox,
    [Parameter(Mandatory, HelpMessage="Type of access to be allowed (All, FullAccess, SendAs)")][string] $AccessRights,
    [Parameter(Mandatory=$false, HelpMessage="Set mailbox automapping")][switch] $AutoMapping
  )
  Switch ($AccessRights) {
    "FullAccess" {
      if ($AutoMapping) {
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) ..."
        Add-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -AutoMapping:$true -Confirm:$false
      } else {
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) without AutoMapping ..."
        Add-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
      }
    }
    "SendAs" {
      Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) ..."
      Add-RecipientPermission $SourceMailbox -Trustee $UserMailbox -AccessRights SendAs -Confirm:$false
    }
    "All" {
      if ($AutoMapping) {
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) ..."
        Add-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -AutoMapping:$true -Confirm:$false
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) ..."
        Add-RecipientPermission $SourceMailbox -Trustee $UserMailbox -AccessRights SendAs -Confirm:$false
      }
      else {
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) without AutoMapping ..."
        Add-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
        Write-Host "Add $($UserMailbox) ($($AccessRights)) on $($SourceMailbox) ..."
        Add-RecipientPermission $SourceMailbox -Trustee $UserMailbox -AccessRights SendAs -Confirm:$false
      }
    }
  }
}

function MboxPermission-Remove {
  param(
    [Parameter(Mandatory, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")][string] $SourceMailbox,
    [Parameter(Mandatory, HelpMessage="E-mail address of the mailbox to which to remove access (e.g. mario.rossi@contoso.com)")][string] $UserMailbox,
    [Parameter(Mandatory, HelpMessage="Type of access to be removed (All, FullAccess, SendAs)")][string] $AccessRights
  )
  Write-Host "Remove $($UserMailbox) ($($AccessRights)) from $($SourceMailbox)..."
  Switch ($AccessRights) {
    "FullAccess" { Remove-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -Confirm:$false }
    "SendAs" { Remove-RecipientPermission $SourceMailbox -Trustee $UserMailbox -AccessRights SendAs -Confirm:$false }
    "All" {
      Remove-MailboxPermission -Identity $SourceMailbox -User $UserMailbox -AccessRights FullAccess -Confirm:$false
      Remove-RecipientPermission $SourceMailbox -Trustee $UserMailbox -AccessRights SendAs -Confirm:$false
    }
  }
}

function SharedMbox-New {
  param(
    [Parameter(Mandatory, HelpMessage="Primary SMTP Address (example: info@contoso.com)")][string] $SharedMailboxSMTPAddress,
    [Parameter(Mandatory, HelpMessage="Mailbox Display Name (example: Contoso srl - Info)")][string] $SharedMailboxDisplayName,
    [Parameter(Mandatory, HelpMessage="Mailbox Alias (example: Contososrl_info)")][string] $SharedMailboxAlias
  )
  New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySMTPAddress $SharedMailboxSMTPAddress
  Write-Host "Set outgoing e-mail copy save for $($SharedMailboxSMTPAddress)" -f "Yellow"
  Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSentAsEnabled $True
	Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSendOnBehalfEnabled $True
  Write-Host "All done, remember to set access and editing rights to the new mailbox."
}

function SmtpExpand {
  param(
    [Parameter(Mandatory, HelpMessage="E-mail address of the mailbox to be analyzed (e.g. info@contoso.com)")][string] $SourceMailbox
  )
  Get-Recipient $SourceMailbox | Select-Object Name -Expand EmailAddresses | Where-Object {$_ -like 'smtp*'}
}

# Modules ==========================================================================================================================================================

function ReloadModule {
  param(
    [Parameter(Mandatory, HelpMessage="Name of the module to reload (e.g. Gioxx.ToyBox)")][string] $Module
  )
  Write-Host "Reload $($Module) module ..."
  Import-Module $Module -Force
  Get-Module | Where-Object { $_.Name -eq "$($Module)" }
}

# Protection =======================================================================================================================================================

function QuarantineRelease {
  param(
    [Parameter(Mandatory, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")][string] $SenderAddress,
    [Parameter(Mandatory=$false, HelpMessage="Unlock emails stuck in quarantine")][switch] $Release
  )
  if ($Release) {
    Write-Host "Release quarantine from known senders: release e-mail(s) from $($SenderAddress) ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Where-Object {$null -ne $_.QuarantinedUser} | Release-QuarantineMessage -ReleaseToAll
    Write-Host "Release quarantine from known senders: verifying e-mail(s) from $($SenderAddress) just released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  } else {
    Write-Host "Find e-mail(s) from known senders quarantined: e-mail(s) from $($SenderAddress) not yet released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  }
}

# Statistics =======================================================================================================================================================

function MboxStatistics-Export {
  param(
    [Parameter(Mandatory=$false, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")][string] $folderCSV,
    [Parameter(Mandatory=$false, HelpMessage="Round up the values of ArchiveWarningQuotaInGB and ArchiveQuotaInGB (by excess).")][switch] $Round
  )
  Set-Variable ProgressPreference Continue
  if ([string]::IsNullOrEmpty($folderCSV)) {
    $folderCSV = "C:\Temp"
  } else {
    $folderCSV = $folderCSV.TrimEnd('\')
  }
  $Today = Get-Date -format yyyyMMdd

  $Result=@()
  $ProcessedCount = 0
  $Mailboxes = Get-Mailbox -ResultSize Unlimited
  $TotalMailboxes = $Mailboxes.Count
  
  $Mailboxes | Foreach-Object {
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

    $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
     UserName = $Mbox.DisplayName
     ServerName = $Mbox.ServerName
     Database = $Mbox.Database
     RecipientTypeDetails = $Mbox.RecipientTypeDetails
     PrimarySmtpAddress = $Mbox.PrimarySmtpAddress
     MailboxSizeInGB = $MailboxSize
     IssueWarningQuotaInGB = if ( $Round ) { [Math]::Ceiling($Mbox.IssueWarningQuota -Replace " GB.*") } else { $Mbox.IssueWarningQuota -Replace " GB.*" }
     ProhibitSendQuotaInGB = if ( $Round ) { [Math]::Ceiling($Mbox.ProhibitSendQuota -Replace " GB.*") } else { $Mbox.ProhibitSendQuota -Replace " GB.*" }
     ArchiveDatabase = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveDatabase } else { $null }
     ArchiveName = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveName } else { $null }
     ArchiveState = if ( $Mbox.ArchiveDatabase -ne $null ) { $Mbox.ArchiveState } else { $null }
     ArchiveMailboxSizeInGB = if ( $Mbox.ArchiveDatabase -ne $null ) { $ArchiveSize } else { $null }
     ArchiveWarningQuotaInGB = if ( $Mbox.ArchiveDatabase -ne $null ) { if ( $Round ) { [Math]::Ceiling($Mbox.ArchiveWarningQuota -Replace " GB.*") } else { $Mbox.ArchiveWarningQuota -Replace " GB.*" } } else { $null }
     ArchiveQuotaInGB = if ( $Mbox.ArchiveDatabase -ne $null ) { if ( $Round ) { [Math]::Ceiling($Mbox.ArchiveQuota -Replace " GB.*") } else { $Mbox.ArchiveQuota -Replace " GB.*" } } else { $null }
     AutoExpandingArchiveEnabled = $Mbox.AutoExpandingArchiveEnabled
    })
  }
  $CSV = SaveFileWithProgressiveNumber("$($folderCSV)\$($Today)_MailboxSize.csv")
  $Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

function MsolAccountSku-Export {
  param(
    [Parameter(Mandatory=$false, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")][string] $folderCSV
  )

  if ( (Get-Module -Name Microsoft.Graph -ListAvailable).count -eq 0 ) {
    Write-Host "Please install the Graph module using this command (then relaunch this script): `nInstall-Module Microsoft.Graph" -f "Yellow"
    exit
  } else { Connect-MgGraph | Out-Null }

  if ( (Get-Module -Name Microsoft.Graph.Users -ListAvailable).count -eq 0 ) {
    Write-Host "Please install the Microsoft.Graph.Users module using this command (then relaunch this script): `nInstall-Module Microsoft.Graph.Users" -f "Yellow"
    exit
  } else {
    if ( (Get-Module -Name Microsoft.Graph.Users).count -eq 0 ) {
      Import-Module Microsoft.Graph.Users
    }
  }

  Set-Variable ProgressPreference Continue
  if ([string]::IsNullOrEmpty($folderCSV)) {
    $folderCSV = "C:\Temp"
  } else {
    $folderCSV = $folderCSV.TrimEnd('\')
  }
  $Today = Get-Date -format yyyyMMdd
  
  $Result=@()
  $ProcessedCount = 0
  $licenseFile = Invoke-RestMethod -Method Get -Uri 'https://raw.githubusercontent.com/gioxx/ps.toybox/main/JSON/M365_licenses.json'
  $Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All

  $Users | Foreach-Object {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
    $User = $_
    Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    $GraphLicense = Get-MgUserLicenseDetail -UserId $User.Id
    if ($GraphLicense -ne $null) {
      ForEach ( $License in $($GraphLicense.SkuPartNumber) ) {
        ForEach ( $licName in $licenseFile ) {
          if ( $licName.licName -eq $License ) {
              $Result += New-Object -TypeName PSObject -Property $([ordered]@{
              DisplayName = $User.DisplayName
              UserPrincipalName = $User.UserPrincipalName
              PrimarySmtpAddress = $User.Mail
              Licenses = $licName.licDisplayName
            })
            break
          }
        }
      }
    }
  }
  $CSV = SaveFileWithProgressiveNumber("$($folderCSV)\O365-User-License-Report_$($Today).csv")
  $Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

# Start your engine ================================================================================================================================================

Export-ModuleMember -Function ConnectEOL
Export-ModuleMember -Function ConnectMSOnline
Export-ModuleMember -Function ExplodeDDG
Export-ModuleMember -Function MboxAlias
Export-ModuleMember -Function MboxPermission
Export-ModuleMember -Function MboxPermission-Add
Export-ModuleMember -Function MboxPermission-Remove
Export-ModuleMember -Function MboxStatistics-Export
Export-ModuleMember -Function MsolAccountSku-Export
Export-ModuleMember -Function QuarantineRelease
Export-ModuleMember -Function ReloadModule
Export-ModuleMember -Function SharedMbox-New
Export-ModuleMember -Function SmtpExpand
