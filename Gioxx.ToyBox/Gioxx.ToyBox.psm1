# Connections ======================================================================================================================================================

function ConnectMSOnline {
  Import-Module MSOnline -UseWindowsPowershell
  Connect-MsolService | Out-Null
}

# Groups ===========================================================================================================================================================

function ExplodeDDG {
  param(
    [Parameter(Mandatory)][string] $DDG,
    [switch] $GridView
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
    [switch] $Remove
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
    [Parameter(Mandatory)][string] $SourceMailbox
  )
  Get-MailboxPermission -Identity $SourceMailbox | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select-Object Identity,User,AccessRights
  Get-RecipientPermission $SourceMailbox -AccessRights SendAs | Where-Object {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select-Object Identity,Trustee,AccessRights | Out-String
  Get-Mailbox $SourceMailbox | Select-Object -Expand GrantSendOnBehalfTo
}

function MboxPermission-Add {
  param(
    [Parameter(Mandatory)][string] $SourceMailbox,
    [Parameter(Mandatory)][string] $UserMailbox,
    [Parameter(Mandatory)][string] $AccessRights,
    [switch] $AutoMapping
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
    [Parameter(Mandatory)][string] $SourceMailbox,
    [Parameter(Mandatory)][string] $UserMailbox,
    [Parameter(Mandatory)][string] $AccessRights
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
    [Parameter(Mandatory, HelpMessage="Mailbox Display Name (example: Contoso srl -Info)")][string] $SharedMailboxDisplayName,
    [Parameter(Mandatory, HelpMessage="Mailbox Alias (example: Contososrl_info)")][string] $SharedMailboxAlias
  )
  New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySMTPAddress $SharedMailboxSMTPAddress
  Write-Host "Set outgoing email copy save for $($SharedMailboxSMTPAddress)" -f "Yellow"
  Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSentAsEnabled $True
	Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSendOnBehalfEnabled $True
  Write-Host "All done, remember to set access and editing rights to the new mailbox."
}

function SmtpExpand {
  param(
    [Parameter(Mandatory)][string] $SourceMailbox
  )
  Get-Recipient $SourceMailbox | Select-Object Name -Expand EmailAddresses | Where-Object {$_ -like 'smtp*'}
}

# Modules ==========================================================================================================================================================

function ReloadModule {
  param(
    [Parameter(Mandatory)][string] $Module
  )
  Write-Host "Reload $($Module) module ..."
  Remove-Module $Module
  Import-Module $Module
  Get-Module | Where-Object { $_.Name -eq "$($Module)" }
}

# Protection =======================================================================================================================================================

function QuarantineRelease {
  param(
    [string] $SenderAddress,
    [switch] $Release
  )
  if ($Release) {
    Write-Host "Release quarantine from known senders: release email(s) from $($SenderAddress) ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Where-Object {$null -ne $_.QuarantinedUser} | Release-QuarantineMessage -ReleaseToAll
    Write-Host "Release quarantine from known senders: verifying email(s) from $($SenderAddress) just released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  } else {
    Write-Host "Find email(s) from known senders quarantined: email(s) from $($SenderAddress) not yet released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $SenderAddress | ForEach-Object {Get-QuarantineMessage -Identity $_.Identity} | Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  }
}

# Start your engine ================================================================================================================================================

Export-ModuleMember -Function ConnectMSOnline
Export-ModuleMember -Function ExplodeDDG
Export-ModuleMember -Function MboxAlias
Export-ModuleMember -Function MboxPermission
Export-ModuleMember -Function MboxPermission-Add
Export-ModuleMember -Function MboxPermission-Remove
Export-ModuleMember -Function QuarantineRelease
Export-ModuleMember -Function ReloadModule
Export-ModuleMember -Function SharedMbox-New
Export-ModuleMember -Function SmtpExpand
