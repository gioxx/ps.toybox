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

function AddMboxPermission {
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

function MboxPermission {
  param(
    [Parameter(Mandatory)][string] $SourceMailbox
  )
  Get-MailboxPermission -Identity $SourceMailbox | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select-Object Identity,User,AccessRights
  Get-RecipientPermission $SourceMailbox -AccessRights SendAs | Where-Object {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select-Object Identity,Trustee,AccessRights | Out-String
  Get-Mailbox $SourceMailbox | Select-Object -Expand GrantSendOnBehalfTo
}

function RemoveMboxPermission {
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

Export-ModuleMember -Function AddMboxPermission
Export-ModuleMember -Function ConnectMSOnline
Export-ModuleMember -Function ExplodeDDG
Export-ModuleMember -Function MboxPermission
Export-ModuleMember -Function QuarantineRelease
Export-ModuleMember -Function ReloadModule
Export-ModuleMember -Function RemoveMboxPermission
Export-ModuleMember -Function SmtpExpand
