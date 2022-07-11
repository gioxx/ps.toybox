# M365: connessioni ======================================================================================================================================================================

function ConnectMSOnline {
  Import-Module MSOnline -UseWindowsPowershell
  Connect-MsolService
}

# Check ACL caselle di posta =============================================================================================================================================================

function MboxPermission {
  param( [string] $sourceMailbox )
  Get-MailboxPermission -Identity $sourceMailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights
  Get-RecipientPermission $sourceMailbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select Identity,Trustee,AccessRights | Out-String
  Get-Mailbox $sourceMailbox | Select -Expand GrantSendOnBehalfTo
}

# Modifica ACL caselle di posta ==========================================================================================================================================================

function AddMboxPermission {
  param( [string] $sourceMailbox )
  Write-Host "Add $($_) on $($sourceMailbox) ..."
  Add-MailboxPermission -Identity $sourceMailbox -User $_ -AccessRights FullAccess -Confirm:$false
  Add-RecipientPermission $sourceMailbox -Trustee $_ -AccessRights SendAs -Confirm:$false
}

function RemoveMboxPermission {
  param( [string] $sourceMailbox )
  Write-Host "Remove $($_) from $($sourceMailbox) ..."
  Remove-MailboxPermission -Identity $sourceMailbox -User $_ -AccessRights FullAccess -Confirm:$false
  Remove-RecipientPermission $sourceMailbox -Trustee $_ -AccessRights SendAs -Confirm:$false
}

# M365: Protection =======================================================================================================================================================================

function QuarantineRelease {
  param(
    [string] $senderAddress,
    [switch] $release
  )
  if ($release) {
    Write-Host "Release quarantine from known senders: release email(s) from $($senderAddress) ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $senderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ? {$_.QuarantinedUser -ne $null} | Release-QuarantineMessage -ReleaseToAll
    Write-Host "Release quarantine from known senders: verifying email(s) from $($senderAddress) just released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $senderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  } else {
    Write-Host "Find email(s) from known senders quarantined: email(s) from $($senderAddress) not yet released ..."
    Get-QuarantineMessage -QuarantineTypes TransportRule -SenderAddress $senderAddress | ForEach {Get-QuarantineMessage -Identity $_.Identity} | ft -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
  }
}

Export-ModuleMember -Function AddMboxPermission
Export-ModuleMember -Function ConnectMSOnline
Export-ModuleMember -Function MboxPermission
Export-ModuleMember -Function QuarantineRelease
Export-ModuleMember -Function RemoveMboxPermission
