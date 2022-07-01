# Check ACL caselle di posta =============================================================================================================================================================

function MboxPermission {
  param( $sourceMailbox )
  Get-MailboxPermission -Identity $sourceMailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -NotLike "S-1-5*" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights
  Get-RecipientPermission $sourceMailbox -AccessRights SendAs | where {$_.Trustee.tostring() -ne "NT AUTHORITY\SELF" -and $_.Trustee.tostring() -NotLike "S-1-5*"} | Select Identity,Trustee,AccessRights | Out-String
  Get-Mailbox $sourceMailbox | Select -Expand GrantSendOnBehalfTo
}

# Modifica ACL caselle di posta ==========================================================================================================================================================

function AddMboxPermission {
  param( $sourceMailbox )
  Write-Host "Add $($_) permission on $($sourceMailbox) ..."
  Add-MailboxPermission -Identity $sourceMailbox -User $_ -AccessRights FullAccess -Confirm:$false
  Add-RecipientPermission $sourceMailbox -Trustee $_ -AccessRights SendAs -Confirm:$false
}

function RemoveMboxPermission {
  param( $sourceMailbox )
  Write-Host "Remove $($_) permission on $($sourceMailbox) ..."
  Remove-MailboxPermission -Identity $sourceMailbox -User $_ -AccessRights FullAccess -Confirm:$false
  Remove-RecipientPermission $sourceMailbox -Trustee $_ -AccessRights SendAs -Confirm:$false
}

Export-ModuleMember -Function MboxPermission
Export-ModuleMember -Function AddMboxPermission
Export-ModuleMember -Function RemoveMboxPermission
