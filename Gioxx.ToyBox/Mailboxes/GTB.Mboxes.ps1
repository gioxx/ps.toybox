# Mailboxes ========================================================================================================================================================

function Add-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="User to edit (e.g. mario.rossi)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Alias to be added (e.g. mario.rossi.alias@contoso.com)")]
    [string] $MailboxAlias
  )
 
  Set-Mailbox $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"}
  Get-Recipient $SourceMailbox | 
      Select-Object Name -Expand EmailAddresses | 
      Where {$_ -like 'smtp*'}
}

function Add-MboxPermission {
  param(
    [Parameter(Mandatory=$True, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to which to allow access (e.g. mario.rossi@contoso.com)")]
    [string[]] $UserMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Type of access to be allowed (All, FullAccess, SendAs)")]
    [string] $AccessRights,
    [Parameter(Mandatory=$False, HelpMessage="Set mailbox automapping")]
    [switch] $AutoMapping
  )

  begin {
    if ([string]::IsNullOrEmpty($AccessRights)) { $AccessRights = "All" }
  }

  process {
    $UserMailbox | ForEach {
      $CurrentUser = $_
      Switch ($AccessRights) {
        "FullAccess" {
          if ($AutoMapping) {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False
          } else {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) without AutoMapping ..."
            Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False
          }
        }
        "SendAs" {
          Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
          Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
        }
        "All" {
          if ($AutoMapping) {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
          }
          else {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) without AutoMapping ..."
            Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
          }
        }
      }
    }
  }
}

function Change-MboxLanguage {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox whose language is to be changed (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Language selected (e.g. it)")]
    [string] $Language,
    [Parameter(Mandatory=$False, HelpMessage="CSV file containing the addresses of the mailboxes to which the language is to be changed - header 'EmailAddress' (e.g. C:\temp\mailboxes.csv)")]
    [string] $CSV
  )
  
  Set-Variable ProgressPreference Continue
  
  if ( [string]::IsNullOrEmpty($Language) ) { $Language = "it-IT" }

  if ( [string]::IsNullOrEmpty($CSV) ) {
    if ( -not([string]::IsNullOrEmpty($SourceMailbox)) ) {
      Write-Progress -Activity "Changing $($SourceMailbox) language to $($Language) ..."
      Set-MailboxRegionalConfiguration $SourceMailbox -LocalizeDefaultFolderName:$True -Language $Language
      Get-MailboxRegionalConfiguration $SourceMailbox
    }
  } else {
    Import-CSV $CSV | ForEach {
      Write-Progress -Activity "Changing $($_.EmailAddress) language to $($Language) ..."
      Set-MailboxRegionalConfiguration $_.EmailAddress -LocalizeDefaultFolderName:$True -Language $Language
      Get-MailboxRegionalConfiguration $_.EmailAddress
    }
  }
}

function Export-MboxAlias {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to be analyzed (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
    [switch] $CSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all mailboxes aliases")]
    [switch] $All,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all mailboxes aliases with a specified domain (e.g. contoso.com)")]
    [string] $Domain
  )

  begin {
    $mboxCounter = 0
    $Result = @()
    Set-Variable ProgressPreference Continue

    if (([string]::IsNullOrEmpty($Domain)) -And ([string]::IsNullOrEmpty($SourceMailbox))) {
      $All = $True
    }

    if (-not([string]::IsNullOrEmpty($Domain))) {
      $SourceMailbox = Get-Recipient -ResultSize Unlimited | 
          Where { $_.RecipientTypeDetails -ne "GuestMailUser" -And $_.EmailAddresses -like "*@" + $Domain }
      $CSV = $True
    }

    if ($All) {
      Write-Host "WARNING: no mailbox(es) specified, I scan all the mailboxes, please be patient." -f "Yellow"
      $SourceMailbox = Get-Recipient -ResultSize Unlimited | 
          Where { $_.RecipientTypeDetails -ne "GuestMailUser" }
      $CSV = $True
    }
    
    if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
    if ($CSV) { $folder = priv_CheckFolder($folderCSV) }
  }

  process {
    $SourceMailbox | ForEach {
      try {
        $CurrentMailbox = $_
        $mboxCounter++
        $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
        Write-Progress -Activity "Processing $((Get-Recipient $CurrentMailbox).PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        $Aliases = Get-Recipient $CurrentMailbox | 
            Select-Object -ExpandProperty EmailAddresses | 
            Where { $_ -clike "smtp*" }
        $UserPrimary = $((Get-Recipient $CurrentMailbox).PrimarySmtpAddress)

        if ($CSV) {
          $Aliases | ForEach {
              $Result += New-Object -TypeName PSObject -Property $([ordered]@{
                PrimarySmtpAddress = $UserPrimary
                Alias = $($_.SubString(5))
              })
          }
        } else {
          Write-Host "Primary: $($UserPrimary)" -f "Green"
          $Aliases | ForEach {
              Write-Host "Alias: $($_.SubString(5))"
          }
        }
      } catch {
        Write-Error $_.Exception.Message
      }
    }
  }

  end {
    if ($CSV) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-Alias-Report.csv")
      $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $Result
    }
  }
}

function Export-MboxPermission {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Which type of box to analyze (User/Shared/Room/All)")]
    [string] $RecipientType,
    [Parameter(Mandatory=$False, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV
  )

  $Result = @()
  $mboxCounter = 0
  Set-Variable ProgressPreference Continue

  Switch ($RecipientType) {
    "User" { $Mailboxes = Get-Mailbox -ResultSize Unlimited | Where { $_.RecipientTypeDetails -eq "UserMailbox" } }
    "Shared" { $Mailboxes = Get-Mailbox -ResultSize Unlimited | Where { $_.RecipientTypeDetails -eq "SharedMailbox" } }
    "Room" { $Mailboxes = Get-Mailbox -ResultSize Unlimited | Where { $_.RecipientTypeDetails -eq "RoomMailbox" } }
    "All" { 
      Write-Host "WARNING: no recipient type specified, I scan all the mailboxes, please be patient." -f "Yellow"
      $Mailboxes = Get-Mailbox -ResultSize Unlimited | 
          Where { $_.RecipientTypeDetails -eq "UserMailbox" -Or $_.RecipientTypeDetails -eq "SharedMailbox" -Or $_.RecipientTypeDetails -eq "RoomMailbox" }
    }
  }

  $Mailboxes | ForEach {
    $CurrentMailbox = $_
    
    $mboxCounter++
    $PercentComplete = (($mboxCounter / $Mailboxes.Count) * 100)
    Write-Progress -Activity "Processing $((Get-Mailbox $CurrentMailbox).PrimarySmtpAddress)" -Status "$mboxCounter out of $($Mailboxes.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

    $MboxPermSendAs = Get-RecipientPermission $(Get-Mailbox $CurrentMailbox).PrimarySmtpAddress -AccessRights SendAs |
        Where { $_.Trustee.ToString() -ne "NT AUTHORITY\SELF" -And $_.Trustee.ToString() -notlike "S-1-5*" } |
        ForEach { $_.Trustee.ToString() }

    $MboxPermFullAccess = Get-MailboxPermission $(Get-Mailbox $CurrentMailbox).PrimarySmtpAddress |
        Where { $_.AccessRights -eq "FullAccess" -and !$_.IsInherited } |
        ForEach { $_.User.ToString() }

    $Result += New-Object -TypeName PSObject -Property $([ordered]@{
      Mailbox = $(Get-Mailbox $CurrentMailbox).DisplayName
      "Mailbox Address" = $(Get-Mailbox $CurrentMailbox).PrimarySmtpAddress
      "Recipient Type" = $(Get-Mailbox $CurrentMailbox).RecipientTypeDetails
      FullAccess = $MboxPermFullAccess -join ", "
      SendAs = $MboxPermSendAs -join ", "
      SendOnBehalfTo = $(Get-Mailbox $CurrentMailbox).GrantSendOnBehalfTo
    })
  }
  $folder = priv_CheckFolder($folderCSV)
  $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MboxPermissions-Report.csv")
  $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

function Get-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to be analyzed (e.g. info@contoso.com)")]
    [string] $SourceMailbox
  )

  Get-Recipient $SourceMailbox | 
      Select-Object Name -Expand EmailAddresses | 
      Where { $_ -like 'smtp*' }
}

function Get-MboxPermission {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox e-mail address or display name (e.g. mario.rossi@contoso.com)")]
    [string] $SourceMailbox
  )

  $Result = @()

  $MboxPermFullAccess = Get-MailboxPermission $(Get-Mailbox $SourceMailbox).PrimarySmtpAddress |
      Where-Object { $_.AccessRights -eq "FullAccess" -and !$_.IsInherited } |
      ForEach-Object {
          $UserMailbox = $_.User.ToString()
          $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox).PrimarySmtpAddress
          $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

          #$existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $UserMailbox }
          $existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
          if ($existingUserObject) {
              $existingUserObject.AccessRights += ", FullAccess"
          } else {
              $Result += [PSCustomObject]@{
                  User = $DisplayName
                  UserMailbox = $PrimarySmtpAddress
                  AccessRights = "FullAccess"
              }
          }
      }

  $MboxPermSendAs = Get-RecipientPermission $(Get-Mailbox $SourceMailbox).PrimarySmtpAddress -AccessRights SendAs |
      Where-Object { $_.Trustee.ToString() -ne "NT AUTHORITY\SELF" -And $_.Trustee.ToString() -notlike "S-1-5*" } |
      ForEach-Object {
          $UserMailbox = $_.Trustee.ToString()
          $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox).PrimarySmtpAddress
          $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

          #$existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $UserMailbox }
          $existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
          if ($existingUserObject) {
              $existingUserObject.AccessRights += ", SendAs"
          } else {
              $Result += [PSCustomObject]@{
                  User = $DisplayName
                  UserMailbox = $PrimarySmtpAddress
                  AccessRights = "SendAs"
              }
          }
      }

  $MboxPermSendOnBehalfTo = $(Get-Mailbox $SourceMailbox).GrantSendOnBehalfTo |
      ForEach-Object {
          $UserMailbox = $_
          $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox).PrimarySmtpAddress
          $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

          #$existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $UserMailbox }
          $existingUserObject = $Result | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
          if ($existingUserObject) {
              $existingUserObject.AccessRights += ", SendOnBehalfTo"
          } else {
              $Result += [PSCustomObject]@{
                  User = $DisplayName
                  UserMailbox = $PrimarySmtpAddress
                  AccessRights = "SendOnBehalfTo"
              }
          }
      }

  Write-Host "`nAccess Rights on $((Get-Mailbox $SourceMailbox).DisplayName) ($((Get-Mailbox $SourceMailbox).PrimarySmtpAddress))" -f "Yellow"
  $Result
}

function New-SharedMailbox {
  param(
    [Parameter(Mandatory=$True, HelpMessage="Primary SMTP Address (example: info@contoso.com)")]
    [string] $SharedMailboxSMTPAddress,
    [Parameter(Mandatory=$True, HelpMessage="Mailbox Display Name (example: Contoso srl - Info)")]
    [string] $SharedMailboxDisplayName,
    [Parameter(Mandatory=$True, HelpMessage="Mailbox Alias (example: Contososrl_info)")]
    [string] $SharedMailboxAlias
  )

  New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySmtpAddress $SharedMailboxSMTPAddress
  Write-Host "Set outgoing e-mail copy save for $($SharedMailboxSMTPAddress)" -f "Yellow"
  Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSentAsEnabled $True
	Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSendOnBehalfEnabled $True
  Write-Host "All done, remember to set access and editing rights to the new mailbox."
}

function Remove-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline, HelpMessage="User to edit (e.g. mario.rossi)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline, HelpMessage="Alias to be removed (e.g. mario.rossi.alias@contoso.com)")]
    [string] $MailboxAlias
  )
  
  Set-Mailbox $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"}
  Get-Recipient $SourceMailbox | 
      Select-Object Name -Expand EmailAddresses | 
      Where {$_ -like 'smtp*'}
}

function Remove-MboxPermission {
  param(
    [Parameter(Mandatory=$True, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to which to remove access (e.g. mario.rossi@contoso.com)")]
    [string[]] $UserMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Type of access to be removed (All, FullAccess, SendAs)")]
    [string] $AccessRights
  )

  begin {
    if ([string]::IsNullOrEmpty($AccessRights)) { $AccessRights = "All" }
  }

  process {
    $UserMailbox | ForEach {
      $CurrentUser = $_
      Switch ($AccessRights) {
        "FullAccess" { Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False }
        "SendAs" { Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False }
        "All" {
          Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False
          Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
        }
      }
    }
  }
}

function Set-MboxRulesQuota {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox address to which to expand space for rules (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox
  )
  
  begin {
    $mboxCounter = 0
    $Result = @()
    Set-Variable ProgressPreference Continue
  }

  process {
    $SourceMailbox | ForEach {
      try {
        $CurrentMailbox = $_
        $GetCM = Get-Mailbox $CurrentMailbox
        
        $mboxCounter++
        $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
        Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        Set-Mailbox $CurrentMailbox -RulesQuota 256KB

        $Result += New-Object -TypeName PSObject -Property $([ordered]@{
          PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
          "Rules Quota" = $GetCM.RulesQuota
        })
      } catch {
        Write-Error $_.Exception.Message
      }
    }
    $Result
  }
}

function Set-SharedMboxCopyForSent {
  # Credits: https://stackoverflow.com/q/51680709
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Shared mailbox address to which to activate the copy of sent emails (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox
  )
  
  begin {
    $mboxCounter = 0
    $Result = @()
    $ResultError = @()
    Set-Variable ProgressPreference Continue
  }

  process {
    $SourceMailbox | ForEach {
      try {
        $CurrentMailbox = $_
        $GetCM = Get-Mailbox $CurrentMailbox
        if ( $GetCM.RecipientTypeDetails -eq "SharedMailbox") {
          $mboxCounter++
          $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
          Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

          Set-Mailbox $CurrentMailbox -MessageCopyForSentAsEnabled $True
          Set-Mailbox $CurrentMailbox -MessageCopyForSendOnBehalfEnabled $True

          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
            PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
            "Copy for SentAs" = $GetCM.MessageCopyForSentAsEnabled
            "Copy for SendOnBehalf" = $GetCM.MessageCopyForSendOnBehalfEnabled
          })
        } else {
          $ResultError += "`e[31m $($CurrentMailbox) is not a Shared Mailbox. `e[0m"
        } 
      } catch {
        Write-Error $_.Exception.Message
      }
    }
    $Result; ""
    $ResultError
  }
}


# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Add-MboxAlias"
Export-ModuleMember -Function "Add-MboxPermission"
Export-ModuleMember -Function "Change-MboxLanguage"
Export-ModuleMember -Function "Export-MboxAlias"
Export-ModuleMember -Function "Export-MboxPermission"
Export-ModuleMember -Function "Get-MboxAlias"
Export-ModuleMember -Function "Get-MboxPermission"
Export-ModuleMember -Function "New-SharedMailbox"
Export-ModuleMember -Function "Remove-MboxAlias"
Export-ModuleMember -Function "Remove-MboxPermission"
Export-ModuleMember -Function "Set-MboxRulesQuota"
Export-ModuleMember -Function "Set-SharedMboxCopyForSent"