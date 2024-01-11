# Mailboxes ========================================================================================================================================================

function Add-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="User to edit (e.g. mario.rossi)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Alias to be added (e.g. mario.rossi.alias@contoso.com)")]
    [string] $MailboxAlias
  )

  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {

    try {
      $GRRTD = (Get-Recipient $SourceMailbox -ErrorAction Stop).RecipientTypeDetails
    } catch {
      Write-Host "`nUsage: Add-MboxAlias -SourceMailbox mailbox@contoso.com -MailboxAlias alias@contoso.com`n" -f "Yellow"
      Write-Error $_.Exception.Message
    }

    Switch ($GRRTD) {
      "MailContact" { Set-MailContact $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
      "MailUser" { Set-MailUser $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
      Default { Set-Mailbox $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
    }

    Get-Recipient $SourceMailbox | 
        Select-Object Name -Expand EmailAddresses | 
        Where {$_ -like 'smtp*'}

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Add-MboxPermission {
  param(
    [Parameter(Mandatory=$True, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to which to allow access (e.g. mario.rossi@contoso.com)")]
    [string[]] $UserMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Type of access to be allowed (All, FullAccess, SendAs, SendOnBehalfTo)")]
    [string] $AccessRights,
    [Parameter(Mandatory=$False, HelpMessage="Set mailbox automapping")]
    [switch] $AutoMapping
  )

  begin {
    if ([string]::IsNullOrEmpty($AccessRights)) { $AccessRights = "All" }
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    
    if ( $eolConnectedCheck -eq $true ) {
      $UserMailbox | ForEach {
        $CurrentUser = $_
        Switch ($AccessRights) {
          "FullAccess" {
            if ($AutoMapping) {
              Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
              Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False | Out-Host
            } else {
              Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) without AutoMapping ..."
              Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False | Out-Host
            }
          }
          "SendAs" {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False | Out-Host
          }
          "SendOnBehalfTo" {
            Write-Host "Add $($CurrentUser) ($($AccessRights)) on $($SourceMailbox) ..."
            Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{add="$($CurrentUser)"} -Confirm:$False | Out-Host
          }
          "All" {
            if ($AutoMapping) {
              Write-Host "Add $($CurrentUser) (FullAccess) on $($SourceMailbox) ..."
              Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False | Out-Host
              Write-Host "Add $($CurrentUser) (SendAs) on $($SourceMailbox) ..."
              Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False | Out-Host
            }
            else {
              Write-Host "Add $($CurrentUser) (FullAccess) on $($SourceMailbox) without AutoMapping ..."
              Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False | Out-Host
              Write-Host "Add $($CurrentUser) (SendAs) on $($SourceMailbox) ..."
              Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False | Out-Host
            }
          }
        }
      }
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
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
  $eolConnectedCheck = priv_CheckEOLConnection
    
  if ( $eolConnectedCheck -eq $true ) {
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

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
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
    $arr_MboxAliases = @()
    Set-Variable ProgressPreference Continue
    $eolConnectedCheck = priv_CheckEOLConnection
    
    if ( $eolConnectedCheck -eq $true ) {

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
    
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
      Return
    }
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
              $arr_MboxAliases += New-Object -TypeName PSObject -Property $([ordered]@{
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
      $arr_MboxAliases | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_MboxAliases | Out-Host
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

  $arr_MboxPerms = @()
  $mboxCounter = 0
  Set-Variable ProgressPreference Continue
  # $warningPrefBackup = priv_HideWarning
  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {

    Switch ($RecipientType) {
      "User" { $Mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where { $_.RecipientTypeDetails -eq "UserMailbox" } }
      "Shared" { $Mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where { $_.RecipientTypeDetails -eq "SharedMailbox" } }
      "Room" { $Mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where { $_.RecipientTypeDetails -eq "RoomMailbox" } }
      "All" { 
        Write-Host "WARNING: no recipient type specified, I scan all the types now (User, Shared, Room), please be patient." -f "Yellow"
        $Mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | 
          Where { $_.RecipientTypeDetails -eq "UserMailbox" -Or $_.RecipientTypeDetails -eq "SharedMailbox" -Or $_.RecipientTypeDetails -eq "RoomMailbox" }
      }
    }

    $Mailboxes | ForEach {
      $CurrentMailbox = $_
      $GetCM = Get-EXOMailbox $CurrentMailbox
      
      $mboxCounter++
      $PercentComplete = (($mboxCounter / $Mailboxes.Count) * 100)
      Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($Mailboxes.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      $MboxPermSendAs = Get-RecipientPermission $GetCM.PrimarySmtpAddress -AccessRights SendAs |
          Where { $_.Trustee.ToString() -ne "NT AUTHORITY\SELF" -And $_.Trustee.ToString() -notlike "S-1-5*" } |
          ForEach { $_.Trustee.ToString() }

      $MboxPermFullAccess = Get-MailboxPermission $GetCM.PrimarySmtpAddress |
          Where { $_.AccessRights -eq "FullAccess" -and !$_.IsInherited } |
          ForEach { $_.User.ToString() }

      $arr_MboxPerms += New-Object -TypeName PSObject -Property $([ordered]@{
        Mailbox = $GetCM.DisplayName
        "Mailbox Address" = $GetCM.PrimarySmtpAddress
        "Recipient Type" = $GetCM.RecipientTypeDetails
        FullAccess = $MboxPermFullAccess -join ", "
        SendAs = $MboxPermSendAs -join ", "
        SendOnBehalfTo = $GetCM.GrantSendOnBehalfTo
      })
    }

    $folder = priv_CheckFolder($folderCSV)
    $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MboxPermissions-Report.csv")
    $arr_MboxPerms | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
  
  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }

  # $WarningPreference = $warningPrefBackup
}

function Get-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to be analyzed (e.g. info@contoso.com)")]
    [string] $SourceMailbox
  )
  
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    
    $arr_Alias = @()
    $getAddresses = Get-Recipient $SourceMailbox -ErrorAction SilentlyContinue

    if ( $getAddresses -ne $null ) {
      $getAddresses | Select-Object Name -Expand EmailAddresses | ForEach-Object {
        if ($_ -clike 'smtp:*') {
          $arr_Alias += [PSCustomObject]@{
            Alias = $_.Replace('smtp:', '')
          }
        } elseif ($_ -clike 'SMTP:*') {
          $getPrimary = $_.Replace('SMTP:', '')
        }
      }

      Write-Host "PrimarySmtpAddress: $($getPrimary)" -f "Cyan" -NoNewLine
      $arr_Alias | ft -HideTableHeaders | Out-Host

    } else {
      Write-Host "Recipient not available or not found." -f "Red"
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Get-MboxPermission {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox e-mail address or display name (e.g. mario.rossi@contoso.com)")]
    [string] $SourceMailbox
  )
  
  $arr_MbxPerms = @()
  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {
    
    $MboxPermFullAccess = Get-MailboxPermission $(Get-Mailbox $SourceMailbox).PrimarySmtpAddress |
        Where-Object { $_.AccessRights -eq "FullAccess" -and !$_.IsInherited } |
        ForEach-Object {
            $UserMailbox = $_.User.ToString()
            $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

            $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
            if ($existingUserObject) {
                $existingUserObject.AccessRights += ", FullAccess"
            } else {
                $arr_MbxPerms += [PSCustomObject]@{
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
            $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

            $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
            if ($existingUserObject) {
                $existingUserObject.AccessRights += ", SendAs"
            } else {
                $arr_MbxPerms += [PSCustomObject]@{
                    User = $DisplayName
                    UserMailbox = $PrimarySmtpAddress
                    AccessRights = "SendAs"
                }
            }
        }

    $MboxPermSendOnBehalfTo = $(Get-Mailbox $SourceMailbox).GrantSendOnBehalfTo |
        ForEach-Object {
            $UserMailbox = $_
            $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $DisplayName = $(Get-User -Identity $UserMailbox).DisplayName

            $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
            if ($existingUserObject) {
                $existingUserObject.AccessRights += ", SendOnBehalfTo"
            } else {
                $arr_MbxPerms += [PSCustomObject]@{
                    User = $DisplayName
                    UserMailbox = $PrimarySmtpAddress
                    AccessRights = "SendOnBehalfTo"
                }
            }
        }

    Write-Host "`nAccess Rights on $((Get-Mailbox $SourceMailbox).DisplayName) ($((Get-Mailbox $SourceMailbox).PrimarySmtpAddress))" -f "Yellow"
    $arr_MbxPerms | Out-Host

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
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

  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySmtpAddress $SharedMailboxSMTPAddress
    Write-Host "Set outgoing e-mail copy save for $($SharedMailboxSMTPAddress)" -f "Yellow"
    Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSentAsEnabled $True
    Set-Mailbox $SharedMailboxSMTPAddress -MessageCopyForSendOnBehalfEnabled $True
    Write-Host "All done, remember to set access and editing rights to the new mailbox."
  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Remove-MboxAlias {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="User to edit (e.g. mario.rossi)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Alias to be removed (e.g. mario.rossi.alias@contoso.com)")]
    [string] $MailboxAlias
  )

  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {

    Switch ($(Get-Recipient $SourceMailbox).RecipientTypeDetails) {
      "MailContact" { Set-MailContact $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
      "MailUser" { Set-MailUser $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
      Default { Set-Mailbox $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
    }
    
    Get-Recipient $SourceMailbox | 
        Select-Object Name -Expand EmailAddresses | 
        Where {$_ -like 'smtp*'}

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Remove-MboxPermission {
  param(
    [Parameter(Mandatory=$True, HelpMessage="E-mail address of the mailbox to which the permissions are to be changed (e.g. info@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="E-mail address of the mailbox to which to remove access (e.g. mario.rossi@contoso.com)")]
    [string[]] $UserMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Type of access to be removed (All, FullAccess, SendAs, SendOnBehalfTo)")]
    [string] $AccessRights
  )

  begin {
    if ([string]::IsNullOrEmpty($AccessRights)) { $AccessRights = "All" }
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    
    if ( $eolConnectedCheck -eq $true ) {
      $UserMailbox | ForEach {
        $CurrentUser = $_
        Switch ($AccessRights) {
          "FullAccess" { Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False }
          "SendAs" { Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False }
          "SendOnBehalfTo" { Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{remove="$($CurrentUser)"} -Confirm:$False }
          "All" {
            Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False
            Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
            Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{remove="$($CurrentUser)"} -Confirm:$False
          }
        }
      }

    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
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
    $arr_MboxRulesQuota = @()
    Set-Variable ProgressPreference Continue
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      $SourceMailbox | ForEach {
        try {
          $CurrentMailbox = $_
          $GetCM = Get-Recipient $CurrentMailbox
          
          $mboxCounter++
          $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
          Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

          Set-Mailbox $CurrentMailbox -RulesQuota 256KB

          $arr_MboxRulesQuota += New-Object -TypeName PSObject -Property $([ordered]@{
            PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
            "Rules Quota" = $GetCM.RulesQuota
          })
        } catch {
          Write-Error $_.Exception.Message
        }
      }
      
      $arr_MboxRulesQuota | Out-Host

    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
    }
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
    $arr_SharedMboxCFS = @()
    $arr_SharedMboxCFSError = @()
    Set-Variable ProgressPreference Continue
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      $SourceMailbox | ForEach {
        try {
          $CurrentMailbox = $_
          $GetCM = Get-Recipient $CurrentMailbox
          if ( $GetCM.RecipientTypeDetails -eq "SharedMailbox") {
            $mboxCounter++
            $PercentComplete = (($mboxCounter / $SourceMailbox.Count) * 100)
            Write-Progress -Activity "Processing $($GetCM.PrimarySmtpAddress)" -Status "$mboxCounter out of $($SourceMailbox.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

            Set-Mailbox $CurrentMailbox -MessageCopyForSentAsEnabled $True
            Set-Mailbox $CurrentMailbox -MessageCopyForSendOnBehalfEnabled $True

            $arr_SharedMboxCFS += New-Object -TypeName PSObject -Property $([ordered]@{
              PrimarySmtpAddress = $GetCM.PrimarySmtpAddress
              "Copy for SentAs" = $GetCM.MessageCopyForSentAsEnabled
              "Copy for SendOnBehalf" = $GetCM.MessageCopyForSendOnBehalfEnabled
            })
          } else {
            $arr_SharedMboxCFSError += "`e[31m $($CurrentMailbox) is not a Shared Mailbox. `e[0m"
          } 
        } catch {
          Write-Error $_.Exception.Message
        }
      }
      $arr_SharedMboxCFS | Out-Host; ""
      $arr_SharedMboxCFSError | Out-Host

    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
    }
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