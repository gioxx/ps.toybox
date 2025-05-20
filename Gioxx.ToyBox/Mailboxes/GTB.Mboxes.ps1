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
      $recipient = Get-Recipient $SourceMailbox -ErrorAction Stop
      $GRRTD = $recipient.RecipientTypeDetails
      $existingAliases = $recipient.EmailAddresses
    } catch {
      Write-Warning "`nUsage: Add-MboxAlias -SourceMailbox mailbox@contoso.com -MailboxAlias alias@contoso.com`n"
      Write-Error $_.Exception.Message
      return
    }

    # Check if the alias already exists
    if ($existingAliases -contains $MailboxAlias) {
      Write-InformationColored "`nAlias '$MailboxAlias' already exists for mailbox '$SourceMailbox'. No action taken." -ForegroundColor "Yellow"
      return
    }

    Switch ($GRRTD) {
      "MailContact" { Set-MailContact $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
      "MailUser" { Set-MailUser $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
      { ($_ -eq "MailUniversalDistributionGroup") -or ($_ -eq "DynamicDistributionGroup") -or ($_ -eq "MailUniversalSecurityGroup") } {
        Set-DistributionGroup $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"}
      }
      Default { Set-Mailbox $SourceMailbox -EmailAddresses @{add="$($MailboxAlias)"} }
    }

    Get-MboxAlias -SourceMailbox $SourceMailbox

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
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
    priv_SetPreferences -Verbose
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    $SourceMailbox = $SourceMailbox.ToLower()
    $UserMailbox = $UserMailbox.ToLower()
    
    if ( $eolConnectedCheck -eq $true ) {
      $UserMailbox | ForEach {
        $CurrentUser = $_

        # Checks if the user exists
        $userExists = Get-User -Identity $CurrentUser -ErrorAction SilentlyContinue
        if (-not $userExists) {
          Write-Error "`nThe mailbox $($CurrentUser) does not exist. Please check the provided e-mail address."
          return
        }

        # Continue with assigning permissions only if the user exists
        Switch ($AccessRights) {
          "FullAccess" {
            $existingPermission = Get-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and $_.IsInherited -eq $false }
            if (-not $existingPermission) {
              if ($AutoMapping) {
                Write-Output "`nAdd $($CurrentUser) (FullAccess) to $($SourceMailbox) ..."
                $addMboxPerm = Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False
              } else {
                Write-Output "`nAdd $($CurrentUser) (FullAccess) to $($SourceMailbox) without AutoMapping ..."
                $addMboxPerm = Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False
              }
              $addMboxPermDN = (Get-User -Identity $addMboxPerm.User).DisplayName
              [PSCustomObject]@{
                Identity = $addMboxPerm.Identity
                User = $addMboxPerm.User
                DisplayName = $addMboxPermDN
                AccessRights = $addMboxPerm.AccessRights
                IsInherited = $addMboxPerm.IsInherited
                Deny = $addMboxPerm.Deny
              } | Out-Host
            } else {
              Write-InformationColored "`n$($CurrentUser) already has FullAccess permission to $($SourceMailbox), skip." -ForegroundColor "Yellow"
            }
          }
          "SendAs" {
            $existingPermission = Get-RecipientPermission -Identity $SourceMailbox -Trustee $CurrentUser -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'SendAs' }
            if (-not $existingPermission) {
              Write-Output "`nAdd $($CurrentUser) (SendAs) to $($SourceMailbox) ..."
              $addMboxPerm = Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
              $addMboxPermDN = (Get-User -Identity $addMboxPerm.Trustee).DisplayName
              [PSCustomObject]@{
                Identity = $addMboxPerm.Identity
                Trustee = $addMboxPerm.Trustee
                DisplayName = $addMboxPermDN
                AccessControlType = $addMboxPerm.AccessControlType
                AccessRights = $addMboxPerm.AccessRights
              } | Out-Host
            } else {
              Write-InformationColored "`n$($CurrentUser) already has SendAs permission to $($SourceMailbox), skip." -ForegroundColor "Yellow"
            }
          }
          "SendOnBehalfTo" {
            $existingPermission = (Get-Mailbox -Identity $SourceMailbox).GrantSendOnBehalfTo | Where-Object { $_ -eq $CurrentUser }
            if (-not $existingPermission) {
              Write-Output "`nAdd $($CurrentUser) (SendOnBehalfTo) to $($SourceMailbox) ..."
              Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{add="$($CurrentUser)"} -Confirm:$False | Out-Host
            } else {
              Write-InformationColored "`n$($CurrentUser) already has SendOnBehalfTo permission to $($SourceMailbox), skip." -ForegroundColor "Yellow"
            }
          }
          "All" {
            # FullAccess
            $existingFullAccess = Get-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and $_.IsInherited -eq $false }
            if (-not $existingFullAccess) {
              if ($AutoMapping) {
                Write-Output "`nAdd $($CurrentUser) (FullAccess) to $($SourceMailbox) ..."
                $addMboxPerm = Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$True -Confirm:$False
              } else {
                Write-Output "`nAdd $($CurrentUser) (FullAccess) to $($SourceMailbox) without AutoMapping ..."
                $addMboxPerm = Add-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -AutoMapping:$False -Confirm:$False
              }
              $addMboxPermDN = (Get-User -Identity $addMboxPerm.User).DisplayName
              [PSCustomObject]@{
                Identity = $addMboxPerm.Identity
                User = $addMboxPerm.User
                DisplayName = $addMboxPermDN
                AccessRights = $addMboxPerm.AccessRights
                IsInherited = $addMboxPerm.IsInherited
                Deny = $addMboxPerm.Deny
              } | Out-Host
            } else {
              Write-InformationColored "`n$($CurrentUser) already has FullAccess permission to $($SourceMailbox), skip." -ForegroundColor "Yellow"
            }

            # SendAs
            $existingSendAs = Get-RecipientPermission -Identity $SourceMailbox -Trustee $CurrentUser -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'SendAs' }
            if (-not $existingSendAs) {
              Write-Output "`nAdd $($CurrentUser) (SendAs) to $($SourceMailbox) ..."
              $addMboxPerm = Add-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
              $addMboxPermDN = (Get-User -Identity $addMboxPerm.Trustee).DisplayName
              [PSCustomObject]@{
                Identity = $addMboxPerm.Identity
                Trustee = $addMboxPerm.Trustee
                DisplayName = $addMboxPermDN
                AccessControlType = $addMboxPerm.AccessControlType
                AccessRights = $addMboxPerm.AccessRights
              } | Out-Host
            } else {
              Write-InformationColored "`n$($CurrentUser) already has SendAs permission to $($SourceMailbox), skip." -ForegroundColor "Yellow"
            }
          }
        }
      }
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }

    priv_RestorePreferences
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
  
  priv_SetPreferences -Verbose
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
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
}

function Check-SharedMailboxCompliance {
  priv_SetPreferences -Verbose
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {

    $mggConnectedCheck = priv_CheckMGGraphModule
    if ( $mggConnectedCheck -eq $true ) {
      Write-InformationColored "`nFinding shared mailboxes... " -NoNewLine
      $Mbx = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
      If ($Mbx) {
          Write-InformationColored ("{0} shared mailboxes found." -f $Mbx.Count) -ForegroundColor "Cyan"
          $mboxCounter = 0
      } Else {
          Write-InformationColored "No shared mailboxes found." -ForegroundColor "Red"
          Break
      }
      # Define the service plan IDs for Exchange Online (Plan 1) and Exchange Online (Plan 2)
      $ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
      $ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 
      $Report = [System.Collections.Generic.List[Object]]::new()

      ForEach ($M in $Mbx) {
          # $mboxCounter++
          # $PercentComplete = (($mboxCounter / $Mbx.Count) * 100)
          # Write-Progress -Activity "Processing $((Get-Recipient $M).DisplayName)" -Status "$mboxCounter out of $($Mbx.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

          $ExoPlan1Found = $false; $ExoPlan2Found = $false; $LogsFound = "No"
          Write-Output ("Checking sign-in records for {0}" -f $M.DisplayName)
          $UserId = $M.ExternalDirectoryObjectId
          [array]$Logs = Get-MgAuditLogSignIn -Filter "userid eq '$UserId'" -All -Top 1
          If ($Logs) {
              
              # Check if there are successful sign-in records (at least one), otherwise, don't keep track of them (these could be attack attempts)
              Write-Output ("Checking for successful sign-in records for {0}" -f $M.DisplayName)
              [array]$search4SuccessfulLogins = Get-MgAuditLogSignIn -Filter "userid eq '$UserId'" -All
              $search4SuccessfulLogins | ForEach-Object {
                  If ($_.Status.ErrorCode -eq "0") {
                      $LogsFound = "Yes"
                  }
              }
              
              If ($LogsFound -eq "Yes") {
                  Write-InformationColored ("Sign-in records found for shared mailbox {0}" -f $M.DisplayName) -ForegroundColor "Red"
                  # Check if the shared mailbox is licensed
                  $User = Get-MgUser -UserId $M.ExternalDirectoryObjectId -Property UserPrincipalName, AccountEnabled, Id, DisplayName, assignedPlans
                  [array]$ExoPlans = $User.AssignedPlans | Where-Object {$_.Service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled'}
                  If ($ExoServicePlan1 -in $ExoPlans.ServicePlanId) {
                      $ExoPlan1Found = $true

                  } ElseIf ($ExoServicePlan2 -in $ExoPlans.ServicePlanId) {
                      $ExoPlan2Found = $true
                  }
              
                  If ($ExoPlan1Found -eq $true) {
                      Write-Output ("Shared mailbox {0} has Exchange Online (Plan 1) license" -f $M.DisplayName)
                  } ElseIf ($ExoPlan2Found -eq $true) {
                      Write-Output ("Shared mailbox {0} has Exchange Online (Plan 2) license" -f $M.DisplayName)
                  }  Else {
                      Write-InformationColored ("Shared mailbox {0} has no Exchange Online license" -f $M.DisplayName) -ForegroundColor "Yellow"
                  }   
              } else {
                  Write-InformationColored ("No successful sign-in records found for shared mailbox {0}" -f $M.DisplayName) -ForegroundColor "Green"
              }
          } 

          $ReportLine = [PSCustomObject] @{ 
              DisplayName                 = $M.DisplayName
              ExternalDirectoryObjectId   = $M.ExternalDirectoryObjectId
              'Sign in Record Found'      = $LogsFound
              'Exchange Online Plan 1'    = $ExoPlan1Found
              'Exchange Online Plan 2'    = $ExoPlan2Found
          }
          $Report.Add($ReportLine)
      }

      $Report | Out-GridView -Title "Shared Mailbox Sign-In Records and Licensing Status" 
      
    } else {
      Write-Error "`nCan't connect or use Microsoft Graph modules. `nPlease check logs."
    }
  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
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
    priv_SetPreferences -Verbose
    $eolConnectedCheck = priv_CheckEOLConnection
    $mboxCounter = 0
    $arr_MboxAliases = @()
    
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
        Write-Warning "WARNING: no mailbox(es) specified, I scan all the mailboxes, please be patient."
        $SourceMailbox = Get-Recipient -ResultSize Unlimited | 
            Where { $_.RecipientTypeDetails -ne "GuestMailUser" }
        $CSV = $True
      }
      
      if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
      if ($CSV) { $folder = priv_CheckFolder($folderCSV) }
    
    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
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

    priv_RestorePreferences
  }

  end {
    if ($CSV) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-Alias-Report.csv")
      $arr_MboxAliases | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_MboxAliases | Out-Host
    }

    priv_RestorePreferences
  }
}

function Export-MboxPermission {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Which type of box to analyze (User/Shared/Room/All)")]
    [string] $RecipientType,
    [Parameter(Mandatory=$False, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV
  )

  priv_SetPreferences -Verbose
  $arr_MboxPerms = @()
  $mboxCounter = 0
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
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
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
      Write-Error "Recipient not available or not found."
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Get-MboxPermission {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox e-mail address or display name (e.g. mario.rossi@contoso.com)")]
    [string] $SourceMailbox,
    [Parameter(Mandatory=$False, HelpMessage="Include a summary of permissions counts")]
    [switch] $IncludeSummary
  )
  
  $arr_MbxPerms = @()
  $fullAccessCount = 0
  $sendAsCount = 0
  $sendOnBehalfToCount = 0
  $eolConnectedCheck = priv_CheckEOLConnection
  priv_SetPreferences -Verbose

  if ( $eolConnectedCheck -eq $true ) {

    # Controlla se la casella esiste
    $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction SilentlyContinue
    if (-not $mailbox) {
      Write-Error "Mailbox '$SourceMailbox' not found."
      return
    }

    $MboxPermFullAccess = Get-MailboxPermission $mailbox.PrimarySmtpAddress | Where-Object { $_.AccessRights -eq "FullAccess" -and !$_.IsInherited } | ForEach-Object {
      $UserMailbox = $_.User.ToString()
      $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
      $DisplayName = $(Get-User -Identity $UserMailbox -ErrorAction SilentlyContinue).DisplayName

      if ($PrimarySmtpAddress) {
        $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
        if ($existingUserObject) {
            $existingUserObject.AccessRights += ", FullAccess"
        } else {
            $arr_MbxPerms += [PSCustomObject]@{
                User = $DisplayName
                UserMailbox = $PrimarySmtpAddress
                AccessRights = "FullAccess"
            }
            $fullAccessCount++
        }
      }
    }
    Write-Progress -Activity "Gathered FullAccess permissions for $($SourceMailbox) ..." -Status "35% Complete" -PercentComplete 35

    $MboxPermSendAs = Get-RecipientPermission $mailbox.PrimarySmtpAddress -AccessRights SendAs -ErrorAction SilentlyContinue | Where-Object { $_.Trustee.ToString() -ne "NT AUTHORITY\SELF" -And $_.Trustee.ToString() -notlike "S-1-5*" } | ForEach-Object {
      $UserMailbox = $_.Trustee.ToString()
      $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
      $DisplayName = $(Get-User -Identity $UserMailbox -ErrorAction SilentlyContinue).DisplayName

      if ($PrimarySmtpAddress) {
        $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
        if ($existingUserObject) {
          $existingUserObject.AccessRights += ", SendAs"
        } else {
          $arr_MbxPerms += [PSCustomObject]@{
            User = $DisplayName
            UserMailbox = $PrimarySmtpAddress
            AccessRights = "SendAs"
          }
          $sendAsCount++
        }
      }
    }
    Write-Progress -Activity "Gathered SendAs permissions for $($SourceMailbox) ..." -Status "50% Complete" -PercentComplete 50

    $MboxPermSendOnBehalfTo = $mailbox.GrantSendOnBehalfTo | ForEach-Object {
      $UserMailbox = $_
      $PrimarySmtpAddress = $(Get-Mailbox $UserMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
      $DisplayName = $(Get-User -Identity $UserMailbox -ErrorAction SilentlyContinue).DisplayName

      if ($PrimarySmtpAddress) {
        $existingUserObject = $arr_MbxPerms | Where-Object { $_.UserMailbox -eq $PrimarySmtpAddress }
        if ($existingUserObject) {
          $existingUserObject.AccessRights += ", SendOnBehalfTo"
        } else {
          $arr_MbxPerms += [PSCustomObject]@{
            User = $DisplayName
            UserMailbox = $PrimarySmtpAddress
            AccessRights = "SendOnBehalfTo"
          }
          $sendOnBehalfToCount++
        }
      }
    }
    Write-Progress -Activity "Gathered SendOnBehalfTo permissions for $($SourceMailbox) ..." -Status "90% Complete" -PercentComplete 90

    Write-InformationColored "`nAccess Rights on $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))" -ForegroundColor "Yellow"
    $arr_MbxPerms | Out-Host

    if ($IncludeSummary) {
      Write-Host "`nSummary of Permissions Found:" -ForegroundColor Cyan
      Write-Host "FullAccess: $fullAccessCount" -ForegroundColor Green
      Write-Host "SendAs: $sendAsCount" -ForegroundColor Green
      Write-Host "SendOnBehalfTo: $sendOnBehalfToCount" -ForegroundColor Green
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
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
    Set-Mailbox $SharedMailboxSMTPAddress -RetainDeletedItemsFor 30
    Write-Host "All done, remember to set access and editing rights to the new mailbox."
  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
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

    try {
      $GRRTD = (Get-Recipient $SourceMailbox -ErrorAction Stop).RecipientTypeDetails
    } catch {
      Write-Warning "`nUsage: Remove-MboxAlias -SourceMailbox mailbox@contoso.com -MailboxAlias alias@contoso.com`n"
      Write-Error $_.Exception.Message
    }

    Switch ($GRRTD) {
      "MailContact" { Set-MailContact $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
      "MailUser" { Set-MailUser $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
      { ($_ -eq "MailUniversalDistributionGroup") -or ($_ -eq "DynamicDistributionGroup") -or ($_ -eq "MailUniversalSecurityGroup") } {
        # Credits: https://stackoverflow.com/a/3493826/2220346
        Set-DistributionGroup $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"}
      }
      Default { Set-Mailbox $SourceMailbox -EmailAddresses @{remove="$($MailboxAlias)"} }
    }

    Get-MboxAlias -SourceMailbox $SourceMailbox

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
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
    priv_SetPreferences -Verbose
    if ([string]::IsNullOrEmpty($AccessRights)) { $AccessRights = "All" }
  }

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    $SourceMailbox = $SourceMailbox.ToLower()
    $UserMailbox = $UserMailbox.ToLower()
    
    if ( $eolConnectedCheck -eq $true ) {
      $UserMailbox | ForEach {
        $CurrentUser = $_

        # Checks if the user exists
        $userExists = Get-User -Identity $CurrentUser -ErrorAction SilentlyContinue
        if (-not $userExists) {
          Write-Error "`nThe mailbox $($CurrentUser) does not exist. Please check the provided e-mail address."
          return
        }

        # Continue with removing permissions only if the user exists
        Switch ($AccessRights) {
          "FullAccess" { 
            Write-Output "Removing Full Access for $($CurrentUser) from $($SourceMailbox) ..."
            Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False 
          }
          "SendAs" { 
            Write-Output "Removing SendAs for $($CurrentUser) from $($SourceMailbox) ..." 
            Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False 
          }
          "SendOnBehalfTo" { 
            Write-Output "Removing SendOnBehalfTo for $($CurrentUser) from $($SourceMailbox) ..."
            Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{remove="$($CurrentUser)"} -Confirm:$False 
          }
          "All" {
            Write-Output "Removing Full Access for $($CurrentUser) from $($SourceMailbox) ..."
            Remove-MailboxPermission -Identity $SourceMailbox -User $CurrentUser -AccessRights FullAccess -Confirm:$False
            Write-Output "Removing SendAs for $($CurrentUser) from $($SourceMailbox) ..."
            Remove-RecipientPermission $SourceMailbox -Trustee $CurrentUser -AccessRights SendAs -Confirm:$False
            Write-Output "Removing SendOnBehalfTo for $($CurrentUser) from $($SourceMailbox) ..."
            Set-Mailbox $SourceMailbox -GrantSendOnBehalfTo @{remove="$($CurrentUser)"} -Confirm:$False
          }
        }
      }

    } else {
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }

    priv_RestorePreferences
  }
}

function Set-MboxRulesQuota {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Mailbox address to which to expand space for rules (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox
  )
  
  begin {
    priv_SetPreferences -Verbose
    $mboxCounter = 0
    $arr_MboxRulesQuota = @()
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
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }

    priv_RestorePreferences
  }
}

function Set-SharedMboxCopyForSent {
  # Credits: https://stackoverflow.com/q/51680709
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Shared mailbox address to which to activate the copy of sent e-mails (e.g. info@contoso.com)")]
    [string[]] $SourceMailbox
  )
  
  begin {
    priv_SetPreferences -Verbose
    $eolConnectedCheck = priv_CheckEOLConnection
    $mboxCounter = 0
    $arr_SharedMboxCFS = @()
    $arr_SharedMboxCFSError = @()
  }

  process {
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
      Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
    }

    priv_RestorePreferences
  }
}


# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Add-MboxAlias"
Export-ModuleMember -Function "Add-MboxPermission"
Export-ModuleMember -Function "Change-MboxLanguage"
Export-ModuleMember -Function "Check-SharedMailboxCompliance"
Export-ModuleMember -Function "Export-MboxAlias"
Export-ModuleMember -Function "Export-MboxPermission"
Export-ModuleMember -Function "Get-MboxAlias"
Export-ModuleMember -Function "Get-MboxPermission"
Export-ModuleMember -Function "New-SharedMailbox"
Export-ModuleMember -Function "Remove-MboxAlias"
Export-ModuleMember -Function "Remove-MboxPermission"
Export-ModuleMember -Function "Set-MboxRulesQuota"
Export-ModuleMember -Function "Set-SharedMboxCopyForSent"