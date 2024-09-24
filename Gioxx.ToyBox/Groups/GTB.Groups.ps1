# Groups ===========================================================================================================================================================

function Export-DG {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Distribution Group e-mail address or display name")]
    [string] $DG,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
    [switch] $CSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all Distribution Groups")]
    [switch] $All,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
    [switch] $GridView
  )

  Set-Variable ProgressPreference Continue
  $DGsCounter = 0
  $arr_ExportedDG = @()
  $eolConnectedCheck = priv_CheckEOLConnection
  
  if ( $eolConnectedCheck -eq $true ) {
    if ( [string]::IsNullOrEmpty($DG) ) { $All = $True } else { $DGs = Get-DistributionGroup $DG }

    if ($All) {
      $DGs = Get-DistributionGroup -ResultSize Unlimited
      $CSV = $True
    }
    
    if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
    if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

    $DGs | ForEach-Object {
      try {
        $CurrentDG = $_
        $GetDG = Get-DistributionGroup $CurrentDG
        $DGsCounter++
        $PercentComplete = (($DGsCounter / $DGs.Count) * 100)
        Write-Progress -Activity "Processing $($GetDG.DisplayName)" -Status "$DGsCounter out of $($DGs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        Get-DistributionGroupMember $CurrentDG | ForEach-Object {
          if ($All) {
            $arr_ExportedDG += New-Object -TypeName PSObject -Property $([ordered]@{
              "Group Name" = $GetDG.DisplayName
              "Group Primary Smtp Address" = $GetDG.PrimarySmtpAddress
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          } else {
            $arr_ExportedDG += New-Object -TypeName PSObject -Property $([ordered]@{
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          }
        }
      } catch {
        Write-Error $_.Exception.Message
      }
    }

    if ( $GridView ) {
      $arr_ExportedDG | Out-GridView -Title "M365 Distribution Groups"
    } elseif ( $CSV ) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-DistributionGroups-Report.csv")
      $arr_ExportedDG | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_ExportedDG | Out-Host
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Export-DDG {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Dynamic Distribution Group e-mail address or display name")]
    [string] $DDG,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
    [switch] $CSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all Dynamic Distribution Groups")]
    [switch] $All,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
    [switch] $GridView
  )

  Set-Variable ProgressPreference Continue
  $DDGsCounter = 0
  $arr_ExportedDDG = @()
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    if ( [string]::IsNullOrEmpty($DDG) ) { $All = $True } else { $DDGs = Get-DynamicDistributionGroup $DDG }

    if ($All) {
      $DDGs = Get-DynamicDistributionGroup -ResultSize Unlimited
      $CSV = $True
    }
    
    if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
    if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

    $DDGs | ForEach-Object {
      try {
        $CurrentDDG = $_
        $GetDDG = Get-DynamicDistributionGroup $CurrentDDG
        $DDGsCounter++
        $PercentComplete = (($DDGsCounter / $DDGs.Count) * 100)
        Write-Progress -Activity "Processing $($GetDDG.DisplayName)" -Status "$DDGsCounter out of $($DDGs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        Get-DynamicDistributionGroupMember $CurrentDDG | ForEach-Object {
          if ($All) {
            $arr_ExportedDDG += New-Object -TypeName PSObject -Property $([ordered]@{
              "Group Name" = $GetDDG.DisplayName
              "Group Primary Smtp Address" = $GetDDG.PrimarySmtpAddress
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          } else {
            $arr_ExportedDDG += New-Object -TypeName PSObject -Property $([ordered]@{
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          }
        }
      } catch {
        Write-Error $_.Exception.Message
      }
    }

    if ( $GridView ) {
      $arr_ExportedDDG | Out-GridView -Title "M365 Dynamic Distribution Groups"
    } elseif ( $CSV ) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-DynamicDistributionGroups-Report.csv")
      $arr_ExportedDDG | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_ExportedDDG | Out-Host
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Export-M365Group {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Microsoft 365 Unified Group e-mail address or display name")]
    [string] $M365Group,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
    [switch] $CSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export all Microsoft 365 Unified Group")]
    [switch] $All,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
    [switch] $GridView
  )

  Set-Variable ProgressPreference Continue
  $M365GsCounter = 0
  $arr_ExportedM365Groups = @()
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    if ( [string]::IsNullOrEmpty($M365G) ) { $All = $True } else { $M365Gs = Get-UnifiedGroup $M365G }

    if ($All) {
      $M365Gs = Get-UnifiedGroup -ResultSize Unlimited
      $CSV = $True
    }
    
    if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
    if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

    $M365Gs | ForEach-Object {
      try {
        $CurrentM365G = $_
        $GetM365G = Get-UnifiedGroup $CurrentM365G
        $M365GsCounter++
        $PercentComplete = (($M365GsCounter / $M365Gs.Count) * 100)
        Write-Progress -Activity "Processing $($GetM365G.DisplayName)" -Status "$M365GsCounter out of $($M365Gs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        $GetM365G | Get-UnifiedGroupLinks -LinkType Member | ForEach-Object {
          if ($All) {
            $arr_ExportedM365Groups += New-Object -TypeName PSObject -Property $([ordered]@{
              "Group Name" = $GetM365G.DisplayName
              "Group Primary Smtp Address" = $GetM365G.PrimarySmtpAddress
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          } else {
            $arr_ExportedM365Groups += New-Object -TypeName PSObject -Property $([ordered]@{
              "Member Display Name" = $_.DisplayName
              "Member FirstName" = $_.FirstName
              "Member LastName" = $_.LastName
              "Member Primary Smtp Address" = $_.PrimarySmtpAddress
              "Member Company" = $_.Company
              "Member City" = $_.City
            })
          }
        }
      } catch {
        Write-Error $_.Exception.Message
      }
    }

    if ( $GridView ) {
      $arr_ExportedM365Groups | Out-GridView -Title "M365 Unified Groups"
    } elseif ( $CSV ) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-UnifiedGroups-Report.csv")
      $arr_ExportedM365Groups | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_ExportedM365Groups | Out-Host
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Get-RoleGroupsMembers {
  Set-Variable ProgressPreference Continue
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    $roleGroups = Get-RoleGroup
    $rgCounter = 0

    $permTable = foreach ($rg in $roleGroups) {
        $rgCounter++
        $PercentComplete = (($rgCounter / $roleGroups.Count) * 100)
        Write-Progress -Activity "Processing $($rg)" -Status "$rgCounter out of $($roleGroups.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

        $rgMembers = (Get-RoleGroupMember $rg).DisplayName -join "`n"
        [PSCustomObject]@{
            "Role Group" = $rg
            Count = (Get-RoleGroupMember $rg).Count
            Members = $rgMembers
        }
    }

    $permTable | Sort-Object Count -Descending | Format-Table -AutoSize -Wrap

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }
}

function Get-UserGroups {
  # Credits: https://infrasos.com/get-mgusermemberof-list-group-memberships-of-azure-ad-user-powershell/
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="User to check")]
    [string] $UserPrincipalName,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show more detailed results in a grid view")]
    [switch] $GridView
  )

  $groupList=@()
  $mggConnectedCheck = priv_CheckMGGraphModule

  if ( $mggConnectedCheck -eq $true ) {
    if ( $UserPrincipalName -inotmatch "@" ) {
      try {
        $emailAddresses = (Get-Recipient $UserPrincipalName).EmailAddresses | Where-Object { $_ -clike 'SMTP:*' }
        if ($emailAddresses.Count -le 0) {
          Write-Host "Recipient not available or not found ($($UserPrincipalName))." -f "Red"
          break
        } elseif ($emailAddresses.Count -gt 1) {
          Write-Host "Complete e-mail address not specified, multiple email addresses found:" -f "Cyan"
          $emailAddresses | ForEach-Object { Write-Host $_.Replace('SMTP:', ' ') }
          Write-Host "Run the command again but specify the full address to perform a more accurate search." -f "Cyan"
          $UserPrincipalName = (Get-Recipient $UserPrincipalName).PrimarySmtpAddress | Select-Object -First 1
          Write-Host "First user selected: $($UserPrincipalName)" -f "Cyan"
        } else {
          $UserPrincipalName = (Get-Recipient $UserPrincipalName).PrimarySmtpAddress
          Write-Host "Complete e-mail address not specified, user found: $($UserPrincipalName)" -f "Cyan"
        }
      } catch {
        Write-Error $_.Exception.Message
      }
    }
    
    try {
      $RecipientType = (Get-Recipient $UserPrincipalName -ErrorAction SilentlyContinue).RecipientTypeDetails
      
      if ( $RecipientType -ne $null ) {
        Switch ($RecipientType) {
          "MailContact" { 
            # If you need to analyze a MailContact you must change query in Get-MgContact instead of Get-MgUser / Get-MgContactMemberOf
            # Credits: https://m365scripts.com/microsoft365/effortlessly-manage-office-365-contacts-using-ms-graph-powershell/
            $userID = Get-MgContact -Filter "Mail eq '$UserPrincipalName'"
            $groups = Get-MgContactMemberOf -OrgContactId $userID.Id | Select-Object *
          }
          "UserMailbox" {
            $userID = Get-MgUser -UserId (Get-Recipient $UserPrincipalName).WindowsLiveID
            $groups = Get-MgUserMemberOf -UserId $userID.Id | Select-Object *
          }
          "MailUniversalDistributionGroup" {
            $userID = Get-MgGroup -Filter "Mail eq '$UserPrincipalName'"
            $groups = Get-MgGroupMemberOf -GroupId $userID.Id | Select-Object *
          }
          Default {
            # If the mailbox is not a "UserMailbox" or a "MailContact" (for example a "SharedMailbox"), then the UPN is the WindowsLiveID value.
            $UserPrincipalName = (Get-Recipient $UserPrincipalName).WindowsLiveID
            $userID = Get-MgUser -UserId $UserPrincipalName
            $groups = Get-MgUserMemberOf -UserId $userID.Id | Select-Object *
          }
        }

        if ( $groups -ne $null ) {
          $groups | ForEach-Object {
            $groupIDs = $_.id
            $otherproperties = $_.AdditionalProperties

            if (($otherproperties.groupTypes).count -eq 0) { 
              $groupType = (Get-Recipient $otherproperties.mail).RecipientTypeDetails
            } else { 
              $groupType = $otherproperties.groupTypes
            }

            if ($GridView) {
              $groupList += New-Object -TypeName PSObject -Property $([ordered]@{ 
                "Group Name" = $otherproperties.displayName
                "Group Description" = $otherproperties.description
                "Group Mail" = $otherproperties.mail
                "Group Mail Nickname" = $otherproperties.mailNickname
                "Group Mail Enabled" = $otherproperties.mailEnabled
                "Group Type" = $groupType
                "Group ID" = $groupIDs
              })
            } else {
              $groupList += New-Object -TypeName PSObject -Property $([ordered]@{ 
                "Group Name" = $otherproperties.displayName
                "Group Mail" = $otherproperties.mail
                "Group Type" = $groupType
              })
            }
            
          }
          
          if ($GridView) { $groupList | Out-GridView -Title "M365 User Groups" } else { $groupList | Sort "Group Name" | Out-Host }
        }

      } else {
        Write-Host "Recipient not available or not found ($($UserPrincipalName))." -f "Red"
      }
    } catch {
      Write-Error $_.Exception.Message
    }
  
  } else {
    Write-Host "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -f "Red"
  }
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Export-DG"
Export-ModuleMember -Function "Export-DDG"
Export-ModuleMember -Function "Export-M365Group"
Export-ModuleMember -Function "Get-RoleGroupsMembers"
Export-ModuleMember -Function "Get-UserGroups"