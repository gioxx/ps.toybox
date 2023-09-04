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
  $Result = @()

  if ( [string]::IsNullOrEmpty($DG) ) {
    $All = $True
  } else {
    $DGs = Get-DistributionGroup $DG
  }

  if ($All) {
    $DGs = Get-DistributionGroup -ResultSize Unlimited
    $CSV = $True
  }
  
  if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
  if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

  $DGs | ForEach {
    try {
      $CurrentDG = $_
      $GetDG = Get-DistributionGroup $CurrentDG
      $DGsCounter++
      $PercentComplete = (($DGsCounter / $DGs.Count) * 100)
      Write-Progress -Activity "Processing $($GetDG.DisplayName)" -Status "$DGsCounter out of $($DGs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      Get-DistributionGroupMember $CurrentDG | ForEach {
        if ($All) {
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
    $Result | Out-GridView -Title "M365 Distribution Groups"
  } elseif ( $CSV ) {
    $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-DistributionGroups-Report.csv")
    $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
  } else {
    $Result
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
  $Result = @()

  if ( [string]::IsNullOrEmpty($DDG) ) {
    $All = $True
  } else {
    $DDGs = Get-DynamicDistributionGroup $DDG
  }

  if ($All) {
    $DDGs = Get-DynamicDistributionGroup -ResultSize Unlimited
    $CSV = $True
  }
  
  if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
  if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

  $DDGs | ForEach {
    try {
      $CurrentDDG = $_
      $GetDDG = Get-DynamicDistributionGroup $CurrentDDG
      $DDGsCounter++
      $PercentComplete = (($DDGsCounter / $DDGs.Count) * 100)
      Write-Progress -Activity "Processing $($GetDDG.DisplayName)" -Status "$DDGsCounter out of $($DDGs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      Get-DynamicDistributionGroupMember $CurrentDDG | ForEach {
        if ($All) {
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
    $Result | Out-GridView -Title "M365 Dynamic Distribution Groups"
  } elseif ( $CSV ) {
    $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-DynamicDistributionGroups-Report.csv")
    $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
  } else {
    $Result
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
  $Result = @()

  if ( [string]::IsNullOrEmpty($M365G) ) {
    $All = $True
  } else {
    $M365Gs = Get-UnifiedGroup $M365G
  }

  if ($All) {
    $M365Gs = Get-UnifiedGroup -ResultSize Unlimited
    $CSV = $True
  }
  
  if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
  if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

  $M365Gs | ForEach {
    try {
      $CurrentM365G = $_
      $GetM365G = Get-UnifiedGroup $CurrentM365G
      $M365GsCounter++
      $PercentComplete = (($M365GsCounter / $M365Gs.Count) * 100)
      Write-Progress -Activity "Processing $($GetM365G.DisplayName)" -Status "$M365GsCounter out of $($M365Gs.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      $GetM365G | Get-UnifiedGroupLinks -LinkType Member | ForEach {
        if ($All) {
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
          $Result += New-Object -TypeName PSObject -Property $([ordered]@{
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
    $Result | Out-GridView -Title "M365 Unified Groups"
  } elseif ( $CSV ) {
    $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-UnifiedGroups-Report.csv")
    $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
  } else {
    $Result
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

  $mggConnectedCheck = priv_CheckMGGraphModule
  if ( $mggConnectedCheck -eq $true ) {
    $groupList=@()

    if ( $UserPrincipalName -notcontains "@" ) {
      #Write-Host "DEBUG: PrimarySmtpAddress not specified, searching $($UserPrincipalName)'s PrimarySmtpAddress ..." -f "Yellow"
      $UserPrincipalName = (Get-Recipient $UserPrincipalName).PrimarySmtpAddress
    }

    if ( (Get-Recipient $UserPrincipalName | Select RecipientTypeDetails) -ne "UserMailbox" ) {
      # Se la casella non è personale, allora l'utente da passare a MgUser è quello del WindowsLiveID.
      $UserPrincipalName = ( Get-Recipient $UserPrincipalName).WindowsLiveID
    }

    $userID = Get-MgUser -UserId $UserPrincipalName
    $groups = Get-MgUserMemberOf -UserId $userID.Id | Select-Object *
    
    $groups | ForEach {
      $groupIDs = $_.id
      $otherproperties = $_.AdditionalProperties

      if ( $GridView ) {
        $groupList += New-Object -TypeName PSObject -Property $([ordered]@{ 
          "Group Name" = $otherproperties.displayName
          "Group Description" = $otherproperties.description
          "Group Mail" = $otherproperties.mail
          "Group Mail Nickname" = $otherproperties.mailNickname
          "Group Mail Enabled" = $otherproperties.mailEnabled
          "Group ID" = $groupIDs
        })
      } else {
        $groupList += New-Object -TypeName PSObject -Property $([ordered]@{ 
          "Group Name" = $otherproperties.displayName
          "Group Mail" = $otherproperties.mail
        })
      }
      
    }
    
    if ( $GridView ) { $groupList | Out-GridView -Title "M365 User Groups" } else { $groupList }

  } else {
    Write-Host "`nCan't connect or use Microsoft Graph Modules. `nPlease check logs." -f "Red"
  }
}

# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Export-DG"
Export-ModuleMember -Function "Export-DDG"
Export-ModuleMember -Function "Export-M365Group"
Export-ModuleMember -Function "Get-UserGroups"