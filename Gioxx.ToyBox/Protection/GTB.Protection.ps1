# Protection =======================================================================================================================================================

function Change-MFAStatus {
  # Credits:
  #   https://technet440.rssing.com/chan-6827930/article18082.html
  Param(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage="User principal name (es. mario.rossi@contoso.com)")]
    [string] $UserPrincipalName
  )

  $previousInformationPreference = $InformationPreference
  Set-Variable InformationPreference Continue

  $msolServiceConnectedCheck = priv_CheckMsolEmbeddedService
  if ( $msolServiceConnectedCheck -eq $true ) {
    
    $authMFA = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $authMFA.RelyingParty = "*"
    $authMFA.State = "Enabled"
    $authMFA.RememberDevicesNotIssuedBefore = (Get-Date)

    if ( -not (Get-MsolUser -UserPrincipalName $UserPrincipalName | Select -ExpandProperty StrongAuthenticationRequirements ) ) {
      Write-InformationColored "No StrongAuthenticationRequirements found. `nSetting new StrongAuthenticationRequirements." -ForegroundColor "Yellow"
      Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements $authMFA
    } else {
      Write-InformationColored "StrongAuthenticationRequirements already enabled. `nDisabling it." -ForegroundColor "Cyan"
      Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements @()
    }

  } else {
    Write-Error "`nCan't connect or use MsolService using Windows PowerShell. `nPlease check logs."
  }

  Set-Variable InformationPreference $previousInformationPreference
}

function Export-MFAStatus {
  # Credits:
  #   https://www.alitajran.com/get-mfa-status-entra/
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV
  )
  
  $previousInformationPreference = $InformationPreference
  Set-Variable InformationPreference Continue
  $previousProgressPreference = $ProgressPreference
  Set-Variable ProgressPreference Continue
  
  $folder = priv_CheckFolder($folderCSV)
  $mggConnectedCheck = priv_CheckMGGraphModule
  
  if ( $mggConnectedCheck -eq $true ) {
    try {
      Connect-MgGraph -Scopes "Reports.Read.All,AuditLog.Read.All"
      # Fetch user registration detail report from Microsoft Graph
      $Users = Get-MgReportAuthenticationMethodUserRegistrationDetail

      # Create custom PowerShell object and populate it with the desired properties
      $Report = foreach ($User in $Users) {
        [pscustomobject]@{
            Id                                           = $User.Id
            UserPrincipalName                            = $User.UserPrincipalName
            UserDisplayName                              = $User.UserDisplayName
            IsAdmin                                      = $User.IsAdmin
            DefaultMfaMethod                             = $User.DefaultMfaMethod
            MethodsRegistered                            = $User.MethodsRegistered -join ','
            IsMfaCapable                                 = $User.IsMfaCapable
            IsMfaRegistered                              = $User.IsMfaRegistered
            IsPasswordlessCapable                        = $User.IsPasswordlessCapable
            IsSsprCapable                                = $User.IsSsprCapable
            IsSsprEnabled                                = $User.IsSsprEnabled
            IsSsprRegistered                             = $User.IsSsprRegistered
            IsSystemPreferredAuthenticationMethodEnabled = $User.IsSystemPreferredAuthenticationMethodEnabled
            LastUpdatedDateTime                          = $User.LastUpdatedDateTime
        }
      }
      
      $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MFA-Status-Report.csv")
      $Report | Out-GridView -Title "Authentication Methods Report"
      $Report | Export-Csv -Path $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

      Write-InformationColored "Report exported successfully to $($CSV)" -ForegroundColor "Green"
    } catch {
      Write-Error "`nAn error occurred: $_"
    }
  } else {
    Write-Error "`nCan't connect or use Microsoft Graph Modules. `nPlease check logs."
  }

  Set-Variable InformationPreference $previousInformationPreference
  Set-Variable ProgressPreference $previousProgressPreference
}

function User-DisableDevices {
  # Credits:
  #   https://alitajran.com/force-sign-out-users-microsoft-365
  param (
    [Parameter(Position = 0, Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Specify the users you want to disable registered devices")]
    [string[]]$UserPrincipalNames
  )

  $mggConnectedCheck = priv_CheckMGGraphModule

  if ( $mggConnectedCheck -eq $true ) {    
    # Filter users based on provided user principal names
    if ($UserPrincipalNames) {
      $Users = $UserPrincipalNames | Foreach-Object { Get-MgUser -Filter "UserPrincipalName eq '$($_)'" }
    } else {
      $Users = @()
      Write-Host "`nNo -UserPrincipalNames or -All parameter provided." -f "Yellow"
    }

    # Check if any provided users were not found
    $UsersNotFound = $UserPrincipalNames | Where-Object { $Users.UserPrincipalName -notcontains $_ }
    foreach ($userNotFound in $UsersNotFound) {
      Write-Host "Can't find Azure AD account for user $userNotFound" -f "Red"
    }

    foreach ($User in $Users) {              
      # Retrieve (and disable) registered devices
      $UserDevices = Get-MgUserRegisteredDevice -UserId $User.Id
      if ($UserDevices) {
        foreach ($Device in $UserDevices) {
          Update-MgDevice -DeviceId $Device.Id -AccountEnabled $false
        }
      }
      Write-Host "Disable registered devices completed for $($User.DisplayName)" -f "Green"
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -f "Red"
  }

}

function User-DisableSignIn {
  # Credits:
  #   https://alitajran.com/force-sign-out-users-microsoft-365
  param (
    [Parameter(Position = 0, Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Specify the users you want to disable sign in")]
    [string[]]$UserPrincipalNames
  )

  $mggConnectedCheck = priv_CheckMGGraphModule

  if ( $mggConnectedCheck -eq $true ) {    
    # Filter users based on provided user principal names
    if ($UserPrincipalNames) {
      $Users = $UserPrincipalNames | Foreach-Object { Get-MgUser -Filter "UserPrincipalName eq '$($_)'" }
    } else {
      $Users = @()
      Write-Host "`nNo -UserPrincipalNames or -All parameter provided." -f "Yellow"
    }

    # Check if any provided users were not found
    $UsersNotFound = $UserPrincipalNames | Where-Object { $Users.UserPrincipalName -notcontains $_ }
    foreach ($userNotFound in $UsersNotFound) {
      Write-Host "Can't find Azure AD account for user $userNotFound" -f "Red"
    }

    foreach ($User in $Users) {              
      # Block sign-in
      Update-MgUser -UserId $User.Id -AccountEnabled:$False
      Write-Host "Disable sign-in completed for account $($User.DisplayName)" -f "Green"
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -f "Red"
  }

}

function User-SignOut {
  # Credits:
  #   https://alitajran.com/force-sign-out-users-microsoft-365
  param (
    [Parameter(Mandatory=$False, HelpMessage="Force sign out for all tenant users")]
    [switch]$All,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Specify all the users you want to exclude from the operation")]
    [string[]]$Exclude,
    [Parameter(Position = 0, Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Specify the users you want to force sign out")]
    [string[]]$UserPrincipalNames
  )

  $mggConnectedCheck = priv_CheckMGGraphModule

  if ( $mggConnectedCheck -eq $true ) {
    # Check if no switches or parameters are provided
    if (-not $All -and -not $Exclude -and -not $UserPrincipalNames) {
      Write-Host "`nNo switches or parameters provided. Please specify the desired action using switches such as -All or provide user principal names using -UserPrincipalNames." -f "Yellow"
      Exit
    }

    # Retrieve all users if -All parameter is specified
    if ($All) {
      $Users = Get-MgUser -All
    } else {
      # Filter users based on provided user principal names
      if ($UserPrincipalNames) {
        $Users = $UserPrincipalNames | Foreach-Object { Get-MgUser -Filter "UserPrincipalName eq '$($_)'" }
      } else {
        $Users = @()
        Write-Host "`nNo -UserPrincipalNames or -All parameter provided." -f "Yellow"
      }
    }

    # Check if any excluded users were not found
    $ExcludedNotFound = $Exclude | Where-Object { $Users.UserPrincipalName -notcontains $_ }
    foreach ($excludedUser in $ExcludedNotFound) {
      Write-Host "Can't find Azure AD account for user $excludedUser" -f "Red"
    }

    # Check if any provided users were not found
    $UsersNotFound = $UserPrincipalNames | Where-Object { $Users.UserPrincipalName -notcontains $_ }
    foreach ($userNotFound in $UsersNotFound) {
      Write-Host "Can't find Azure AD account for user $userNotFound" -f "Red"
    }

    foreach ($User in $Users) {
      # Check if the user should be excluded
      if ($Exclude -contains $User.UserPrincipalName) {
        Write-Host "Skipping user $($User.UserPrincipalName)" -f "Cyan"
        continue
      }
      
      # Revoke all signed in sessions and refresh tokens for the account
      $SignOutStatus = Revoke-MgUserSignInSession -UserId $User.Id
      Write-Host "Sign-out completed for account $($User.DisplayName)" -f "Green"
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -f "Red"
  }

}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Change-MFAStatus"
Export-ModuleMember -Function "Export-MFAStatus"
Export-ModuleMember -Function "User-DisableDevices"
Export-ModuleMember -Function "User-DisableSignIn"
Export-ModuleMember -Function "User-SignOut"