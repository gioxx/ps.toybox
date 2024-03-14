# Protection =======================================================================================================================================================

function DECOMMISSIONING_Export-MFAStatus {
  # Credits:
  #   https://activedirectorypro.com/mfa-status-powershell
  #   https://lazyadmin.nl
  #   https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Extract into CSV all users (even those with MFA disabled)")]
    [switch] $All
  )
  
  Set-Variable ProgressPreference Continue
  $folder = priv_CheckFolder($folderCSV)
  $mggConnectedCheck = priv_CheckMGGraphModule
  
  if ( $mggConnectedCheck -eq $true ) {
    $arr_MFAStatus = @()
    $ProcessedCount = 0

    $select = @(
      'id',
      'DisplayName',
      'userprincipalname',
      'mail'
    )
    $properties = $select + "AssignedLicenses"
    $filter = "UserType eq 'member'"

    $Users = Get-MgUser -Filter $filter -Property $properties -All | 
      Where { ($_.AssignedLicenses).count -gt 0 } | 
      Select-Object $select

    $Users | ForEach {
      $ProcessedCount++
      $PercentComplete = ( ($ProcessedCount / $totalUsers) * 100 )
      $User = $_
      Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
      
      $MFAMethods = [PSCustomObject][Ordered]@{
        status = ""
        authApp = ""
        phoneAuth = ""
        fido = ""
        helloForBusiness = ""
        emailAuth = ""
        tempPass = ""
        passwordLess = ""
        softwareAuth = ""
        authDevice = ""
        authPhoneNr = ""
        SSPREmail = ""
      }

      $MFAData = Get-MgUserAuthenticationMethod -UserId $User.UserPrincipalName -ErrorAction SilentlyContinue

      ForEach ( $method in $MFAData ) {
        Switch ( $method.AdditionalProperties["@odata.type"] ) {
          "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
            # Microsoft Authenticator App
            $MFAMethods.authApp = $True
            $MFAMethods.authDevice = $method.AdditionalProperties["displayName"] 
            $MFAMethods.status = "enabled"
          } 
          "#microsoft.graph.phoneAuthenticationMethod" { 
            # Phone authentication
            $MFAMethods.phoneAuth = $True
            $MFAMethods.authPhoneNr = $method.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
            $MFAMethods.status = "enabled"
          } 
          "#microsoft.graph.fido2AuthenticationMethod" { 
            # FIDO2 key
            $MFAMethods.fido = $True
            $fifoDetails = $method.AdditionalProperties["model"]
            $MFAMethods.status = "enabled"
          } 
          "#microsoft.graph.passwordAuthenticationMethod" { 
            # Password
            # When only the password is set, then MFA is disabled.
            if ($MFAMethods.status -ne "enabled") {$MFAMethods.status = "disabled"}
          }
          "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
            # Windows Hello
            $MFAMethods.helloForBusiness = $True
            $helloForBusinessDetails = $method.AdditionalProperties["displayName"]
            $MFAMethods.status = "enabled"
          } 
          "#microsoft.graph.emailAuthenticationMethod" { 
            # Email Authentication
            $MFAMethods.emailAuth =  $True
            $MFAMethods.SSPREmail = $method.AdditionalProperties["emailAddress"] 
            $MFAMethods.status = "enabled"
          }               
          "microsoft.graph.temporaryAccessPassAuthenticationMethod" { 
            # Temporary Access pass
            $MFAMethods.tempPass = $True
            $tempPassDetails = $method.AdditionalProperties["lifetimeInMinutes"]
            $MFAMethods.status = "enabled"
          }
          "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { 
            # Passwordless
            $MFAMethods.passwordLess = $True
            $passwordLessDetails = $method.AdditionalProperties["displayName"]
            $MFAMethods.status = "enabled"
          }
          "#microsoft.graph.softwareOathAuthenticationMethod" { 
            # ThirdPartyAuthenticator
            $MFAMethods.softwareAuth = $True
            $MFAMethods.status = "enabled"
          }
        }
      }

      if ( $All ) {
        $arr_MFAStatus += New-Object -TypeName PSObject -Property $([ordered]@{ 
          Name = $User.DisplayName
          "Email Address" = $User.mail
          UserPrincipalName = $User.UserPrincipalName
          "MFA Status" = $MFAMethods.status
        # "MFA Default type" = ""  - Not yet supported by MgGraph
          "Phone Authentication" = $MFAMethods.phoneAuth
          "Authenticator App" = $MFAMethods.authApp
          "Passwordless" = $MFAMethods.passwordLess
          "Hello for Business" = $MFAMethods.helloForBusiness
          "FIDO2 Security Key" = $MFAMethods.fido
          "Temporary Access Pass" = $MFAMethods.tempPass
          "Authenticator device" = $MFAMethods.authDevice
          "Phone number" = $MFAMethods.authPhoneNr
          "Recovery email" = $MFAMethods.SSPREmail
        })
      } else {
        if ( $MFAMethods.status -eq "enabled" ) {
          $arr_MFAStatus += New-Object -TypeName PSObject -Property $([ordered]@{ 
            Name = $User.DisplayName
            "Email Address" = $User.mail
            UserPrincipalName = $User.UserPrincipalName
          # "MFA Status" = $MFAMethods.status - It's unnecessary because in this case you filter out only those who have it active
          # "MFA Default type" = ""  - Not yet supported by MgGraph
            "Phone Authentication" = $MFAMethods.phoneAuth
            "Authenticator App" = $MFAMethods.authApp
            "Passwordless" = $MFAMethods.passwordLess
            "Hello for Business" = $MFAMethods.helloForBusiness
            "FIDO2 Security Key" = $MFAMethods.fido
            "Temporary Access Pass" = $MFAMethods.tempPass
            "Authenticator device" = $MFAMethods.authDevice
            "Phone number" = $MFAMethods.authPhoneNr
            "Recovery email" = $MFAMethods.SSPREmail
          }) 
        }
      }

    }

    $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MFA-Status-Report.csv")
    $arr_MFAStatus | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

  } else {
    Write-Host "`nCan't connect or use Microsoft Graph Modules. `nPlease check logs." -f "Red"
  }
}

function Export-MFAStatus {
  # Credits:
  #   https://www.alitajran.com/get-mfa-status-entra/
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV
  )
  
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

      Write-Host "Report exported successfully to $($CSV)" -f "Green"
    } catch {
      Write-Host "`nAn error occurred: $_" -f "Red"
    }
  } else {
    Write-Host "`nCan't connect or use Microsoft Graph Modules. `nPlease check logs." -f "Red"
  }
}

function DECOMMISSIONING_Export-MFAStatusDefaultMethod {
  # Credits: https://thesysadminchannel.com/get-per-user-mfa-status-using-powershell
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Extract into CSV all users (even those with MFA disabled)")]
    [switch] $All
  )

  Set-Variable ProgressPreference Continue
  $folder = priv_CheckFolder($folderCSV)

  if ( -not (Get-MsolDomain -ErrorAction SilentlyContinue) ) {
    Write-Error "You must connect to the MSolService to continue" -ErrorAction Stop
  }

  $arr_MFAStatusDefaultMethod = @()
  $ProcessedCount = 0
  $MsolUserList = Get-MsolUser -All -ErrorAction Stop | 
      Where { $_.UserType -ne 'Guest' -And $_.DisplayName -notmatch 'On-Premises Directory Synchronization' }
  $totalUsers = $MsolUserList.Count

  ForEach ( $User in $MsolUserList ) {
    $ProcessedCount++
    $PercentComplete = ( ($ProcessedCount / $totalUsers) * 100 )
    Write-Progress -Activity "Processing $User" -Status "$ProcessedCount out of $totalUsers completed ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    
    if ( $User.StrongAuthenticationRequirements ) {
      $PerUserMFAState = $User.StrongAuthenticationRequirements.State
    } else {
      $PerUserMFAState = 'Disabled'
    }

    $MethodType = $User.StrongAuthenticationMethods | 
        Where { $_.IsDefault -eq $True } | 
        Select-Object -ExpandProperty MethodType

    if ( $MethodType ) {
      switch ( $MethodType ) {
        'OneWaySMS' {$DefaultMethodType = 'SMS Text Message'}
        'TwoWayVoiceMobile' {$DefaultMethodType = 'Call to Phone'}
        'PhoneAppOTP' {$DefaultMethodType = 'TOTP'}
        'PhoneAppNotification' {$DefaultMethodType = 'Authenticator App'}
      }
    } else {
      $DefaultMethodType = 'Not Enabled'
    }

    if ( $All ) {
      $arr_MFAStatusDefaultMethod += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserPrincipalName = $User.UserPrincipalName
        DisplayName = $User.DisplayName
        PerUserMFAState = $PerUserMFAState
        DefaultMethodType = $DefaultMethodType
      })
      
      $MethodType = $null
    } else {
      if ( !($PerUserMFAState -eq 'Disabled') ) {
        $arr_MFAStatusDefaultMethod += New-Object -TypeName PSObject -Property $([ordered]@{ 
          UserPrincipalName = $User.UserPrincipalName
          DisplayName = $User.DisplayName
          DefaultMethodType = $DefaultMethodType
        })
        
        $MethodType = $null
      }
    }

  }

  $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MFA-DefaultAuthMethod-Report.csv")
  $arr_MFAStatusDefaultMethod | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"

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

# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Export-MFAStatus"
Export-ModuleMember -Function "User-DisableDevices"
Export-ModuleMember -Function "User-DisableSignIn"
Export-ModuleMember -Function "User-SignOut"