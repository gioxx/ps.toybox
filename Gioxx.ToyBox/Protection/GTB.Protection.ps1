# Protection =======================================================================================================================================================

function Export-MFAStatus {
  # Credits:
  #   https://activedirectorypro.com/mfa-status-powershell
  #   https://lazyadmin.nl
  #   https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Extract into CSV all users (even those with MFA disabled).")]
    [switch] $All
  )
  
  Set-Variable ProgressPreference Continue
  $folder = priv_CheckFolder($folderCSV)
  priv_CheckMGGraphModule
  $Result = @()
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
      Where {($_.AssignedLicenses).count -gt 0} | 
      Select-Object $select

  $Users | ForEach {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
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

    ForEach ($method in $MFAData) {
      Switch ($method.AdditionalProperties["@odata.type"]) {
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

    if ($All) {
      $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
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
      if ($MFAMethods.status -eq "enabled") {
        $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
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
  $Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

function Export-MFAStatusDefaultMethod {
  # Credits: https://thesysadminchannel.com/get-per-user-mfa-status-using-powershell
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Extract into CSV all users (even those with MFA disabled).")]
    [switch] $All
  )

  Set-Variable ProgressPreference Continue
  $folder = priv_CheckFolder($folderCSV)

  if (-not (Get-MsolDomain -ErrorAction SilentlyContinue)) {
    Write-Error "You must connect to the MSolService to continue" -ErrorAction Stop
  }

  $Result = @()
  $ProcessedCount = 0
  $MsolUserList = Get-MsolUser -All -ErrorAction Stop | 
      Where {$_.UserType -ne 'Guest' -And $_.DisplayName -notmatch 'On-Premises Directory Synchronization'}
  $totalUsers = $MsolUserList.Count

  ForEach ($User in $MsolUserList) {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
    Write-Progress -Activity "Processing $User" -Status "$ProcessedCount out of $totalUsers completed ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    
    if ($User.StrongAuthenticationRequirements) {
      $PerUserMFAState = $User.StrongAuthenticationRequirements.State
    } else {
      $PerUserMFAState = 'Disabled'
    }

    $MethodType = $User.StrongAuthenticationMethods | 
        Where {$_.IsDefault -eq $True} | 
        Select-Object -ExpandProperty MethodType

    if ($MethodType) {
      switch ($MethodType) {
        'OneWaySMS' {$DefaultMethodType = 'SMS Text Message'}
        'TwoWayVoiceMobile' {$DefaultMethodType = 'Call to Phone'}
        'PhoneAppOTP' {$DefaultMethodType = 'TOTP'}
        'PhoneAppNotification' {$DefaultMethodType = 'Authenticator App'}
      }
    } else {
      $DefaultMethodType = 'Not Enabled'
    }

    if ($All) {
      $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserPrincipalName = $User.UserPrincipalName
        DisplayName = $User.DisplayName
        PerUserMFAState = $PerUserMFAState
        DefaultMethodType = $DefaultMethodType
      })
      
      $MethodType = $null
    } else {
      if (!($PerUserMFAState -eq 'Disabled')) {
        $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
          UserPrincipalName = $User.UserPrincipalName
          DisplayName = $User.DisplayName
          DefaultMethodType = $DefaultMethodType
        })
        
        $MethodType = $null
      }
    }

  }
  $CSV = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-MFA-DefaultAuthMethod-Report.csv")
  $Result | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

function Export-QuarantineEML {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export the EML file (e.g. C:\Temp)")]
    [string] $folder,
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="The ID of the message to be exported (example: 20230617142935.F5B74194B266E458@contoso.com)")]
    [string]$messageID
  )
  
  if (-not($messageID.StartsWith('<'))) { $messageID = '<' + $messageID }
  if (-not($messageID.EndsWith('>'))) { $messageID += '>' }
  $exportFolder = priv_CheckFolder($folder)

  $e = Get-QuarantineMessage -MessageId $($messageID) | 
      Export-QuarantineMessage; $bytes = [Convert]::FromBase64String($e.eml); [IO.File]::WriteAllBytes("$($exportFolder)\QuarantineEML.eml", $bytes)

  Invoke-Item "$($exportFolder)\QuarantineEML.eml"
  Start-Sleep -s 3
  Remove-Item "$($exportFolder)\QuarantineEML.eml"
  
  $message = "Should I release the message to all recipients?"
  $option_y = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Release message."
  $option_n = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not release the message."
  $options = [System.Management.Automation.Host.ChoiceDescription[]]($option_y, $option_n)
  $options_result = $host.ui.PromptForChoice("", $message, $options, 0)
  if ($options_result -eq 0) {
    Get-QuarantineMessage -MessageId $($messageID) | 
        Release-QuarantineMessage -ReleaseToAll
  }
}

function Get-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    ForEach ($CurrentSender in $SenderAddress) {
      try {
        Write-Host "Find e-mail(s) from known senders quarantined: e-mail(s) from $($SenderAddress) not yet released ..."
        Get-QuarantineMessage -SenderAddress $SenderAddress | 
            ForEach {Get-QuarantineMessage -Identity $_.Identity} | 
            Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
      } catch {
        Write-Error $_.Exception.Message
      }
    }
  }
}

function Get-QuarantineToRelease {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Number of days to be analyzed from today (maximum 30)")]
    [ValidateNotNullOrEmpty()]
    [int]$Interval,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
    [switch] $GridView
  )

  Set-Variable ProgressPreference Continue

  if ($Interval -gt 30) {
    # Credits: https://stackoverflow.com/a/18311838
    $Interval = 30
  } else {
    $Interval = -$($Interval)
  }

  $Result = @()
  $ProcessedCount = 0
  $Page = 1
  
  $startDate = (Get-Date).AddDays($interval)
  $endDate = Get-Date
  Write-Host "Quarantine analysis with Start Date: $($startDate.Date) • End Date: $($endDate)" -f "Yellow"
  
  do {
    # Credits: https://community.spiceworks.com/topic/2343368-merge-eop-quarantine-pages#entry-9354845
    $QuarantinedMessages = Get-QuarantineMessage -StartReceivedDate $startDate.Date -EndReceivedDate $endDate -PageSize 1000 -ReleaseStatus NotReleased -Page $Page
    $Page++
    $QuarantinedMessagesAll += $QuarantinedMessages
  } until ( $QuarantinedMessages -eq $null )

  Write-Host "Total items: $($QuarantinedMessagesAll.Count)" -f "Yellow"

  $QuarantinedMessagesAll | ForEach {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $($QuarantinedMessagesAll.Count)) * 100)
    $Message = $_
    Write-Progress -Activity "Processing $($Message.Subject)" -Status "$ProcessedCount out of $($QuarantinedMessagesAll.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    $Result += New-Object -TypeName PSObject -Property $([ordered]@{
      SenderAddress = $Message.SenderAddress
      Subject = $Message.Subject
      QuarantineTypes = $Message.QuarantineTypes
      Released = $Message.Released
    })
  }
  if ( $GridView ) { $Result | Out-GridView } else { $Result }
}

function Release-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    $SenderAddress | ForEach {
      try {
        $CurrentSender = $_
        Write-Host "Release quarantine from known senders: release e-mail(s) from $($CurrentSender) ..."
        Get-QuarantineMessage -SenderAddress $CurrentSender | 
            ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
            Where-Object { $null -ne $_.QuarantinedUser -and $_.ReleaseStatus -ne "RELEASED" } | 
            Release-QuarantineMessage -ReleaseToAll
        Write-Host "Release quarantine from known senders: verifying e-mail(s) from $($CurrentSender) just released ..."
        Get-QuarantineMessage -SenderAddress $CurrentSender | 
            ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
            Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
      } catch {
        Write-Error $_.Exception.Message
      }
    }
  }
}

# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Export-MFAStatus"
#Export-ModuleMember -Function "Export-MFAStatusDefaultMethod"
Export-ModuleMember -Function "Export-QuarantineEML"
Export-ModuleMember -Function "Get-QuarantineFrom"
Export-ModuleMember -Function "Get-QuarantineToRelease"
Export-ModuleMember -Function "Release-QuarantineFrom"