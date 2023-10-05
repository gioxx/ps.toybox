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

function Export-QuarantineEML {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="The ID of the message to be exported (example: 20230617142935.F5B74194B266E458@contoso.com)")]
    [string]$messageID,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export the EML file (e.g. C:\Temp)")]
    [string] $folder
  )
  
  if ( -not($messageID.StartsWith('<')) ) { $messageID = '<' + $messageID }
  if ( -not($messageID.EndsWith('>')) ) { $messageID += '>' }
  
  $exportFolder = priv_CheckFolder($folder)
  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    if ( (Get-QuarantineMessage -MessageId $($messageID)).QuarantineTypes -eq "Malware" ) {
      $e = Get-QuarantineMessage -MessageId $($messageID) | Export-QuarantineMessage
      $bytes = [Convert]::FromBase64String($e.eml)
      [IO.File]::WriteAllBytes("$($exportFolder)\QuarantineEML.eml", $bytes)
    } else {
      $e = Get-QuarantineMessage -MessageId $($messageID) | Export-QuarantineMessage
      $txt = [System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml))
      [IO.File]::WriteAllText("$($exportFolder)\QuarantineEML.eml", $txt)
    }    

    Invoke-Item "$($exportFolder)\QuarantineEML.eml"
    Start-Sleep -s 3
    Remove-Item "$($exportFolder)\QuarantineEML.eml"
    
    $options_result = priv_TakeDecisionOptions "Should I release the message to all recipients?" "&Yes" "&No" "Release the message." "Do not release the message." 1
    if ($options_result -eq 0) {
      $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1
      if ( $reportFalsePositive -eq 0 ) {
        Get-QuarantineMessage -MessageId $($messageID) | Release-QuarantineMessage -ReleaseToAll -ReportFalsePositive -Confirm:$false
      } else { 
        Get-QuarantineMessage -MessageId $($messageID) | Release-QuarantineMessage -ReleaseToAll
      }
    } else {
      Write-Host "Message not released (aborted by user)." -f "Yellow"
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Get-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      ForEach ( $CurrentSender in $SenderAddress ) {
        try {
          Write-Host "Find e-mail(s) from known senders quarantined: e-mail(s) from $($SenderAddress) ..."
          Get-QuarantineMessage -SenderAddress $SenderAddress | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
    }
  }
}

function Get-QuarantineFromDomain {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail domain in quarantine (e.g. contoso.com)")]
    [string[]]$SenderDomain
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection

    if ( $eolConnectedCheck -eq $true ) {
      ForEach ( $CurrentSender in $SenderDomain ) {
        try {
          Write-Host "Find e-mail(s) from known domains quarantined: e-mail(s) from $($SenderDomain) ..."
          Get-QuarantineMessage | Where-Object { $_.SenderAddress -like "*@$($SenderDomain)" } | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Format-Table -AutoSize Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
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

  if ( $Interval -gt 30 ) { $Interval = 30 } else { $Interval = $($Interval) }
  $arr_QuarantineToRelease = @()
  $ReleaseQuarantinePreview = @()
  $ReleaseQuarantineReleased = @()
  $ReleaseQuarantineDeleted = @()
  $Page = 1
  
  $startDate = (Get-Date).AddDays(-$Interval)
  $endDate = Get-Date

  $eolConnectedCheck = priv_CheckEOLConnection

  if ( $eolConnectedCheck -eq $true ) {
    Write-Host "Quarantine report from $($startDate.Date) to $($endDate)" -f "Yellow"
    
    do {
      # Credits: https://community.spiceworks.com/topic/2343368-merge-eop-quarantine-pages#entry-9354845
      $QuarantinedMessages = Get-QuarantineMessage -StartReceivedDate $startDate.Date -EndReceivedDate $endDate -PageSize 1000 -ReleaseStatus NotReleased -Page $Page
      $Page++
      $QuarantinedMessagesAll += $QuarantinedMessages
    } until ( $QuarantinedMessages -eq $null )

    Write-Host "Total items: $($QuarantinedMessagesAll.Count)" -f "Yellow"

    $QuarantinedMessagesAll | ForEach {
      $Message = $_
      $arr_QuarantineToRelease += New-Object -TypeName PSObject -Property $([ordered]@{
        SenderAddress = $Message.SenderAddress
        RecipientAddress = $Message.RecipientAddress
        Subject = $Message.Subject
        ReceivedTime = $Message.ReceivedTime
        QuarantineTypes = $Message.QuarantineTypes
        Released = $Message.Released
        MessageId = $Message.MessageId
        Identity = $Message.Identity
      })
    }

    if ( $GridView ) {
      # Credits: https://stackoverflow.com/a/51033908
      $ReleaseQuarantine = $arr_QuarantineToRelease | Sort-Object -Descending ReceivedTime | Out-GridView -Title "$($startDate.Date) to $($endDate) • $($Interval) days • $($QuarantinedMessagesAll.Count) items" -PassThru

      $ProcessedCount = 0
      
      if ( $ReleaseQuarantine -ne $null ) {
        if ( $ReleaseQuarantine.Count -eq 1 ) {
          $single_menu = priv_TakeDecisionOptions "What do you want to do with $($ReleaseQuarantine.Subject)?" "&Release" "&Analyze" "Release the message." "Export a copy of the EML file for analysis."

          Switch ( $single_menu ) {
            0 { 
                $decision = priv_TakeDecisionOptions "Do you really want to release $($ReleaseQuarantine.Subject)?" "&Yes" "&No" "Release the message." "Do not release message."
                if ($decision -eq 0) {
                  
                  $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1
                  if ( $reportFalsePositive -eq 0 ) {
                    Release-QuarantineMessage -Identity $ReleaseQuarantine.Identity -ReleaseToAll -ReportFalsePositive -Confirm:$false
                  } else { 
                    Release-QuarantineMessage -Identity $ReleaseQuarantine.Identity -ReleaseToAll -Confirm:$false
                  }
                  
                  $released = Get-QuarantineMessage -Identity $ReleaseQuarantine.Identity
                  
                  $releasedResults = @()
                  $releasedResults += New-Object -TypeName PSObject -Property $([ordered]@{
                    Subject = priv_MaxLenghtSubString $released.Subject 40
                    SenderAddress = priv_MaxLenghtSubString $released.SenderAddress $MaxFieldLength
                    Released = $released.Released
                    ReleasedUser = $released.ReleasedUser
                  })
                  $releasedResults | Sort-Object Subject | Out-Host
                }
              }
            1 { Export-QuarantineEML -messageID "$($ReleaseQuarantine.MessageId)" }
          }
        } else {
          $ReleaseQuarantine | ForEach {
            $QuarantinedMessage = $_
            $ReleaseQuarantinePreview += New-Object -TypeName PSObject -Property $([ordered]@{
              Subject = priv_MaxLenghtSubString $QuarantinedMessage.Subject 50
              SenderAddress = priv_MaxLenghtSubString $QuarantinedMessage.SenderAddress $MaxFieldLength
              Released = $QuarantinedMessage.Released
            })
          }
          
          ""; Write-Host "$($ReleaseQuarantine.Count) items selected, take a look at the preview below:" -f "Cyan"
          $ReleaseQuarantinePreview | Sort-Object Subject | Select-Object Subject,SenderAddress,Released | Out-Host

          $release_or_delete = priv_TakeDecisionOptions "Do you want to release or delete $($ReleaseQuarantine.Count) selected items?" "&Release" "&Delete" "Release messages" "Delete messages"
          
          if ( $release_or_delete -eq 1 ) {
            # DELETE QUARANTINED EMAILS SELECTED
            $decision = priv_TakeDecisionOptions "Do you really want to permanently delete $($ReleaseQuarantine.Count) selected items?" "&Yes" "&No" "Delete message(s)." "Do not delete message(s)."
            $ReleaseQuarantine | ForEach {
              if ($decision -eq 0) {
                $QuarantinedMessageToDelete = Get-QuarantineMessage -Identity $_.Identity
                
                $ProcessedCount++
                $PercentComplete = ( ($ProcessedCount / $ReleaseQuarantine.Count) * 100 )
                Write-Progress -Activity "Deleting $(priv_MaxLenghtSubString $QuarantinedMessageToDelete.Subject $MaxFieldLength)" -Status "$ProcessedCount out of $($ReleaseQuarantine.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

                $QuarantinedMessageToDelete | Delete-QuarantineMessage -Confirm:$false
              }
            }
            Write-Host "Done." -f "Green"
          } else {
            # RELEASE QUARANTINED EMAILS SELECTED
            $decision = priv_TakeDecisionOptions "Do you really want to release $($ReleaseQuarantine.Count) selected items?" "&Yes" "&No" "Release messages." "Do not release messages."
            $reportFalsePositive = priv_TakeDecisionOptions "Do you want to report false positive to Microsoft?" "&Yes" "&No" "Report false positive message to Microsoft." "Do not report false positive message to Microsoft." 1

            $ReleaseQuarantine | ForEach {
              if ( $decision -eq 0 ) {
                if ( $reportFalsePositive -eq 0 ) {
                  Release-QuarantineMessage -Identity $_.Identity -ReleaseToAll -ReportFalsePositive -Confirm:$false
                } else { 
                  Release-QuarantineMessage -Identity $_.Identity -ReleaseToAll -Confirm:$false
                }
                
                $QuarantinedMessageReleased = Get-QuarantineMessage -Identity $_.Identity
                
                $ProcessedCount++
                $PercentComplete = (($ProcessedCount / $ReleaseQuarantine.Count) * 100)
                Write-Progress -Activity "Processing $(priv_MaxLenghtSubString $QuarantinedMessageReleased.Subject $MaxFieldLength)" -Status "$ProcessedCount out of $($ReleaseQuarantine.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
                
                $ReleaseQuarantineReleased += New-Object -TypeName PSObject -Property $([ordered]@{
                  Subject = priv_MaxLenghtSubString $QuarantinedMessageReleased.Subject $MaxFieldLength
                  SenderAddress = priv_MaxLenghtSubString $QuarantinedMessageReleased.SenderAddress $MaxFieldLength
                  Released = $QuarantinedMessageReleased.Released
                  ReleasedUser = $QuarantinedMessageReleased.ReleasedUser
                })
              }
            }
            Write-Host "Done, please take a look below." -f "Green"
            $ReleaseQuarantineReleased | Sort-Object Subject | Select-Object Subject,SenderAddress,Released,ReleasedUser | Out-Host
          }
        }
      }
    } else {
      $arr_QuarantineToRelease | Sort-Object Subject | Select-Object SenderAddress,RecipientAddress,Subject,QuarantineTypes,Released
    }

  } else {
    Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
  }
}

function Release-QuarantineFrom {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="Sender's e-mail address locked in quarantine (e.g. mario.rossi@contoso.com)")]
    [string[]]$SenderAddress
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    if ( $eolConnectedCheck -eq $true ) {

      $releasedResults = @()
      $SenderAddress | ForEach {
        try {
          $CurrentSender = $_
          Write-Host "Release quarantine from known senders: release e-mail(s) from $($CurrentSender) ..."
          Get-QuarantineMessage -SenderAddress $CurrentSender | 
              ForEach { Get-QuarantineMessage -Identity $_.Identity } | 
              Where-Object { $null -ne $_.QuarantinedUser -and $_.ReleaseStatus -ne "RELEASED" } | 
              Release-QuarantineMessage -ReleaseToAll
          Get-QuarantineMessage -SenderAddress $CurrentSender | 
            ForEach { 
              $released = Get-QuarantineMessage -Identity $_.Identity
              $row = "" | Select-Object Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
              $row.Subject = priv_MaxLenghtSubString $released.Subject 40
              $row.SenderAddress = $released.SenderAddress
              $row.ReceivedTime = $released.ReceivedTime
              $row.Released = $released.Released
              $row.ReleasedUser = $released.ReleasedUser
              $releasedResults += $row
            } 
          $releasedResults | Format-Table -AutoSize
        } catch {
          Write-Error $_.Exception.Message
        }
      }
    
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
    }
  }
}

function Release-QuarantineMessageId {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="ID of the message locked in quarantine (e.g. CAH_w85uSio_cz4HsFxJAGQDd-kzxGijLaMagZU95m3A1G8hWBA@mail.contoso.com)")]
    [string[]]$MessageId
  )

  process {
    $eolConnectedCheck = priv_CheckEOLConnection
    
    if ( $eolConnectedCheck -eq $true ) {
      $releasedResults = @()
      $MessageId | ForEach {
        try {
          $CurrentMessage = "<$($_)>"
          $ReleaseId = Get-QuarantineMessage -MessageId $CurrentMessage | 
              Where-Object { $null -ne $_.QuarantinedUser -and $_.ReleaseStatus -ne "RELEASED" }
          if ( $ReleaseId.Count -ge 1) {
            $ReleaseId | Release-QuarantineMessage -ReleaseToAll
            Write-Host "Release quarantine message with id $($CurrentMessage) ..."
            Get-QuarantineMessage -MessageId $CurrentMessage | 
              ForEach { 
                $released = Get-QuarantineMessage -Identity $_.Identity
                $row = "" | Select-Object Subject,SenderAddress,ReceivedTime,Released,ReleasedUser
                $row.Subject = priv_MaxLenghtSubString $released.Subject 40
                $row.SenderAddress = $released.SenderAddress
                $row.ReceivedTime = $released.ReceivedTime
                $row.Released = $released.Released
                $row.ReleasedUser = $released.ReleasedUser
                $releasedResults += $row
              }
            $releasedResults | Format-Table -AutoSize
          } else {
            Write-Host "No quarantined messages to release with id $($CurrentMessage) (already released or not found yet)." -f "Yellow"
          }
        } catch {
          Write-Error $_.Exception.Message
        }
      }
      
    } else {
      Write-Host "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -f "Red"
    }
  }
}

# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Export-MFAStatus"
#Export-ModuleMember -Function "Export-MFAStatusDefaultMethod"
Export-ModuleMember -Function "Export-QuarantineEML"
Export-ModuleMember -Function "Get-QuarantineFrom"
Export-ModuleMember -Function "Get-QuarantineFromDomain"
Export-ModuleMember -Function "Get-QuarantineToRelease"
Export-ModuleMember -Function "Release-QuarantineFrom"
Export-ModuleMember -Function "Release-QuarantineMessageId"