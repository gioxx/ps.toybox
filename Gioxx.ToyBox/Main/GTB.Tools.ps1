# Tools for Gioxx's ToyBox

function priv_CheckEOLConnection {
    $eolConnected = $false

    if ( (Get-Module -Name ExchangeOnlineManagement -ListAvailable).count -gt 0 ) {
        try {
            Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
            $eolConnected = $true
        } catch {
            $userConnected = priv_MailSearcher
            if ( $userConnected -ne "notfound" ) {
                Write-InformationColored "Please wait until I load Microsoft Exchange Online Management.`nConnecting using $($userConnected) ..." -ForegroundColor Yellow
                Connect-EOL -UserPrincipalName $userConnected
            } else {
                Write-InformationColored "Please wait until I load Microsoft Exchange Online Management." -ForegroundColor Yellow
                Connect-EOL
            }
            $eolConnected = $true
        }
    } else {
        Write-Warning "Microsoft Exchange Online Management module is not available."
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No "
        if ( $Confirm -match "[yY]" ) {
            try {
                Write-InformationColored "Installing Microsoft Exchange Online Management PowerShell module ..." -ForegroundColor Yellow
                Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber -Force
                if ( $userConnected -ne "notfound" ) { 
                    Write-InformationColored "Please wait until I load Microsoft Exchange Online Management.`nConnecting using $($userConnected) ..." -ForegroundColor Yellow
                    Connect-EOL -UserPrincipalName $userConnected
                } else {
                    Connect-EOL
                }
                $eolConnected = $true
            } catch {
                Write-Error "`nCan't install Exchange Online Management modules. `nPlease check logs."
                exit
            }
        } else {
            Write-Error "`nMicrosoft Exchange Online Management module is required to run this script. `nPlease install module using Install-Module ExchangeOnlineManagement cmdlet."
            exit
        }
    }

    return $eolConnected
}

function priv_CheckFolder($path) {
    if ([string]::IsNullOrEmpty($path)) {
        $path = $PWD
    } else {
        $path = $path.TrimEnd('\')
    }
    return $path
}

function priv_CheckMGGraphModule {
    $mggConnected = $false
    priv_CheckEOLConnection

    if ( (Get-Module -Name Microsoft.Graph -ListAvailable).count -gt 0 ) {
        try {
            Get-MgUser -ErrorAction Stop
            $mggConnected = $true
        } catch {
            Write-InformationColored "Please wait until I load Microsoft Graph, the operation may take a minute or more." -ForegroundColor Yellow
            # Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
            # Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            Connect-MgGraph
            $mggConnected = $true
        }
    } else {
        Write-Warning "Microsoft Graph PowerShell module is not available."
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No "
        if ( $Confirm -match "[yY]" ) {
            try {
                Write-InformationColored "Installing Microsoft Graph PowerShell module ..." -ForegroundColor Yellow
                Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
                # Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
                # Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
                Connect-MgGraph
                $mggConnected = $true
            } catch {
                Write-Error "`nCan't install and import Graph modules. `nPlease check logs."
                exit
            }
        } else {
            Write-Error "`nMicrosoft Graph PowerShell module is required to run this script. `nPlease install module using Install-Module Microsoft.Graph cmdlet."
            exit
        }
    }

    return $mggConnected
}

function priv_GUI_TextBox ($headerMessage, $defaultText) {
    # Credits: https://github.com/n2501r/spiderzebra/blob/master/PowerShell/GUI_Text_Box.ps1

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'OoO Message' 
    $form.Size = New-Object System.Drawing.Size(600,400)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $form.Topmost = $True

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(90,320)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(10,320)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.AutoSize = $True
    $label.Text = $headerMessage
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(560,275)
    $textBox.Multiline = $True
    $textbox.AcceptsReturn = $True
    $textBox.ScrollBars = "Vertical"
    $textBox.Text = $defaultText
    $form.Controls.Add($textBox)

    $form.Add_Shown({$textBox.Select()})
    $ShowDialogResult = $form.ShowDialog()

    if ($textBox.Text -eq '') {
        # Empty TextBox
        Write-Error "Message can't be empty, operation canceled."
        break
    } else {
        if ( $ShowDialogResult -eq [System.Windows.Forms.DialogResult]::OK ) {
            $x = $textBox.Lines | Where{$_} | ForEach{ $_.Trim() }
            $ShowDialogResult_array = @()
            $ShowDialogResult_array = $x -split "`r`n"
            $AbsenceMessageHTMLOutput = $ShowDialogResult_array -join "<br>"
            #Return $ShowDialogResult_array | Where-Object {$_ -ne ''}
            return $AbsenceMessageHTMLOutput.Trim()
        }

        if ( $ShowDialogResult -eq [System.Windows.Forms.DialogResult]::Cancel ) {
            Write-Error "Operation canceled (Aborted by user)."
            break
        }
    }
}

function priv_HideWarning {
    $warningPrefBackup = $WarningPreference
    $WarningPreference = "SilentlyContinue"
    return $warningPrefBackup
}

function priv_MailSearcher {
    # Credits: https://powershellmagazine.com/2012/11/14/pstip-how-to-get-the-email-address-of-the-currently-logged-on-user/
    # $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    # if ( $searcher -ne $null ) {
    #     $mailAddress = $searcher.FindOne().Properties.mail
    #     return $mailAddress
    # } else {
    #     return "notfound"
    # }

    try {
        $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
        if ($searcher -ne $null) {
            $mailAddress = $searcher.FindOne().Properties.mail
            if ($mailAddress) {
                return $mailAddress
            } else {
                return "notfound"
            }
        } else {
            return "notfound"
        }
    } catch {
        Write-Warning "Unable to automatically find e-mail address for current user."
        $mailAddress = Read-Host "Specify your e-mail address"
        return $mailAddress
    }

}

function priv_MaxLenghtSubString($string, $maxchars) {
    if ($string.Length -gt $maxchars) { 
        return "$($string.substring(0, $maxchars))..."
    } else {
        return $string
    }
}

function priv_RestorePreferences {
    if ($global:PreviousInformationPreference -ne $null) {
        Set-Variable -Name InformationPreference -Value $global:PreviousInformationPreference -Scope Global
    }

    if ($global:PreviousProgressPreference -ne $null) {
        Set-Variable -Name ProgressPreference -Value $global:PreviousProgressPreference -Scope Global
    }
}

function priv_SaveFileWithProgressiveNumber($path) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $extension = [System.IO.Path]::GetExtension($path)
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $count = 1
    while (Test-Path $path)
    {
        $fileName = $baseName + "_$count" + $extension
        $path = Join-Path -Path $directory -ChildPath $fileName
        $count++
    }
    return $path
}

function priv_SetPreferences {
    param (
        [switch]$Verbose
    )
    $global:PreviousInformationPreference = $InformationPreference
    $global:PreviousProgressPreference = $ProgressPreference
    
    if ($Verbose) {
        Set-Variable -Name InformationPreference -Value Continue -Scope Global
        Set-Variable -Name ProgressPreference -Value Continue -Scope Global
    }
}

function priv_TakeDecisionOptions($message, $yes, $no, $yesHint, $noHint, $defaultOption=0) {
    $option_1 = New-Object System.Management.Automation.Host.ChoiceDescription "$($yes)", "$($yesHint)"
    $option_2 = New-Object System.Management.Automation.Host.ChoiceDescription "$($no)", "$($noHint)"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($option_1, $option_2)
    $options_result = $Host.UI.PromptForChoice("", "`n$message", $options, $defaultOption)
    return $options_result
}