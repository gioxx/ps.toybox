# Tools for Gioxx's ToyBox

function priv_CheckEOLConnection {
    $eolConnected = $false
    $userConnected = priv_MailSearcher

    if ( (Get-Module -Name ExchangeOnlineManagement -ListAvailable).count -gt 0 ) {
        try {
            Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
            $eolConnected = $true
        } catch {
            if ( $userConnected -ne "notfound" ) {
                Write-Host "Please wait until I load Microsoft Exchange Online Management.`nConnecting using $($userConnected) ..." -f "Yellow"
                Connect-EOL -UserPrincipalName $userConnected 
            } else {
                Write-Host "Please wait until I load Microsoft Exchange Online Management." -f "Yellow"
                Connect-EOL 
            }
            $eolConnected = $true
        }
    } else {
        Write-Host "Microsoft Exchange Online Management module is not available."  -f "Yellow" 
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No "
        if ( $Confirm -match "[yY]" ) {
            try {
                Write-host "Installing Microsoft Exchange Online Management PowerShell module ..."
                Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber -Force
                if ( $userConnected -ne "notfound" ) { 
                    Write-Host "Please wait until I load Microsoft Exchange Online Management.`nConnecting using $($userConnected) ..." -f "Yellow"
                    Connect-EOL -UserPrincipalName $userConnected
                } else {
                    Connect-EOL
                }
                $eolConnected = $true
            } catch {
                ""; Write-Host "Can't install Exchange Online Management modules. `nPlease check logs." -f "Red"
                exit
            }
        } else {
            ""; Write-Host "Microsoft Exchange Online Management module is required to run this script. `nPlease install module using Install-Module ExchangeOnlineManagement cmdlet." -f "Red"
            exit
        }
    }

    return $eolConnected
}

function priv_CheckFolder($path) {
    if ([string]::IsNullOrEmpty($path)) {
        $path = ".\"
    } else {
        $path = $path.TrimEnd('\')
    }
    return $path
}

function priv_CheckMGGraphModule {
    $mggConnected = $false

    if ( (Get-Module -Name Microsoft.Graph -ListAvailable).count -gt 0 ) {
        try {
            Get-MgUser -ErrorAction Stop
            $mggConnected = $true
        } catch {
            Write-Host "Please wait until I load Microsoft Graph, the operation can take a minute or more." -f "Yellow"
            Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            Connect-MgGraph
            $mggConnected = $true
        }
    } else {
        Write-Host "Microsoft Graph PowerShell module is not available."  -f "Yellow" 
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No "
        if ( $Confirm -match "[yY]" ) {
            try {
                Write-host "Installing Microsoft Graph PowerShell module ..."
                Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
                Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
                Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
                Connect-MgGraph
                $mggConnected = $true
            } catch {
                ""; Write-Host "Can't install and import Graph modules. `nPlease check logs." -f "Red"
                exit
            }
        } else {
            ""; Write-Host "Microsoft Graph PowerShell module is required to run this script. `nPlease install module using Install-Module Microsoft.Graph cmdlet." -f "Red"
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
    $result = $form.ShowDialog()

    if ($textBox.Text -eq '') {
        # Empty TextBox
        Write-Host "Message can't be empty, operation canceled." -f "Yellow"
        break
    } else {
        if ( $result -eq [System.Windows.Forms.DialogResult]::OK ) {
            $x = $textBox.Lines | Where{$_} | ForEach{ $_.Trim() }
            $array = @()
            $array = $x -split "`r`n"
            $AbsenceMessageHTMLOutput = $array -join "<br>"
            #Return $array | Where-Object {$_ -ne ''}
            return $AbsenceMessageHTMLOutput.Trim()
        }

        if ( $result -eq [System.Windows.Forms.DialogResult]::Cancel ) {
            Write-Host "Operation canceled (Aborted by user)." -f "Yellow"
            break
        }
    }
}

function priv_MailSearcher {
    # Credits: https://powershellmagazine.com/2012/11/14/pstip-how-to-get-the-email-address-of-the-currently-logged-on-user/
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    if ( $searcher -ne $null ) {
        $mailAddress = $searcher.FindOne().Properties.mail
        return $mailAddress
    } else {
        return "notfound"
    }
}

function priv_MaxLenghtSubString($string, $maxchars) {
    if ($string.Length -gt $maxchars) { 
        return "$($string.substring(0, $maxchars))..."
    } else {
        return $string
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

function priv_TakeDecisionOptions($message, $yes, $no, $yesHint, $noHint, $defaultOption=0) {
    $option_1 = New-Object System.Management.Automation.Host.ChoiceDescription "$($yes)", "$($yesHint)"
    $option_2 = New-Object System.Management.Automation.Host.ChoiceDescription "$($no)", "$($noHint)"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($option_1, $option_2)
    $options_result = $Host.UI.PromptForChoice("", "`n$message", $options, $defaultOption)
    return $options_result
}