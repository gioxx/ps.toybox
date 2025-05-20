# Original source: https://petri.com/more-efficient-powershell-with-psreadline/
Set-PSReadLineKeyHandler -Key F7 -BriefDescription HistoryList -Description "Shows history and allows searching" -ScriptBlock {
    $historyPath = (Get-PSReadlineOption).HistorySavePath
    if (-not (Test-Path $historyPath)) {
        return
    }

    # Reads history, removes consecutive duplicates and shows only rows with text
    $history = Get-Content $historyPath | Where-Object { $_ -match '\S' } | Get-Unique

    if ($history) {
        # Asks the user to enter a search term and filters the history
        $searchTerm = Read-Host "Enter a keyword to filter the history (leave blank to show everything)"
        if ($searchTerm) {
            $filteredHistory = $history | Where-Object { $_ -match [regex]::Escape($searchTerm) }
        } else {
            $filteredHistory = $history
        }

        if (-not $filteredHistory -or $filteredHistory.Count -eq 0) {
            Write-Host "No command found with '$searchTerm'."
            return
        }

        # Shows the filtered history and asks the user to choose a command
        $indexedHistory = $filteredHistory | ForEach-Object -Begin { $i = 1 } -Process { "$i. $_"; $i++ }
        $indexedHistory | Out-Host
        $selectionIndex = Read-Host "Choose a command to execute (1-$($filteredHistory.Count))"

        # Checks whether the input is valid
        if ($selectionIndex -match '^\d+$' -and [int]$selectionIndex -gt 0 -and [int]$selectionIndex -le $filteredHistory.Count) {
            $command = $filteredHistory[[int]$selectionIndex - 1]
            [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
            [Microsoft.PowerShell.PSConsoleReadLine]::Insert($command)
        } else {
            Write-Host "Invalid selection."
        }
    }
}

# Credits
# https://www.sharepointdiary.com/2020/04/powershell-generate-random-password.html
# Modified to include a "less complex" version upon request.
# Utilization examples:
# Get-RandomPassword -PasswordLength 12                          -> One complex password
# Get-RandomPassword -PasswordLength 12 -Simple                  -> One simplified password
# Get-RandomPassword -PasswordLength 12 -Count 5                 -> Five complex passwords
# Get-RandomPassword -PasswordLength 12 -Simple -Count 3         -> Three simplified passwords
# Get-RandomPassword -PasswordLength 12 -Count 5 | Set-Clipboard -> Five complex passwords copied to clipboard
Function Get-RandomPassword {
    param(
        [Parameter(ValueFromPipeline = $false)]
        [ValidateRange(1, 256)]
        [int]$PasswordLength = 10,
        [switch]$Simple,
        [int]$Count = 1
    )

    if ($Simple) {
        $AllowedSpecialChars = [char[]]'!,.@#$_-'
    } else {
        $AllowedSpecialChars = [char[]](33..47 + 58..64 + 91..96 + 123..126)
    }

    $CharacterSet = @{
        Lowercase   = (97..122) | Get-Random -Count 10 | % { [char]$_ }
        Uppercase   = (65..90)  | Get-Random -Count 10 | % { [char]$_ }
        Numeric     = (48..57)  | Get-Random -Count 10 | % { [char]$_ }
        SpecialChar = ($AllowedSpecialChars) | Get-Random -Count 10
    }

    $StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar

    for ($i = 0; $i -lt $Count; $i++) {
        -join (Get-Random -Count $PasswordLength -InputObject $StringSet)
    }
}

# Credits
# https://www.reddit.com/r/PowerShell/comments/160185w/assembly_with_same_name_is_already_loaded/
# https://reddit.com/r/PowerShell/comments/160185w/comment/k69dgyn
# [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object Location | Sort-Object -Property FullName | Select-Object -Property FullName, Location, GlobalAssemblyCache, IsFullyTrusted
function User-CloseAllPSSessions {
    param()
    $PSSessions = Get-PSSession
    if ($PSSessions) {
        ForEach ($Session in $PSSessions) { 
            Remove-PSSession $Session.Id 
        }
    }
    Disconnect-ExchangeOnline -Confirm:$false
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Get-RandomPassword"
Export-ModuleMember -Function "User-CloseAllPSSessions"