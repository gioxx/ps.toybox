# Source: https://petri.com/more-efficient-powershell-with-psreadline/
Set-PSReadlineKeyHandler -Key F7 -BriefDescription HistoryList -Description "Show command history with Out-Gridview. [$($env:username)]" -ScriptBlock {
    $pattern = $null
    [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$pattern, [ref]$null)
    if ($pattern) {
        $pattern = [regex]::Escape($pattern)
    }
    $history = [System.Collections.ArrayList]@(
        $last = ''
        $lines = ''
        foreach ($line in [System.IO.File]::ReadLines((Get-PSReadlineOption).HistorySavePath)) {
            if ($line.EndsWith('`')) {
                $line = $line.Substring(0, $line.Length - 1)
                $lines = if ($lines) {
                    "$lines`n$line"
                } else {
                    $line
                }
                continue
            }
            if ($lines) {
                $line = "$lines`n$line"
                $lines = ''
            }
            if (($line -cne $last) -and (!$pattern -or ($line -match $pattern))) {
                $last = $line
                $line
            }
        }
    )
    $history.Reverse()
    $command = $history | Select-Object -unique | Out-GridView -Title "PSReadline History - Select a command to insert at the prompt" -OutputMode Single
    if ($command) {
        [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert(($command -join "`n"))
    }
}

# Credits
# https://www.sharepointdiary.com/2020/04/powershell-generate-random-password.html
Function Get-RandomPassword {
    #define parameters
    param([Parameter(ValueFromPipeline=$false)][ValidateRange(1,256)][int]$PasswordLength = 10)
 
    #ASCII Character set for Password
    $CharacterSet = @{
        Lowercase   = (97..122) | Get-Random -Count 10 | % {[char]$_}
        Uppercase   = (65..90)  | Get-Random -Count 10 | % {[char]$_}
        Numeric     = (48..57)  | Get-Random -Count 10 | % {[char]$_}
        SpecialChar = (33..47)+(58..64)+(91..96)+(123..126) | Get-Random -Count 10 | % {[char]$_}
    }
 
    #Frame Random Password from given character set
    $StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar
 
    -join(Get-Random -Count $PasswordLength -InputObject $StringSet)
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