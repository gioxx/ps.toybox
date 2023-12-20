# Credits: https://petri.com/more-efficient-powershell-with-psreadline/
Set-PSReadlineKeyHandler -Key F7 -BriefDescription HistoryList -Description "Show command history with Out-Gridview. [$($env:username)]" -ScriptBlock {
    $pattern = $null
    [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$pattern, [ref]$null)
    if ($pattern)
    {
        $pattern = [regex]::Escape($pattern)
    }
    $history = [System.Collections.ArrayList]@(
        $last = ''
        $lines = ''
        foreach ($line in [System.IO.File]::ReadLines((Get-PSReadlineOption).HistorySavePath))
        {
            if ($line.EndsWith('`'))
            {
                $line = $line.Substring(0, $line.Length - 1)
                $lines = if ($lines)
                {
                    "$lines`n$line"
                }
                else
                {
                    $line
                }
                continue
            }
            if ($lines)
            {
                $line = "$lines`n$line"
                $lines = ''
            }
            if (($line -cne $last) -and (!$pattern -or ($line -match $pattern)))
            {
                $last = $line
                $line
            }
        }
    )
    $history.Reverse()
    $command = $history | Select-Object -unique | Out-GridView -Title "PSReadline History - Select a command to insert at the prompt" -OutputMode Single
    if ($command)
    {
        [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert(($command -join "`n"))
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

# EXPORT MODULES ===================================================================================================================================================

Export-ModuleMember -Function "User-CloseAllPSSessions"