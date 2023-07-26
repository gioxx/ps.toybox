# Vars =============================================================================================================================================================

$GTB = [ordered]@{
  LicensesJSON = 'https://raw.githubusercontent.com/gioxx/ps.toybox/main/JSON/M365_licenses.json'
}
New-Variable -Name GTBVars -Value $GTB -Scope Script -Force


function Update-PS7 {
  iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
}

# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Update-PS7"