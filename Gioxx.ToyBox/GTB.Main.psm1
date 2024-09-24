# DATA =============================================================================================================================================================
$GTB = [ordered]@{
  LicenseFilePath = 'JSON/M365_licenses.json'
  RepoName = 'ps.toybox'
  RepoOwner = 'gioxx'
}
New-Variable -Name GTBVars -Value $GTB -Scope Script -Force # Lista licenze M365 utilizzata in Export-MsolAccountSku

# FUNCTIONS ========================================================================================================================================================
# function GTBDebug {
#   priv_MailSearcher
# }

function Update-PS7 {
  iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI" # Aggiornamento PowerShell 7
}

# VARS =============================================================================================================================================================
New-Variable -Name MaxFieldLength -Value 35 -Scope Script -Force # Quantità caratteri richiamata / usata in priv_MaxLenghtSubString

# EXPORT MODULES ===================================================================================================================================================

# Export-ModuleMember -Function "GTBDebug"
Export-ModuleMember -Function "Update-PS7"