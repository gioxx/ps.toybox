# FUNCTIONS ========================================================================================================================================================

function Update-PS7 {
  iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
}


# VARS =============================================================================================================================================================

  # Lista licenze M365 utilizzata in Export-MsolAccountSku
    $GTB = [ordered]@{
      LicensesJSON = 'https://raw.githubusercontent.com/gioxx/ps.toybox/main/JSON/M365_licenses.json'
    }
    New-Variable -Name GTBVars -Value $GTB -Scope Script -Force

  # Quantità caratteri richiamata / usata in priv_MaxLenghtSubString
    New-Variable -Name MaxFieldLength -Value 35 -Scope Script -Force


# EXPORT MODULES ===================================================================================================================================================

Export-ModuleMember -Function "Update-PS7"