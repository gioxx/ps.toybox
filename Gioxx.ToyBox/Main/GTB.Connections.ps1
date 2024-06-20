# Connections ======================================================================================================================================================

function Connect-EOL {
  param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage="User to connect to Exchange Online with")]
    [string] $UserPrincipalName
  )

  if ( (Get-Module -Name ExchangeOnlineManagement -ListAvailable).count -eq 0 ) {
    Write-Host "Install the ExchangeOnlineManagement module using this command (then relaunch this script): `nInstall-Module ExchangeOnlineManagement" -f "Yellow"
  } else {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$False
  }
}

function Connect-MSOnline {
  if ( (Get-Module -Name MSOnline -ListAvailable).count -eq 0 ) {
    Write-Host "Install the MSOnline module using this command (then relaunch this script): `nInstall-Module MSOnline" -f "Yellow"
  } else {
    Import-Module MSOnline -UseWindowsPowershell
    Connect-MsolService | Out-Null
    Import-Module MSOnline
  }
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Connect-EOL"
Export-ModuleMember -Function "Connect-MSOnline"