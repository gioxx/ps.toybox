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
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$False -SkipLoadingCmdletHelp
  }
}

# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Connect-EOL"