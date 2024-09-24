# Module manifest for module 'Gioxx.ToyBox'
# Generated by: Gioxx
# Generated on: 01/07/2022

@{
  RootModule = '.\GTB.Main.psm1'
  NestedModules = @(
    ".\Calendar\GTB.Calendar.ps1",
    ".\Groups\GTB.Groups.ps1",
    ".\Mailboxes\GTB.Mboxes.ps1",
    '.\Main\GTB.Connections.ps1',
    '.\Main\GTB.Tools.ps1',
    ".\Protection\GTB.Protection.ps1",
    ".\Protection\GTB.Quarantine.ps1",
    ".\Rooms\GTB.Room.ps1",
    ".\Statistics\GTB.Stats.ps1",
    '.\Xtras\GTB.Custom.ps1',
    '.\Xtras\Write-InformationColored.ps1'
  )

  ModuleVersion = '0.5'
  GUID = '17aadfab-2909-411f-9266-29177b510755'
  Author = 'Gioxx'
  CompanyName = 'Gioxx.org'
  Copyright = '(c)opyleft, since the dawn of time, Gioxx.org'

  # Description of the functionality provided by this module
  # Description = ''

  PowerShellVersion = '7.0'

  FunctionsToExport = @(
    "Add-MboxAlias",
    "Add-MboxPermission",
    "Change-MboxLanguage",
    "Change-MFAStatus",
    "Clone-OoOMessage",
    "Connect-EOL",
    "Connect-MSOnline",
    "Export-CalendarPermission",
    "Export-DDG",
    "Export-DG",
    "Export-M365Group",
    "Export-MboxAlias",
    "Export-MboxPermission",
    "Export-MboxStatistics",
    "Export-MFAStatus",
    "Export-MsolAccountSku",
    "Export-QuarantineEML",
    "Get-MboxAlias",
    "Get-MboxPermission",
    "Get-QuarantineFrom",
    "Get-QuarantineFromDomain",
    "Get-QuarantineToRelease",
    "Get-RandomPassword",
    "Get-RoomsDetails",
    "Get-RoleGroupsMembers",
    "Get-UserGroups",
    "New-SharedMailbox",
    "Release-QuarantineFrom",
    "Release-QuarantineMessageId",
    "Remove-MboxAlias",
    "Remove-MboxPermission",
    "Set-MboxRulesQuota",
    "Set-OoO",
    "Set-SharedMboxCopyForSent",
    "Update-PS7",
    "User-CloseAllPSSessions",
    "User-DisableDevices",
    "User-DisableSignIn",
    "User-SignOut"
  )
  
  CmdletsToExport = @()

  VariablesToExport = '*' # Variables to export from this module
  AliasesToExport = @(
    "rqf"
  ) # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.

  PrivateData = @{
    PSData = @{
      LicenseUri = 'https://github.com/gioxx/gioxx.toybox/blob/main/LICENSE'
      ProjectUri = 'https://github.com/gioxx/gioxx.toybox/'

      # ReleaseNotes of this module
      # ReleaseNotes = ''

      Prerelease = '20240924(3)'
      RequireLicenseAcceptance = $False

    }

  }

  # HelpInfo URI of this module
  # HelpInfoURI = ''

}
