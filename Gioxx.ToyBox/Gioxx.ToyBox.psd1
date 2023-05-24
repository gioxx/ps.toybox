# Module manifest for module 'Gioxx.ToyBox'
# Generated by: Gioxx
# Generated on: 01/07/2022

@{
  RootModule = '.\Gioxx.ToyBox.psm1'
  ModuleVersion = '0.2'
  GUID = '17aadfab-2909-411f-9266-29177b510755'
  Author = 'Gioxx'
  CompanyName = 'Gioxx.org'
  Copyright = '(c)opyleft, since the dawn of time, Gioxx.org'

  # Description of the functionality provided by this module
  # Description = ''

  PowerShellVersion = '7.0'

  FunctionsToExport = @(
    "ConnectEOL",
    "ConnectMSOnline",
    "ExplodeDDG",
    "MboxAlias",
    "MboxPermission-Add",
    "MboxPermission-Remove",
    "MboxPermission",
    "MboxStatistics-Export",
    "MsolAccountSku-Export",
    "QuarantineRelease",
    "ReloadModule",
    "SharedMbox-New",
    "SmtpExpand"
  )
  CmdletsToExport = @(
    "ConnectEOL",
    "ConnectMSOnline",
    "ExplodeDDG",
    "MboxAlias",
    "MboxPermission-Add",
    "MboxPermission-Remove",
    "MboxPermission",
    "MboxStatistics-Export",
    "MsolAccountSku-Export",
    "QuarantineRelease",
    "ReloadModule",
    "SharedMbox-New",
    "SmtpExpand"
  )
  VariablesToExport = '*' # Variables to export from this module
  AliasesToExport = @() # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.

  PrivateData = @{
    PSData = @{
      LicenseUri = 'https://github.com/gioxx/gioxx.toybox/blob/main/LICENSE'
      ProjectUri = 'https://github.com/gioxx/gioxx.toybox/'

      # ReleaseNotes of this module
      # ReleaseNotes = ''

      Prerelease = 'Preview12'
      RequireLicenseAcceptance = $false

      }

  }

  # HelpInfo URI of this module
  # HelpInfoURI = ''

}
