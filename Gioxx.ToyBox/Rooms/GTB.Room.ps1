# Rooms ============================================================================================================================================================

function Get-RoomsDetails {
  param(
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Export results in a CSV file")]
    [switch] $CSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Folder where export CSV file (e.g. C:\Temp)")]
    [string] $folderCSV,
    [Parameter(Mandatory=$False, ValueFromPipeline=$True, HelpMessage="Show results in a grid view")]
    [switch] $GridView
  )

  priv_SetPreferences -Verbose
  $eolConnectedCheck = priv_CheckEOLConnection
  $mboxCounter = 0
  $arr_RoomDetails = @()

  if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
  if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

  if ( $eolConnectedCheck -eq $true ) {
    $Locations = Get-DistributionGroup -RecipientTypeDetails RoomList
    $Locations | ForEach {
      $CurrentUser = $_

      $mboxCounter++
      $PercentComplete = (($mboxCounter / $Locations.Count) * 100)
      Write-Progress -Activity "Processing $($CurrentUser.DisplayName)" -Status "$mboxCounter out of $($Locations.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

      Get-DistributionGroupMember $CurrentUser.PrimarySmtpAddress | ForEach {
        $arr_RoomDetails += New-Object -TypeName PSObject -Property $([ordered]@{
            "Location" = $CurrentUser.Name
            "Location PrimarySmtpAddress" = $CurrentUser.PrimarySmtpAddress
            "Room Display Name" = $_.DisplayName
            "Room PrimarySmtpAddress" = $_.PrimarySmtpAddress
            "Room Capacity" = ($(Get-Mailbox $_.PrimarySmtpAddress).ResourceCapacity)
        })
      }
    }

    if ( $GridView ) {
      $arr_RoomDetails | Out-GridView -Title "M365 Rooms Details"
    } elseif ( $CSV ) {
      $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-Rooms.csv")
      $arr_RoomDetails | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
      $arr_RoomDetails | Out-Host
    }

  } else {
    Write-Error "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs."
  }

  priv_RestorePreferences
}


# Export Modules and Aliases =======================================================================================================================================

Export-ModuleMember -Alias *
Export-ModuleMember -Function "Get-RoomsDetails"