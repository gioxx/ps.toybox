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

  $mboxCounter = 0
  $Result = @()
  Set-Variable ProgressPreference Continue

  if (-not([string]::IsNullOrEmpty($folderCSV))) { $CSV = $True }
  if ($CSV) { $folder = priv_CheckFolder($folderCSV) }

  $Locations = Get-DistributionGroup -RecipientTypeDetails RoomList
  $Locations | ForEach {
    $CurrentUser = $_

    $mboxCounter++
    $PercentComplete = (($mboxCounter / $Locations.Count) * 100)
    Write-Progress -Activity "Processing $($CurrentUser.DisplayName)" -Status "$mboxCounter out of $($Locations.Count) ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

    Get-DistributionGroupMember $CurrentUser.PrimarySmtpAddress | ForEach {
      $Result += New-Object -TypeName PSObject -Property $([ordered]@{
          "Location" = $CurrentUser.Name
          "Location PrimarySmtpAddress" = $CurrentUser.PrimarySmtpAddress
          "Room Display Name" = $_.DisplayName
          "Room PrimarySmtpAddress" = $_.PrimarySmtpAddress
          "Room Capacity" = ($(Get-Mailbox $_.PrimarySmtpAddress).ResourceCapacity)
      })
    }
  }

  if ( $GridView ) {
    $Result | Out-GridView -Title "M365 Rooms Details"
  } elseif ( $CSV ) {
    $CSVfile = priv_SaveFileWithProgressiveNumber("$($folder)\$((Get-Date -format "yyyyMMdd").ToString())_M365-Rooms.csv")
    $Result | Export-CSV $CSVfile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
  } else {
    $Result
  }
}


# Export Modules ===================================================================================================================================================

Export-ModuleMember -Function "Get-RoomsDetails"