# Define variables
$siteUrl = "https://frcidevtest.sharepoint.com/sites/ContractMgt"
$listName = "Contract_Request"
$fieldName = "TypeOfContract"
$newValue = "Addendum"

# Connect to the site
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get all items in the list
$listItems = Get-PnPListItem -List $listName -PageSize 1000

# Update each item
foreach ($item in $listItems) {
    Set-PnPListItem -List $listName -Identity $item.Id -Values @{$fieldName = $newValue}
    Write-Output "Updated item with ID: $($item.Id)"
}

Write-Output "All items updated successfully."
