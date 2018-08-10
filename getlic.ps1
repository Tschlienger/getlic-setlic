# Define variables
$licFilePath = ".\lic.csv"

# Get product info and key from current PC
$product = (Get-WmiObject Win32_OperatingSystem).Caption
$key = (Get-WmiObject SoftwareLicensingService).OA3xOriginalProductKey

# Import key list
$list = Import-Csv $licFilePath

# Determine if key already exists in list
If ($null -ne ($list | Where-Object {$_.Key -eq $key})) {
    # Notify the user and abort
    (New-Object -ComObject Wscript.Shell).Popup("The product key for $product is already listed!", 0, "Error", 0+48)
} Else {
    # Write product info and key to key list and notify the user
    [PSCustomObject]@{Product = $product; Key = $key} | Export-Csv $licFilePath -Append -NoTypeInformation -Encoding ASCII
    (New-Object -ComObject Wscript.Shell).Popup("Product key for $product was successfully stored in license file.", 0, "Success", 0+64)
}
