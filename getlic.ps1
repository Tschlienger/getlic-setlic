# Get product info and key from current PC
$product = (Get-WmiObject Win32_OperatingSystem).Caption
$key = (Get-WmiObject SoftwareLicensingService).OA3xOriginalProductKey

# Import key list
$list = Import-Csv ".\lic.csv"

# Determine if key already exists in list
If ($list | where {$_.Key -eq $key}) {
    # Notify the user and abort
    (New-Object -ComObject Wscript.Shell).Popup("This product key for $product is already listed!", 0, "Error", 0+48)
} Else {
    # Write product info and key to key list and notify the user
    ($product + "," + $key) >> "lic.csv"
    (New-Object -ComObject Wscript.Shell).Popup("Product key for $product was successfully stored in license file.", 0, "Success", 0+64)
}