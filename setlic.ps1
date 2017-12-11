# Get product info and key from current PC
$product = (Get-WmiObject Win32_OperatingSystem).Caption

# Import key list
$list = Import-Csv ".\lic.csv"

# Try to get license row with product info matching current PC
Try {$license = ($list | Where {$_.Product -eq $product})[0]}
# Notify the user and abort
Catch {(New-Object -ComObject Wscript.Shell).Popup("Couldn't find a key matching $product!", 0, "Error", 0+48);break}
# Notify the user
(New-Object -ComObject Wscript.Shell).Popup("Product key for $product was found. Activating Windows now.", 0, "Success", 0+64)
# Store key from license in variable
$key = $license.Key

# Activate Windows with selected key
$service = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService"
If ($service.InstallProductKey($key)) {
    (New-Object -ComObject Wscript.Shell).Popup("$product was activated successfully.", 0, "Success", 0+64)
} Else {
    (New-Object -ComObject Wscript.Shell).Popup("There was a problem activating $product!", 0, "Error", 0+48)
}
$service.RefreshLicenseStatus()

# Write key list to file removing used one
$list | Where {$_.Key -ne $key} | Export-Csv ".\lic.csv" -NoTypeInformation