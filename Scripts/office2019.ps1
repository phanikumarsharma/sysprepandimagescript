$url = "https://office19rgdisks.blob.core.windows.net/office19-iso/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso"
$output = "C:\office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
Start-Process -FilePath "C:\office.iso" -ArgumentList '/extract:"E:\" /quiet'
Start-Sleep -Seconds 5
Set-Location "E:\Office"
Start-Process -FilePath ".\Setup64.exe"


