$url = "https://unmstrgaccnt9.blob.core.windows.net/iso/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99%20(1).iso"
$output = "D:/office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
Start-Process -FilePath "D:\office.iso" -ArgumentList '/extract:"D:/Office2019" /quiet'
Set-Location "E:\"
Start-Process -FilePath ".\Setup.exe"
