$url = "https://office19rgdisks.blob.core.windows.net/office19-iso/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso"
$output = "C:\office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
Mount-DiskImage -ImagePath "C:\office.iso"
Set-Location "E:\Office"
Start-Process -FilePath ".\Setup64.exe"


