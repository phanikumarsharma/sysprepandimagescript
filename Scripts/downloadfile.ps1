$url = "https://notepad-plus-plus.org/repository/7.x/7.6/npp.7.6.Installer.x64.exe"
$output = "C:\Program Files\Notepad++\npp.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output} -Credential $cred
