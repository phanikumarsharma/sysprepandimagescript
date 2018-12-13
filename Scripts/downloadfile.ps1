$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "C:\Program Files\odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
$path=Start-Process -FilePath "C:\Program Files\odt.exe" -ArgumentList '/extract:"C:\Program Files\O365" /quiet'
Start-Process -FilePath "$path\setup.exe" -ArgumentList '/configure "$path\configuration-Office365-x64.xml"'

