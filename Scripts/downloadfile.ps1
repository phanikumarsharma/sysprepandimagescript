$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "C:\Program Files\odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "C:\Program Files\odt.exe" -ArgumentList '/extract:"C:\Program Files\O365" /quiet'
Start-Sleep -s 10
Start-Process -FilePath "C:\Program Files\O365\setup.exe" -ArgumentList '/configure "C:\Program Files\O365\configuration-Office365-x64.xml" /quiet'

