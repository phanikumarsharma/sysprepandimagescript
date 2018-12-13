$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:\Program Files\ODT\odt.exe"

Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
