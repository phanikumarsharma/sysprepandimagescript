Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:\odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:\odt.exe" -ArgumentList '/extract:"D:\O365" /quiet'
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
#Invoke-Expression -Command "cmd.exe /c D:/O365/setup.exe /configure 'D:/O365/configuration-Office365-x64.xml'"

#Start-Job -ScriptBlock { Start-Process -FilePath "$env:Temp\O365\setup.exe" -ArgumentList '/configure "$env:Temp\O365\configuration-Office365-x64.xml"'}

