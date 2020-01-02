# Script to setup golden image with Azure Image Builder

# Create a folder
New-Item -Path 'C:\temp' -ItemType Directory -Force | Out-Null

# Install VSCode
Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?Linkid=852157' -OutFile 'c:\temp\VScode.exe'
Invoke-Expression -Command 'c:\temp\VScode.exe /verysilent'

#Start sleep
Start-Sleep -Seconds 10

#Install Notepad++
Invoke-WebRequest -Uri 'https://notepad-plus-plus.org/repository/7.x/7.7.1/npp.7.7.1.Installer.x64.exe' -OutFile 'c:\temp\notepadplusplus.exe'
Invoke-Expression -Command 'c:\temp\notepadplusplus.exe /S'

#Start sleep
Start-Sleep -Seconds 10

# Install FSLogix
Invoke-WebRequest -Uri 'https://aka.ms/fslogix_download' -OutFile 'c:\temp\fslogix.zip'
Start-Sleep -Seconds 10
Expand-Archive -Path 'C:\temp\fslogix.zip' -DestinationPath 'C:\temp\fslogix\'  -Force
Invoke-Expression -Command 'C:\temp\fslogix\x64\Release\FSLogixAppsSetup.exe /install /quiet /norestart'

# Install Office 365  
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12130-20272.exe"
$output = 'C:\temp\odt.exe'
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
New-Item -Path "C:\temp\O365" -ItemType Directory -Force -ErrorAction SilentlyContinue
Start-Process -FilePath "C:\temp\odt.exe" -ArgumentList '/extract:"C:\temp\O365" /quiet'
Start-Sleep -Seconds 15
Set-Location "C:\temp\O365"
Invoke-Expression -Command "cmd.exe /c 'C:\temp\O365\setup.exe' /download 'C:\temp\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'C:\temp\O365\setup.exe' /configure 'C:\temp\O365\configuration-Office365-x64.xml'" 

