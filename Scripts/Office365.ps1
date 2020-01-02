#Script to setup golden image with Azure Image Builder


#Create temp folder
New-Item -Path 'C:\temp' -ItemType Directory -Force | Out-Null


#Install VSCode
Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?Linkid=852157' -OutFile 'c:\temp\VScode.exe'
Invoke-Expression -Command 'c:\temp\VScode.exe /verysilent'

#Start sleep
Start-Sleep -Seconds 10

#InstallNotepadplusplus
Invoke-WebRequest -Uri 'https://notepad-plus-plus.org/repository/7.x/7.7.1/npp.7.7.1.Installer.x64.exe' -OutFile 'c:\temp\notepadplusplus.exe'
Invoke-Expression -Command 'c:\temp\notepadplusplus.exe /S'

#Start sleep
Start-Sleep -Seconds 10

#InstallFSLogix
Invoke-WebRequest -Uri 'https://aka.ms/fslogix_download' -OutFile 'c:\temp\fslogix.zip'
Start-Sleep -Seconds 10
Expand-Archive -Path 'C:\temp\fslogix.zip' -DestinationPath 'C:\temp\fslogix\'  -Force
Invoke-Expression -Command 'C:\temp\fslogix\x64\Release\FSLogixAppsSetup.exe /install /quiet /norestart'

    
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12130-20272.exe"
$output = 'C:\temp'
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
New-Item -Path "C:\temp\O365" -ItemType Directory -Force -ErrorAction SilentlyContinue
Start-Process -FilePath "C:\temp\O365\odt.exe" -ArgumentList '/extract:"C:\temp\O365" /quiet'
Start-Sleep -Seconds 15
Set-Location "C:\temp\O365"
Invoke-Expression -Command "cmd.exe /c 'C:\temp\O365\setup.exe' /download 'C:\temp\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'C:\temp\O365\setup.exe' /configure 'C:\temp\O365\configuration-Office365-x64.xml'" 

<#

.Synposys
Install O365 and Enable WinRm

.Description
This script is used to install the Office 365 and to enable the WinRm.

.Permission
Administrator
#>
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
<#
$dnsName= (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
Enable-PSRemoting
$Cert = New-SelfSignedCertificate -CertstoreLocation Cert:\LocalMachine\My -DnsName $dnsName
Export-Certificate -Cert $Cert -FilePath 'C:\exch.cer'
New-Item -Path "WSMan:\LocalHost\Listener" -Transport HTTPS -Address * -CertificateThumbPrint $Cert.Thumbprint -Force
New-NetFirewallRule -Name "winrm_https" -DisplayName "winrm_https" -Enabled True -Profile Any -Action Allow -Direction Inbound -LocalPort 5986 -Protocol TCP
#>
# Install O365
<#function O365
{
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
New-Item -Path "D:\O365" -ItemType Directory -Force -ErrorAction SilentlyContinue
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
Start-Sleep -Seconds 15
Set-Location "D:\O365"
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
}
Out-Null | O365#>
<#
Enable-PSRemoting
$Cert = New-SelfSignedCertificate -CertstoreLocation Cert:\LocalMachine\My -DnsName $dnsName
Export-Certificate -Cert $Cert -FilePath 'C:\exch.cer'
New-Item -Path "WSMan:\LocalHost\Listener" -Transport HTTPS -Address * -CertificateThumbPrint $Cert.Thumbprint -Force
New-NetFirewallRule -Name "winrm_https" -DisplayName "winrm_https" -Enabled True -Profile Any -Action Allow -Direction Inbound -LocalPort 5986 -Protocol TCP
#>
<#
function sysprep
{
$checkinstalled=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Office 365 ProPlus - en-us"}
if($checkinstalled)
{
Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet'
}
}
Out-Null | sysprep
#>

