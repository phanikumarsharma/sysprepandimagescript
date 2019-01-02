Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
function O365
{
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
New-Item -Path "D:\O365" -ItemType Directory -Force -ErrorAction SilentlyContinue
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
Set-Location D:\O365
Invoke-Expression -Command "cmd.exe /c .\setup.exe' /download D:\O365\configuration-Office365-x64.xml" 
Invoke-Expression -Command "cmd.exe /c .\setup.exe /configure D:\O365\configuration-Office365-x64.xml" 
}
Out-Null | O365
function sysprep
{
$checkinstalled=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Office 365 ProPlus - en-us"}
if($checkinstalled)
  {
  Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet'
  }
}
Out-Null | sysprep
<#
function office365Installation
{
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
Set-Location "D:\O365"
Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
}
Out-Null | office365Installation
function second
{
$checkinstalled=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Office 365 ProPlus - en-us"}
if($checkinstalled)
{
Write-Host "yes its there and this command execuites after installation of O365"
}
}
second
#>
<#
function sysprep
{
$checkinstalled = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Office 365 ProPlus - en-us"} | Select-Object DisplayName
if($checkinstalled)
{
Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet'
}
}
sysprep
#>
