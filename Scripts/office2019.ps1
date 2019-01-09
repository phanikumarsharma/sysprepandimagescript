    <#
    
    .Synposys
    Install Office2019 and Enable WinRm
    
    .Description
    This script is used to install the Office 365 and to enable the WinRm.
    
    .Permission
    Administrator
    #>
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
    $dnsName= (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
    Enable-PSRemoting
    $Cert = New-SelfSignedCertificate -CertstoreLocation Cert:\LocalMachine\My -DnsName $dnsName
    Export-Certificate -Cert $Cert -FilePath 'C:\exch.cer'
    New-Item -Path "WSMan:\LocalHost\Listener" -Transport HTTPS -Address * -CertificateThumbPrint $Cert.Thumbprint -Force
    New-NetFirewallRule -Name "winrm_https" -DisplayName "winrm_https" -Enabled True -Profile Any -Action Allow -Direction Inbound -LocalPort 5986 -Protocol TCP
    # Install O365
    function Office2019
    {
    $url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
    $output = "D:/odt.exe"
    Import-Module BitsTransfer
    Start-BitsTransfer -Source $url -Destination $output
    #to extract the ODT tool
    New-Item -Path "D:\Office2019" -ItemType Directory -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/Office2019" /quiet'
    Start-Sleep -Seconds 15
    Set-Location "D:\Office2019"
    Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /download 'D:\Office2019\configuration-Office365-x64.xml'" 
    Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\Office2019\configuration-Office365-x64.xml'" 
    }
    Out-Null | O365
