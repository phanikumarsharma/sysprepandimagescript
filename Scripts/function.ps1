
Param( 

    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $azureLoginId,

    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $azureLoginPassword,
    
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $subscriptionId,
    
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $resourceGroup,

    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $location,

    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string] $vmName
    )
    
    # Install AzureRM Module   
        
    Write-Output "Checking if AzureRm module is installed.."
    $azureRmModule = Get-Module AzureRM -ListAvailable | Select-Object -Property Name -ErrorAction SilentlyContinue
    if (!$azureRmModule.Name) 
    {
        Write-Output "AzureRM module Not Available. Installing AzureRM Module"
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
        Install-Module Azure -Force
        Install-Module AzureRm -Force 
        Write-Output "Installed AzureRM Module successfully"
    } 
    else
    {
        Write-Output "AzureRM Module Available"
    }

    # Import AzureRM Module

    Write-Output "Importing AzureRm Module.."
    Import-Module AzureRm -ErrorAction SilentlyContinue -Force
    Import-Module AzureRM.profile
    Import-Module AzureRM.resources
    Import-Module AzureRM.Compute

    # Login to AzureRM Account

    Write-Output "Login Into Azure RM.."
    
    $Psswd = $Password | ConvertTo-SecureString -asPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($UserName,$Psswd)
    Login-AzureRmAccount -Credential $Credential

    # Select the AzureRM Subscription

    Write-Output "Selecting Azure Subscription.."
    Select-AzureRmSubscription -SubscriptionId $SubscriptionId
    $rg = Get-AzureRmResourceGroup -Name $resourceGroup
    $vm = Get-AzureRmVM -ResourceGroupName $rg.ResourceGroupName -Name $vmName
    $nic = Get-AzureRmNetworkInterface -ResourceGroupName $rg.ResourceGroupName -Name $(Split-Path -Leaf $VM.NetworkProfile.NetworkInterfaces[0].Id)
    $nic | Get-AzureRmNetworkInterfaceIpConfig | Select-Object @{'label'='DNS name';Expression={Set-Variable -name DNS -scope Global -value $(Split-Path -leaf $_.PublicIpAddress.Id);$dns}}
    $dnsName=(Get-AzureRmPublicIpAddress -ResourceGroupName $rg.ResourceGroupName -Name $dns).DnsSettings.Fqdn
    
    Enable-PSRemoting
    $Cert = New-SelfSignedCertificate -CertstoreLocation Cert:\LocalMachine\My -DnsName $dnsName
    Export-Certificate -Cert $Cert -FilePath 'C:\exch.cer'
    New-Item -Path "WSMan:\LocalHost\Listener" -Transport HTTPS -Address * -CertificateThumbPrint $Cert.Thumbprint -Force
    New-NetFirewallRule -Name "winrm_https" -DisplayName "winrm_https" -Enabled True -Profile Any -Action Allow -Direction Inbound -LocalPort 5986 -Protocol TCP

    # Install O365
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
    function O365
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
    Out-Null | O365
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

