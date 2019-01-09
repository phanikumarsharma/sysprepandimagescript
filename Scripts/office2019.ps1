    function O365
    {
    $url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
    $output = "D:/odt.exe"
    Import-Module BitsTransfer
    Start-BitsTransfer -Source $url -Destination $output
    #to extract the ODT tool
    New-Item -Path "D:\Office2019" -ItemType Directory -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
    Start-Sleep -Seconds 15
    Set-Location "D:\O365"
    Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
    Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
    }
    Out-Null | O365
