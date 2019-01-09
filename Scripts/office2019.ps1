    function Office2019
    {
    $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
    $output = "D:/odt.exe"
    Import-Module BitsTransfer
    Start-BitsTransfer -Source $url -Destination $output
    #to extract the ODT tool
    New-Item -Path "D:\Office2019" -ItemType Directory -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/Office2019" /quiet'
    Start-Sleep -Seconds 15
    Set-Location "D:\Office2019"
    Invoke-Expression -Command "cmd.exe /c 'D:\Office2019\setup.exe' /download 'D:\Office2019\configuration-Office2019Enterprise.xml'" 
    Invoke-Expression -Command "cmd.exe /c 'D:\Office2019\setup.exe' /configure 'D:\Office2019\configuration-Office2019Enterprise.xml'" 
    }
    Out-Null | Office2019
