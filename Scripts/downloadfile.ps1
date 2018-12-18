<#Set-Location "E:\"
Invoke-Expression -Command "cmd.exe /c .\Setup.exe /configure D:\O365\configuration-Office365-x64.xml"
#>
<#
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
#$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'

Set-Location "D:/O365"
Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:/O365/configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:/O365/setup.exe' /configure 'D:/O365/configuration-Office365-x64.xml'" 
#Invoke-Expression -Command "cmd.exe /c D:/O365/setup.exe /configure 'D:/O365/configuration-Office365-x64.xml'"

#Start-Job -ScriptBlock { Start-Process -FilePath "$env:Temp\O365\setup.exe" -ArgumentList '/configure "$env:Temp\O365\configuration-Office365-x64.xml"'}

#>


$url = "https://public.bn.files.1drv.com/y4mVCpfV4oQUE_jLVB5TS4rHhXiW8kFtiGnZMf2TlfdoxXhA8_1utJI66FJdY86Rvcxi26dp2lBul6cQy4NMDrSVbYFHj7KydCmyC5-DTWt8i1WydA1yyuzuKkSnTgxTHieOgA1Oa6ZbETM1LFQx3ZZUqtdC_n6-_DJYhvEChmafOfUT6n36EQwlo4d9FeobRLG/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso?access_token=EwAQA61DBAAUcSSzoTJJsy%2bXrnQXgAKO5cj4yc8AAVO%2fl9TiLvaS5SbKZE4KDYatDbz6PAtCIm8M%2fkA1ciugpTiZRiyEqlvRiN9kyLVb2CsFC%2b4Ncju1oSLR6lq2Z1Lug0MXYM21nPbxncaXlTZtl7%2fhFka3srzSvOE%2bVb7wgf8zzVXrsncTjHgEdvphdOv6mkjVNdw152hMaYYY690blRFtYs98A18oHA%2byfWLWWp%2bk%2fub%2buFN3fwMFRB6TE%2bnV5T57gFicMDNd3yOxrpXLNG2Z%2byH7HTfoYXwdJ3PinVO3vqiRdODcn63py5xZ%2fxa3k%2f%2bvpsWIzR2hFCZ3VArfwqQfj6BMeuVLpSYoJzxgIqyhGv9da4B8tMfLLRx5BmoDZgAACK1bGJ1a1wVO4AELsIwdamrFQpssi25Cc17fbvEg%2f3xkK1HpUP46dBnHf0enTqcqzlG%2fZqOwvEHJlPAsErel6kNiSaBBr3Ao6Ok1GBYxqJAxVze%2bTEW8b2BX7X4c5ka6zkB7Q9mPrzYexbOU7U760DyF8bY9E6Fd6AvwNOYaHrUEQ7oVsKSxD48NaucoqW6wQX6c8kyGYTh0ZGP5lR%2fHanwLF9yonTCQbbejbbkOy%2b8eBd6M3uI7zfaBzB1EyJBupMUt3Tk6UtpTIHa7MiZikO%2fUCQLqn76ZBSZy0dmtrZgX05fmYcREPmLzJ3Q%2fnRl0oJFkDp0HqcBPYW95Mqu%2b3KHcY5rS%2fjm73%2fa%2fBhBcRTYTqJfuynb40X07buZ93tlUTBBKc%2fxz2oGLdBriWmW1gX09XFQSQfK4laeshkAyEsDL1gzEIf4eMBWQM5vdIoLl%2b9arELqq61rxZJqsBdnNcSwcq40iHYafEecMFzCURmOhFoJUqGytkzVDrBLTRa138erZhpszVDcu7Kr0f9sgb7XXi%2bTPwWvvDBj%2bFEzg1Zbjir%2f%2fWbo1AHPZQE%2fuCAWGnktENNOlSPH5wVMgJSBkeNtrMVPuPDZhQvJNNTWwlN78cKG%2fEUiIdF8H7PKVUU63kHxerwU5hi497jweAg%3d%3d"
$output = "D:/office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/office.iso" -ArgumentList '/extract:"D:/Office" /quiet'

Set-Location "E:\"
#Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:/Office/configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c .\Setup.exe /configure 'D:/configuration-Office365-x64.xml'" 
#Invoke-Expression -Command "cmd.exe /c D:/O365/setup.exe /configure 'D:/O365/configuration-Office365-x64.xml'"

#Start-Job -ScriptBlock { Start-Process -FilePath "$env:Temp\O365\setup.exe" -ArgumentList '/configure "$env:Temp\O365\configuration-Office365-x64.xml"'}

