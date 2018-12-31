
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
Start-Job -Name Job1 -ScriptBlock {$url = "https://public.bn.files.1drv.com/y4miLmAlb9h827b-Z8h6WVZPw7x6_iqkESQAyZG1TEfAiM3IPrcKK1XjXhowOgqQAodxoVdGKc9eT5z2ifW3OfwMwPnb32_C9bnwrZPNQyK84_g6Yjqc8Rl18AW5fmjr3K3_oMOyQH3vKwrE67rhSz9QoH-7OzpGesKnuUQt73JrSVY2fP_JIrNZ3YCNaEkAjXI/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso?access_token=EwAQA61DBAAUcSSzoTJJsy%2bXrnQXgAKO5cj4yc8AAS%2f4A4f2RVsn4%2fQnQ9aU%2fqlO7mZ%2bUtGVaIqD1W%2b6WIuACeV69fth4OyQMhst9Y9ZaiUE9WTIY3%2b7Wg1qzDHWtPTCZvLehKkoP9NTGy2GLC2p8i3Rm7iOUGea26%2f7zvEJ4w%2fRPzMmvjcUADUhIeR%2fBupnw7nSXgLJUko3nqEWfPGb6qBt42S3YfJ6eXhhSTG16QxT9AaQ532BwkUtuWan1cmj16C8Nq6b616Vpqhvgs4YvF3jI%2bAmbW0nzkExlgO8GL1AWRqro2hM3nazX0wzGmwttodgtdkDIxWFHhZG%2f8nT2inX8MqdODFvFT8rH8hCTMSz13h%2bQp1YmzQAUd1jhWEDZgAACDTm0sU9h4JW4AFu0NOxlaRiX0d3bEXRLnoPB7qzIwoDq6DGl1ASpUA0Sm%2bxguhtVqEQe%2f%2f1iAPK%2bWIygxHQeoZn3Mh8dUHT%2fTekcs%2fQQp5nVDrAnkoDsq%2fNxxJX4gSXSOZZBE92r5yOzEnyOD%2bw%2fHcL3KoEjg6DAF1AojvImwaLVgR8AtY4fwDPlERsBe01tNWNeo6WHOAX33OBudlIM4bBghmD1G8XSTaOYz87BQG0vz8iTaLzRsWU%2boHN8hCx8mV5GVA%2bAxJQdojgwVdHV2oDcQQ%2f2HHLVt5aAFwy9KHPRbSgS8oVUJ9Js3CqDy%2b6hnpAFiZDJzbSZC%2fR7NYzQ0QmJikeIOU6qgIcgyCjCGfnicZa1iw3uh4z4c1Nm3l1SO46w1SwUlsnqV%2fxD%2bSw%2b8btqwaYUWpG6EefXuzDMpZ%2fF%2fjbhjJI8hyMbUtuUmenNOq3ioDxp4oXEFBvwEesvGhWznP6OfXu8wvUx0vcgTmlIUz4fTqHgr2ehzrTrnUIrklxxwIETFQkyR7x8N1Xt5IYhjwXdGRXNVJWZZ5Hz84A2HI6waCB%2b%2fHn4bpQM91CrEYFdMeLlm29TPmOO3Y8%2favKrntwt03T2%2b71UXLvNitPcA%2bqQy48SlosMMGWlvGHS7YJqlW%2fVCKIIxweAg%3d%3d"
$output = "C:\office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
Mount-DiskImage -ImagePath "C:\office.iso"
Set-Location "E:\Office"
Start-Process -FilePath ".\Setup64.exe"
}
Wait-Job -Name Job1
Start-Job -Name Job2 -ScriptBlock {Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet'}
Wait-Job -Name Job2
Stop-Job -Name "Job2"
<#Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
function office365Installation{
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
Set-Location "D:/O365"
Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
}
office365Installation
function RunSysprep{
  Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet'
}
RunSysprep


$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe"
$output = "D:/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/odt.exe" -ArgumentList '/extract:"D:/O365" /quiet'
Set-Location "D:\O365"
Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:\O365\configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c 'D:\O365\setup.exe' /configure 'D:\O365\configuration-Office365-x64.xml'" 
#>
