
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
Start-Job -Name Job1 -ScriptBlock {$url = "https://public.bn.files.1drv.com/y4mzRYMyVNKNUtTmjgfa-pQ3CjDDS7pL9d3JSLkPYMWdpB3OP9sM9vFAAfmEaBT2H9oBLA2flTYVHItJstifVVbMsQiSwttzxVzGk3ybsf-4_h4Ko0dqvxA7Ym0gjjODgVKxdhhWZ8kTEoj-GG1RzZNgRUkWjvEdTAVs5aNsnfU6GwPw9XrwV1rHPz3sf-9wbii/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso?access_token=EwAYA61DBAAUcSSzoTJJsy%2bXrnQXgAKO5cj4yc8AAQELxq%2fzpBI9ZqB3Dn5CWtcnHMAXlv24IW8fm2wgjoFyMzMsB74pPZmdI5aXxwGk3cYM6Xyd0sK8Hp4HXZLLvKhmxV33n20ZJAuf8o4tyDSl99XUspVpjdwlqVsmDB5wQ5KuZ1U22HROeUS5O%2fqWkblwwJUgYHgbiPJZ1XGksm4IlrOUTeyY1%2ftWZK1X0qYUH01JBWhPbKPrMzzeX%2fOG7Um38fHnVcHOTJNyteQl5qO2qFi6VLnu1olBBrT6pF6dt0XanZ%2bkm6oGusL19pYT%2f%2bB3ZdIXzHhNedZ6HjNrYdgacu7o5bWh2vRJsvsdk3WOD9dDxWrNoLTREyjNTXrNiRwDZgAACPC7M0hrEt%2bg6AHI%2bBw0PnvTela3GF4P%2bk56xIdkbqqp2RkW0oueHJCA7H2pKT2f9nKJAY9mpcFhh48im1DLV4r7RC3WGucT8afeRQg6MNtNNwEZS2oC4%2f4sHtMZChlZQO3PWK1wIjqJOcdp8avBJKrzGeJxcAakcLXNEVdwua3clBWOgj8VV%2bqqeHcamq3fLLPsUk4A2uMKm82u8DRyJV47CTi1X50Lcv8SSsX%2btxwz9GVKnrgvOz3uy6X5jX%2b6h1PnBSy95kSns3Kb1X5mKGIKoHBS6%2b3%2bs3oew7X1%2bqY0MGuH%2fFZHFVJvefEq4HEuGNek1VGELbaAqqX0i3KOcSFU7NsF%2fq9TtzLZVFHliwyufXbg1b9u0t9o5BbH0%2fRYnAhQ1F8v5TFLU1e8QYafCDUdyXYliS6nksxqgAZa0n2QgAaC2Te64Zd%2fbUwfshzcs3mPHxcjqmAkOv26sxV6%2bDd4YUIbZivzW82DcaqI1Nt0FHRevOyflO2Gty9mONJeScvTlql1w72v28WAiex8%2fi9ut%2bn5LVZhLf2Paraz9lOn0iEQ6FmaIF%2fSmgg3s6OUmYXD4gYbbPSnqBEAlJ9vk8H79x3WAQWhINVHpXZ07wKH7KOuj92YvCrekSx8w5mFkQ7CwLxYIM0v2RXwdf6kGnh98R4C"
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
