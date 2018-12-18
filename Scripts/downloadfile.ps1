$url = "https://public.bn.files.1drv.com/y4meAan5lYibg8HV6m4Y0eexecP381-yQ_Dp9cx7abdrCI78KbJsDEaFpXts4Kxf45o8Kl6nlMPteFpVTY4mL7fO1WC6FLwnW9oX-tpqRb80GaR2LG76nDXkLspB0pWi1pYbg9KPV-j0v8YmldjvPVWoYPBgXknj9c6NgaT8iKz3_8yB58eVrqSgUwK7Ns9P-su/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso?access_token=EwAQA61DBAAUcSSzoTJJsy%2bXrnQXgAKO5cj4yc8AARkvASLnuD7PBsD6YH8tMnLWEaRpH1QXQ2jYXeCn1HbMkQk5QVahADoMKilHiSW8mKC%2b%2b2zsRG2FjID2MW0V3AvrF796RI48ocPrIkb7JKVhZWBHuvcGBFxmdY6jlggqBePhRFGg8OhYZu33QJhm8iSeWSJcR76Er53loNWaqDO3oR2b336jNvuwUU%2f0GQw6LdhnZsViiW53ktNHIoWB79xiISBTO4Pkihrkb1fhuCEztK%2byedfqIdL9rZMlrJxjCMowixo9vbAyzvoMujzqGLPcra9A8LRdOhLH89nvy7mq9TUi1nmxutIiJmKYeOwLVHb63nCJScITCrr43rJfL04DZgAACASKYb%2fyZinV4AGGoXGAZ%2fGGJ2d5BsH8kIwn3R%2faWiX5yfcA3CiTndgsWTkPSqD78swpgUY1XlySwFUvoRSUhPb6IxAwWax6%2fJCS3YDXpDnVGRBJTUc0zg437r%2fiZXu2Gs9WggvsLlHasOT%2bs3D02%2ba40p0boCZQVf%2fiNrxlZ6bgu2Fp3h17BhK5weNtyP0f0cVRM1ZywuL4XDr7sR8s6gh2wEyQccL0K6b3zFBfSpukpFSnHJNnUGA6VVF0aRtNyfi9UML%2b2k7kUr7e6JSaf2cu46Fh0aSvRl6ZiWQQbgyXzOj%2f2mjDlA%2foYylMsrUZGM%2bbP9hrJxOUUFD70bxPTnCaN1HFrRdEGDBW%2buQtEnAhuvQUYrOfwEmTfAC%2fb5OA6mfalf5ht%2benQzDXAN2TCxntFPJZsPoRU96gkidj5O2c4oWkxFnFnAptAhzBH4dT3OYD3lgVUV5XiL18%2bsEFlFSa9ppO5azSVPgWj3aRVVgk1tV%2b%2bRugQQjf2qmN9R2Gp1ijztzWV9qn7xuQdC3PPOp4qnZocUp%2fiUTgA4p4gUiqFIQ4TeyE9hgVQwLZwhfAwtNCMgE1pSGCPEeORf8ZTnd%2fcTwUtWUk0GSMdURC8dk7bETWWh2i0hc2xigu1dc4K9JRPOzIS5U0YxMeAg%3d%3d"
$output = "D:/office.iso"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
#to extract the ODT tool
Start-Process -FilePath "D:/office.iso" -ArgumentList '/extract:"D:/Office" /quiet'

Set-Location "E:\"
#Invoke-Expression -Command "cmd.exe /c '.\setup.exe' /download 'D:/Office/configuration-Office365-x64.xml'" 
Invoke-Expression -Command "cmd.exe /c .\Setup.exe /configure 'D:/configuration-Office365-x64.xml'" 

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


$url = "https://public.bn.files.1drv.com/y4meAan5lYibg8HV6m4Y0eexecP381-yQ_Dp9cx7abdrCI78KbJsDEaFpXts4Kxf45o8Kl6nlMPteFpVTY4mL7fO1WC6FLwnW9oX-tpqRb80GaR2LG76nDXkLspB0pWi1pYbg9KPV-j0v8YmldjvPVWoYPBgXknj9c6NgaT8iKz3_8yB58eVrqSgUwK7Ns9P-su/en_office_professional_plus_2019_x86_x64_dvd_7ea28c99.iso?access_token=EwAQA61DBAAUcSSzoTJJsy%2bXrnQXgAKO5cj4yc8AARkvASLnuD7PBsD6YH8tMnLWEaRpH1QXQ2jYXeCn1HbMkQk5QVahADoMKilHiSW8mKC%2b%2b2zsRG2FjID2MW0V3AvrF796RI48ocPrIkb7JKVhZWBHuvcGBFxmdY6jlggqBePhRFGg8OhYZu33QJhm8iSeWSJcR76Er53loNWaqDO3oR2b336jNvuwUU%2f0GQw6LdhnZsViiW53ktNHIoWB79xiISBTO4Pkihrkb1fhuCEztK%2byedfqIdL9rZMlrJxjCMowixo9vbAyzvoMujzqGLPcra9A8LRdOhLH89nvy7mq9TUi1nmxutIiJmKYeOwLVHb63nCJScITCrr43rJfL04DZgAACASKYb%2fyZinV4AGGoXGAZ%2fGGJ2d5BsH8kIwn3R%2faWiX5yfcA3CiTndgsWTkPSqD78swpgUY1XlySwFUvoRSUhPb6IxAwWax6%2fJCS3YDXpDnVGRBJTUc0zg437r%2fiZXu2Gs9WggvsLlHasOT%2bs3D02%2ba40p0boCZQVf%2fiNrxlZ6bgu2Fp3h17BhK5weNtyP0f0cVRM1ZywuL4XDr7sR8s6gh2wEyQccL0K6b3zFBfSpukpFSnHJNnUGA6VVF0aRtNyfi9UML%2b2k7kUr7e6JSaf2cu46Fh0aSvRl6ZiWQQbgyXzOj%2f2mjDlA%2foYylMsrUZGM%2bbP9hrJxOUUFD70bxPTnCaN1HFrRdEGDBW%2buQtEnAhuvQUYrOfwEmTfAC%2fb5OA6mfalf5ht%2benQzDXAN2TCxntFPJZsPoRU96gkidj5O2c4oWkxFnFnAptAhzBH4dT3OYD3lgVUV5XiL18%2bsEFlFSa9ppO5azSVPgWj3aRVVgk1tV%2b%2bRugQQjf2qmN9R2Gp1ijztzWV9qn7xuQdC3PPOp4qnZocUp%2fiUTgA4p4gUiqFIQ4TeyE9hgVQwLZwhfAwtNCMgE1pSGCPEeORf8ZTnd%2fcTwUtWUk0GSMdURC8dk7bETWWh2i0hc2xigu1dc4K9JRPOzIS5U0YxMeAg%3d%3d"
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

