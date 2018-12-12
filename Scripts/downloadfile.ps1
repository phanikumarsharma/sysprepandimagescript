$storageDir = "C:\Program Files"
$webclient = New-Object System.Net.WebClient
$url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
$file = "$storageDir\odt.exe"
$webclient.DownloadFile($url,$file)
