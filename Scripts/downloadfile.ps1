$url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
$output = "C:/Users/vmadmin/Desktop/odt.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output
