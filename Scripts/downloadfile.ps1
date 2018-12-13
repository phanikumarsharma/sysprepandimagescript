$output = "C:\Program Files\ODT\odt.exe"

Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output

