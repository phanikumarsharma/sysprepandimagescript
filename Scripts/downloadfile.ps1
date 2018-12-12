$url = "https://www.mozilla.org/en-US/firefox/download/thanks/"
$output = "C:\Program Files\Notepad++\firefox.exe"
Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output} -Credential $cred
