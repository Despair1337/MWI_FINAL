$url = "https://github.com/TheWover/donut/releases/download/v1.1/donut_v1.1.zip"
$output = "donut_v1.1.zip"

Invoke-WebRequest -Uri $url -OutFile $output