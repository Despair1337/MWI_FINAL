.bin payload to b64:
Powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes($fn)) | clip