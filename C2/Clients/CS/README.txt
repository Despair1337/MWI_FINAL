.bin payload to b64:
Powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("D:\REPO\MWI_FINAL\C2\Clients\CS\loader.bin")) | clip