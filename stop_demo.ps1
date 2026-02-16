$ErrorActionPreference = "SilentlyContinue"

$stopped = @()

$portPids = @()
$portPids += (netstat -ano | Select-String ":8000" | ForEach-Object { ($_ -split "\s+")[-1] } | Where-Object { $_ -match '^\d+$' } | Select-Object -Unique)

foreach ($procId in $portPids) {
    Stop-Process -Id ([int]$procId) -Force
    $stopped += "port8000:$procId"
}

Get-Process -Name cloudflared -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.Id -Force
    $stopped += "cloudflared:$($_.Id)"
}

if ($stopped.Count -eq 0) {
    "No demo processes were running."
} else {
    "Stopped: $($stopped -join ', ')"
}
