param(
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$webOut = Join-Path $root ".tmp_web_stdout.log"
$webErr = Join-Path $root ".tmp_web_stderr.log"
$tunnelOut = Join-Path $root ".tmp_tunnel_stdout.log"
$tunnelErr = Join-Path $root ".tmp_tunnel_stderr.log"

# Stop processes currently bound to the app port for a clean restart.
$portPids = @()
try {
    $portPids = (netstat -ano | Select-String (":{0}" -f $Port) | ForEach-Object {
        ($_ -split "\s+")[-1]
    } | Where-Object { $_ -match '^\d+$' } | Select-Object -Unique)
} catch {
}
foreach ($procId in $portPids) {
    Stop-Process -Id ([int]$procId) -Force -ErrorAction SilentlyContinue
}

# Stop active cloudflared processes to avoid stale demo links.
Get-Process -Name cloudflared -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
}

foreach ($p in @($webOut, $webErr, $tunnelOut, $tunnelErr)) {
    if (Test-Path $p) {
        Remove-Item $p -Force -ErrorAction SilentlyContinue
        if (Test-Path $p) {
            Clear-Content -Path $p -ErrorAction SilentlyContinue
        }
    }
}

$env:PORT = "$Port"
$web = Start-Process -FilePath python `
    -ArgumentList @("run_web.py") `
    -WorkingDirectory $root `
    -RedirectStandardOutput $webOut `
    -RedirectStandardError $webErr `
    -PassThru

$healthy = $false
for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    try {
        $res = Invoke-WebRequest -UseBasicParsing -Uri ("http://127.0.0.1:{0}" -f $Port) -TimeoutSec 2
        if ($res.StatusCode -ge 200 -and $res.StatusCode -lt 500) {
            $healthy = $true
            break
        }
    } catch {
    }
}

if (-not $healthy) {
    throw "Web app failed to start on port $Port. Check $webErr"
}

$cloudflaredPath = Join-Path $root "cloudflared.exe"
if (-not (Test-Path $cloudflaredPath)) {
    throw "cloudflared.exe not found at $cloudflaredPath"
}

$tunnel = Start-Process -FilePath $cloudflaredPath `
    -ArgumentList @("tunnel", "--url", ("http://127.0.0.1:{0}" -f $Port)) `
    -WorkingDirectory $root `
    -RedirectStandardOutput $tunnelOut `
    -RedirectStandardError $tunnelErr `
    -PassThru

$publicUrl = $null
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    if (Test-Path $tunnelErr) {
        $errText = Get-Content $tunnelErr -Raw
        if ($errText -match "failed to request quick Tunnel") {
            throw "Cloudflare quick tunnel failed. Check $tunnelErr"
        }
        $matches = [regex]::Matches($errText, 'https://([a-zA-Z0-9-]+)\.trycloudflare\.com')
        foreach ($match in $matches) {
            if ($match.Groups[1].Value -ne "api") {
                $publicUrl = $match.Value
                break
            }
        }
        if ($publicUrl) { break }
    }
}

if (-not $publicUrl) {
    throw "Tunnel started but no public URL found yet. Check $tunnelErr"
}

[pscustomobject]@{
    public_url = $publicUrl
    local_url = "http://127.0.0.1:$Port"
    web_pid = $web.Id
    tunnel_pid = $tunnel.Id
    web_log = $webErr
    tunnel_log = $tunnelErr
}
