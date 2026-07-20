param([string]$Mode = "local")
$OutputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Done = $false

Write-Host "=== PPT Render Engine ===" -ForegroundColor Cyan

# ---- Start Render Engine ----
Write-Host "[1/2] Starting Render Engine..." -ForegroundColor Yellow
$engineJob = Start-Job -Name "RenderEngine" -ScriptBlock {
    param($dir)
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $OutputEncoding = [Text.Encoding]::UTF8
    Set-Location $dir
    $env:PYTHONUNBUFFERED = "1"
    $env:PYTHONIOENCODING = "utf-8"
    uv run python run.py 2>&1
} -ArgumentList "$Root\ppt_render_engine"

# Wait for Render Engine to be ready
$ready = $false
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    foreach ($line in Receive-Job $engineJob) {
        Write-Host "[ENGINE] $line" -ForegroundColor Cyan
    }
    try {
        $r = Invoke-WebRequest -Uri "http://127.0.0.1:8000/health" -Method GET -UseBasicParsing -TimeoutSec 2
        if ($r.StatusCode -eq 200) { $ready = $true; break }
    } catch {}
}
if (-not $ready) {
    Write-Host "Render Engine failed to start within 30s" -ForegroundColor Red
    Stop-Job $engineJob -ErrorAction SilentlyContinue; Remove-Job $engineJob -ErrorAction SilentlyContinue
    exit 1
}
Write-Host "[ENGINE] -> http://127.0.0.1:8000" -ForegroundColor Green

# ---- Start MCP Server ----
Write-Host "[2/2] Starting MCP Server..." -ForegroundColor Yellow
$mcpJob = Start-Job -Name "MCPServer" -ScriptBlock {
    param($dir)
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $OutputEncoding = [Text.Encoding]::UTF8
    Set-Location $dir
    $env:PYTHONUNBUFFERED = "1"
    $env:PYTHONIOENCODING = "utf-8"
    uv run python -m mcp_server --mode $using:Mode 2>&1
} -ArgumentList "$Root\mcp_server"

Start-Sleep -Seconds 1
Write-Host "[MCP] -> http://127.0.0.1:8010" -ForegroundColor Green
Write-Host "Press Ctrl+C to stop both services." -ForegroundColor Cyan

# Loop: print logs from both services until Ctrl+C
try {
    while (-not $Done) {
        Start-Sleep -Milliseconds 200
        foreach ($line in Receive-Job $engineJob) {
            if ($line) { Write-Host "[ENGINE] $line" -ForegroundColor Cyan }
        }
        foreach ($line in Receive-Job $mcpJob) {
            if ($line) { Write-Host "[MCP   ] $line" -ForegroundColor Magenta }
        }
        # Check if either job died
        $engineJob.State -eq "Running" -and $mcpJob.State -eq "Running" | Out-Null
        if ($engineJob.State -ne "Running" -or $mcpJob.State -ne "Running") {
            Write-Host "`nA service has stopped unexpectedly." -ForegroundColor Red
            break
        }
    }
} finally {
    Write-Host "`nStopping services..." -ForegroundColor Yellow
    Stop-Job $engineJob -ErrorAction SilentlyContinue; Remove-Job $engineJob -ErrorAction SilentlyContinue
    Stop-Job $mcpJob -ErrorAction SilentlyContinue; Remove-Job $mcpJob -ErrorAction SilentlyContinue
    Write-Host "Stopped." -ForegroundColor Green
}
