#Requires -Version 5.1
<#
.SYNOPSIS
  Deploy & kelola aplikasi GotongRoyong di Windows memakai PM2.

.DESCRIPTION
  - Tanpa flag : install dependency + PM2, lalu start aplikasi.
  - ADMIN_PASSWORD & PORT bisa diubah (disimpan ke file .env).

.EXAMPLE
  .\deploy.ps1 -AdminPassword 'Rahasia123'
  .\deploy.ps1 -AdminPassword 'Rahasia123' -Port 8080
  .\deploy.ps1 -Restart
  .\deploy.ps1 -Status
  .\deploy.ps1 -Logs
  .\deploy.ps1 -Stop
#>
[CmdletBinding()]
param(
    [string]$AdminPassword = '',
    [string]$Port = '',
    [switch]$Start,
    [switch]$Stop,
    [switch]$Restart,
    [switch]$Status,
    [switch]$Logs,
    [switch]$InstallDeps
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$NPM = 'npm.cmd'
$APP_NAME = 'gotongroyong'
$ENV_FILE = Join-Path $PSScriptRoot '.env'

function Write-Step { param([string]$Msg) Write-Host "[GotongRoyong] $Msg" -ForegroundColor Cyan }
function Write-Err  { param([string]$Msg) Write-Host "[GotongRoyong] GAGAL: $Msg" -ForegroundColor Red }

function Test-NodeVersion {
    $v = (node -v 2>$null) -replace 'v', ''
    if (-not $v) {
        throw 'Node.js tidak ditemukan. Install Node.js 22.5+ dari https://nodejs.org lalu ulangi.'
    }
    $num = [version]$v
    if ($num -lt [version]'22.5.0') {
        throw "Node.js versi $v terlalu lama. Butuh minimal 22.5.0. Upgrade dulu."
    }
    Write-Step "Node.js v$v terdeteksi (OK)."
}

function Test-Pm2 {
    cmd /c "pm2 -v" *> $null
    return $LASTEXITCODE -eq 0
}

function Ensure-Pm2 {
    if (Test-Pm2) {
        Write-Step 'PM2 sudah terinstall.'
    } else {
        Write-Step 'Menginstall PM2 secara global...'
        $code = Invoke-Cmd @($NPM, 'install', '-g', 'pm2')
        if ($code -ne 0) { throw 'Gagal install PM2. Coba: npm install -g pm2' }
        if (-not (Test-Pm2)) { throw 'PM2 terinstall tapi tidak terdeteksi. Restart terminal lalu coba lagi.' }
    }
}

function Install-Deps {
    $nodeModules = Join-Path $PSScriptRoot 'node_modules'
    if (Test-Path -LiteralPath $nodeModules) {
        Write-Step 'node_modules sudah ada, lewati install.'
    } else {
        Write-Step 'Menginstall dependencies (npm install)...'
        $code = Invoke-Cmd @($NPM, 'install', '--omit=dev')
        if ($code -ne 0) { throw 'npm install gagal. Periksa koneksi internet.' }
        Write-Step 'Dependencies selesai terinstall.'
    }
}

function Save-Env {
    param([string]$Pw, [string]$Pt)
    if (-not $Pw -and -not $Pt) { return }
    $lines = @()
    if (Test-Path -LiteralPath $ENV_FILE) {
        $lines = @(Get-Content -LiteralPath $ENV_FILE)
    }
    $setLine = {
        param($arr, $key, $val)
        $idx = $null
        for ($i = 0; $i -lt $arr.Count; $i++) {
            if ($arr[$i] -match "^$key=") { $idx = $i; break }
        }
        if ($null -eq $idx) { $arr += "$key=$val" }
        else { $arr[$idx] = "$key=$val" }
        return $arr
    }
    if ($Pw) { $lines = & $setLine $lines 'ADMIN_PASSWORD' $Pw }
    if ($Pt) { $lines = & $setLine $lines 'PORT' $Pt }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($ENV_FILE, [string[]]$lines, $utf8NoBom)
    Write-Step "Variabel disimpan ke .env ($ENV_FILE)"
}

function Invoke-Cmd {
    param([string[]]$Argv)
    $line = ($Argv | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    & cmd /c $line | ForEach-Object { Write-Host $_ }
    return $LASTEXITCODE
}

function Pm2 {
    param([string[]]$Com)
    return Invoke-Cmd (@('pm2') + $Com)
}

function Start-App {
    Write-Step 'Menjalankan aplikasi via PM2...'
    $code = Pm2 'start' 'ecosystem.config.js'
    if ($code -ne 0) { throw 'pm2 start gagal.' }
    Pm2 'save' | Out-Null
    $port = '3000'
    if (Test-Path -LiteralPath $ENV_FILE) {
        $port = (Select-String -Path $ENV_FILE -Pattern '^PORT=' | Select-Object -First 1).Line -replace '^PORT=', ''
        if (-not $port) { $port = '3000' }
    }
    Write-Step 'Aplikasi berjalan!'
    Write-Host "    Halaman publik:  http://localhost:$port/"           -ForegroundColor Green
    Write-Host "    Halaman admin:   http://localhost:$port/admin.html" -ForegroundColor Green
    Write-Host '    Untuk lihat log:  .\deploy.ps1 -Logs'               -ForegroundColor Gray
}

# ---------- Perintah khusus ----------
if ($Stop) {
    Pm2 'stop' $APP_NAME
    & cmd /c "pm2 delete $APP_NAME" *> $null
    Write-Step 'Aplikasi berhenti.'
    return
}
if ($Restart) {
    if (-not (Test-Pm2)) { Write-Err 'PM2 belum terinstall. Jalankan .\deploy.ps1 tanpa flag dulu.'; return }
    Pm2 'restart' $APP_NAME
    Pm2 'logs' $APP_NAME '--lines' '20' '--nostream'
    return
}
if ($Status) {
    & cmd /c "pm2 status"
    return
}
if ($Logs) {
    if (-not (Test-Pm2)) { Write-Err 'PM2 belum terinstall.'; return }
    Pm2 'logs' $APP_NAME '--lines' '50' '--nostream'
    return
}

# ---------- Deploy utama ----------
try {
    if (-not ($Start -or $InstallDeps)) { Write-Step 'Mode deploy lengkap (install dependency + start).' }
    Test-NodeVersion
    Save-Env -Pw $AdminPassword -Pt $Port
    Ensure-Pm2
    Install-Deps
    Start-App
} catch {
    Write-Err $_.Exception.Message
    exit 1
}