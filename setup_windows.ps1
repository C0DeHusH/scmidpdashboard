$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Quote-DotEnv([string]$Value) {
    return "'" + ($Value -replace "'", "\\'") + "'"
}

Write-Host "SCM IDP Dashboard v2.46.1 - Professional Audit + Reliability Setup" -ForegroundColor Cyan
Write-Host "No Supabase account or key is required." -ForegroundColor DarkGray
Write-Host ""

$adminPassword = Read-Host "Choose SCM Admin password"
if ([string]::IsNullOrWhiteSpace($adminPassword)) { throw "SCM Admin password is required." }

$secret = python -c "import secrets; print(secrets.token_hex(32))"
if ([string]::IsNullOrWhiteSpace($secret)) { throw "Unable to generate SCM_SECRET_KEY. Verify Python is installed." }

$lines = @(
    "SCM_SECRET_KEY=$(Quote-DotEnv $secret)",
    "SCM_ADMIN_PASSWORD=$(Quote-DotEnv $adminPassword)",
    "SCM_UTC_OFFSET_HOURS='8'",
    "SCM_HOST='127.0.0.1'",
    "SCM_SERVER_THREADS='8'"
)
$lines | Set-Content -Encoding UTF8 .env

Write-Host ""
Write-Host ".env created successfully." -ForegroundColor Green
Write-Host "Saved imports and planning state will be stored locally under the uploads folder." -ForegroundColor Green
Write-Host "Start the system with: run_windows.bat" -ForegroundColor Cyan
