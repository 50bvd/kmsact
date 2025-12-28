# Run.ps1 - Execute KMS Activator (Dev Mode)

$ScriptPath = Join-Path $PSScriptRoot "KMS_Activator_GUI.ps1"

if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: KMS_Activator_GUI.ps1 not found!" -ForegroundColor Red
    pause
    exit 1
}

# Check admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Relaunching as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$ScriptPath`"" -Verb RunAs
    exit
}

# Run the script
Write-Host "Starting MS KMS Activator..." -ForegroundColor Green
& $ScriptPath
