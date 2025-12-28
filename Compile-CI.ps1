# Compile-CI.ps1 - Non-interactive compilation for GitHub Actions

$ScriptRoot = $PSScriptRoot
$BuildDir = "$ScriptRoot\Build"
$TempDir = "$BuildDir\Temp"
$OutputEXE = "$BuildDir\KMS_Activator.exe"
$LauncherCS = "$ScriptRoot\src\Launcher.cs"
$IconFile = "$ScriptRoot\icon.ico"

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  MS KMS Activator - CI Build" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# Clean build directory
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
if (Test-Path $OutputEXE) { Remove-Item $OutputEXE -Force }
New-Item -ItemType Directory -Path $BuildDir -Force | Out-Null
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

# Step 1: Check icon
Write-Host "[1/5] Checking icon..." -ForegroundColor Yellow

if (-not (Test-Path $IconFile)) {
    $IconFile = "$ScriptRoot\assets\key.ico"
    if (-not (Test-Path $IconFile)) {
        Write-Host "  WARNING: icon.ico not found, skipping icon" -ForegroundColor Yellow
        $IconFile = $null
    } else {
        Write-Host "  Using fallback icon: assets\key.ico" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Icon ready: icon.ico" -ForegroundColor Green
}

# Step 2: Copy application files
Write-Host "`n[2/5] Copying application files..." -ForegroundColor Yellow
Copy-Item "$ScriptRoot\KMS_Activator_GUI.ps1" -Destination $TempDir
Copy-Item "$ScriptRoot\Resources" -Destination "$TempDir\Resources" -Recurse
Copy-Item "$ScriptRoot\Modules" -Destination "$TempDir\Modules" -Recurse
Copy-Item "$ScriptRoot\UI" -Destination "$TempDir\UI" -Recurse
Copy-Item "$ScriptRoot\locales" -Destination "$TempDir\locales" -Recurse
Copy-Item "$ScriptRoot\assets" -Destination "$TempDir\assets" -Recurse
Write-Host "  Files copied" -ForegroundColor Green

# Step 3: Create payload ZIP
Write-Host "`n[3/5] Creating payload ZIP..." -ForegroundColor Yellow
$zipFile = "$BuildDir\payload.zip"
if (Test-Path $zipFile) { Remove-Item $zipFile -Force }
Compress-Archive -Path "$TempDir\*" -DestinationPath $zipFile -CompressionLevel Optimal
$zipSize = [math]::Round((Get-Item $zipFile).Length / 1KB, 0)
Write-Host "  ZIP created: $zipSize KB" -ForegroundColor Green

# Step 4: Find C# compiler
Write-Host "`n[4/5] Finding C# compiler..." -ForegroundColor Yellow
$csc = Get-ChildItem "C:\Windows\Microsoft.NET\Framework64" -Recurse -Filter "csc.exe" -ErrorAction SilentlyContinue | 
    Where-Object { $_.FullName -match "v4\." } | 
    Select-Object -First 1 -ExpandProperty FullName

if (-not $csc) {
    $csc = Get-ChildItem "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter "csc.exe" -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}

if (-not $csc) {
    Write-Host "  ERROR: C# compiler not found!" -ForegroundColor Red
    exit 1
}

Write-Host "  Compiler: $csc" -ForegroundColor Green

# Step 5: Compile
Write-Host "`n[5/5] Compiling to EXE..." -ForegroundColor Yellow

$compileArgs = @(
    "/target:winexe"
    "/out:$OutputEXE"
    "/resource:$zipFile,KMSActivator.payload.zip"
    "/reference:System.IO.Compression.dll"
    "/reference:System.IO.Compression.FileSystem.dll"
    "/reference:System.Windows.Forms.dll"
    "/reference:System.Drawing.dll"
    "/platform:x64"
    "/optimize+"
    $LauncherCS
)

# Add icon if available
if ($IconFile) {
    $compileArgs = @("/win32icon:$IconFile", "/resource:$IconFile,KMSActivator.app.ico") + $compileArgs
}

& $csc $compileArgs 2>&1 | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Host "  Compilation FAILED!" -ForegroundColor Red
    exit 1
}

Write-Host "  EXE compiled: $([math]::Round((Get-Item $OutputEXE).Length / 1MB, 2)) MB" -ForegroundColor Green

# Step 6: Code Signing (if certificate provided)
if ($env:CERT_PATH -and $env:CERT_PASSWORD) {
    Write-Host "`n[6/6] Code Signing..." -ForegroundColor Cyan
    
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($env:CERT_PATH, $env:CERT_PASSWORD)
        
        Write-Host "  Certificate: $($cert.Subject)" -ForegroundColor Green
        Write-Host "  Signing with timestamp..." -ForegroundColor Gray
        
        $signature = Set-AuthenticodeSignature -FilePath $OutputEXE -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
        
        if ($signature.Status -eq "Valid") {
            Write-Host "  Signed successfully!" -ForegroundColor Green
        } else {
            Write-Host "  WARNING: Signature status: $($signature.Status)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "`n[6/6] Code Signing skipped (no certificate)" -ForegroundColor Gray
}

# Cleanup
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $zipFile -Force -ErrorAction SilentlyContinue

# Generate release info
$hash = (Get-FileHash $OutputEXE -Algorithm SHA256).Hash
$isSigned = (Get-AuthenticodeSignature $OutputEXE).Status -eq "Valid"
$size = [math]::Round((Get-Item $OutputEXE).Length / 1MB, 2)

$releaseInfo = @"
MS KMS Activator - Build Info
========================================

File    : KMS_Activator.exe
Size    : $size MB
Signed  : $isSigned
Icon    : $(if ($IconFile) { 'Yes' } else { 'No' })
SHA256  : $hash

"@

if ($isSigned) {
    $sig = Get-AuthenticodeSignature $OutputEXE
    $releaseInfo += @"
Certificate:
  Subject : $($sig.SignerCertificate.Subject)
  Issuer  : $($sig.SignerCertificate.Issuer)
  Valid Until : $($sig.SignerCertificate.NotAfter)

"@
}

$releaseInfo += @"
Built   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
By      : 50bvd (GitHub Actions)
Platform: Windows x64
"@

$releaseInfo | Out-File -FilePath "$BuildDir\RELEASE_INFO.txt" -Encoding UTF8

# Success message
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  BUILD SUCCESS!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Output : $OutputEXE" -ForegroundColor White
Write-Host "Size   : $size MB" -ForegroundColor White
Write-Host "Icon   : $(if ($IconFile) { 'Yes' } else { 'No' })" -ForegroundColor $(if ($IconFile) { 'Green' } else { 'Yellow' })
Write-Host "Signed : $(if ($isSigned) { 'Yes' } else { 'No' })" -ForegroundColor $(if ($isSigned) { 'Green' } else { 'Yellow' })
Write-Host "SHA256 : $hash" -ForegroundColor White
Write-Host "`nReady for release!" -ForegroundColor Green
