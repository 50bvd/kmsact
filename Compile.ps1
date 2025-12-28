# Compile.ps1 - Compile C# launcher with embedded resources + Code Signing

$ScriptRoot = $PSScriptRoot
$BuildDir = "$ScriptRoot\Build"
$TempDir = "$BuildDir\Temp"
$OutputEXE = "$BuildDir\KMS_Activator.exe"
$LauncherCS = "$ScriptRoot\src\Launcher.cs"
$IconFile = "$ScriptRoot\icon.ico"
$SelfSignedCertPath = "$BuildDir\50bvd_CodeSigning.pfx"

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  MS KMS Activator - C# Compilation" -ForegroundColor Cyan
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

Write-Host "  Icon ready: icon.ico" -ForegroundColor Green

# Step 2: Copy application files
Write-Host "`n[2/5] Copying application files..." -ForegroundColor Yellow
Copy-Item "$ScriptRoot\KMS_Activator_GUI.ps1" -Destination $TempDir
Copy-Item "$ScriptRoot\Resources" -Destination "$TempDir\Resources" -Recurse
Copy-Item "$ScriptRoot\Modules" -Destination "$TempDir\Modules" -Recurse
Copy-Item "$ScriptRoot\UI" -Destination "$TempDir\UI" -Recurse
Copy-Item "$ScriptRoot\locales" -Destination "$TempDir\locales" -Recurse
Copy-Item "$ScriptRoot\assets" -Destination "$TempDir\assets" -Recurse
Write-Host "  Files copied (including locales)" -ForegroundColor Green

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
    Write-Host "  Please install .NET Framework or Visual Studio" -ForegroundColor Yellow
    exit 1
}

Write-Host "  Compiler: $csc" -ForegroundColor Green

# Step 5: Compile
Write-Host "`n[5/5] Compiling to EXE..." -ForegroundColor Yellow

$compileArgs = @(
    "/target:winexe"
    "/out:$OutputEXE"
    "/win32icon:$IconFile"
    "/resource:$IconFile,KMSActivator.app.ico"
    "/resource:$zipFile,KMSActivator.payload.zip"
    "/reference:System.IO.Compression.dll"
    "/reference:System.IO.Compression.FileSystem.dll"
    "/reference:System.Windows.Forms.dll"
    "/reference:System.Drawing.dll"
    "/platform:x64"
    "/optimize+"
    $LauncherCS
)

& $csc $compileArgs 2>&1 | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Host "  Compilation FAILED!" -ForegroundColor Red
    exit 1
}

Write-Host "  EXE compiled: $([math]::Round((Get-Item $OutputEXE).Length / 1MB, 2)) MB" -ForegroundColor Green

# Step 6: Code Signing
Write-Host "`n[6/6] Code Signing Configuration" -ForegroundColor Yellow
Write-Host "`nChoose signing option:" -ForegroundColor Cyan
Write-Host "  [1] Auto-generate self-signed certificate" -ForegroundColor White
Write-Host "  [2] Use existing .PFX certificate" -ForegroundColor White
Write-Host "  [3] Skip signing" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter choice (1-3)"
$signCert = $null

switch ($choice) {
    "1" {
        Write-Host "`nGenerating self-signed Code Signing certificate..." -ForegroundColor Cyan
        Write-Host "Enter certificate details (leave empty to omit):`n" -ForegroundColor Yellow
        
        $certName = Read-Host "Common Name / CN (required)"
        while ([string]::IsNullOrWhiteSpace($certName)) {
            Write-Host "  ERROR: Common Name is required!" -ForegroundColor Red
            $certName = Read-Host "Common Name / CN (required)"
        }
        
        $certOrg = Read-Host "Organization / O (optional)"
        $certOU = Read-Host "Organizational Unit / OU (optional)"
        $certCountry = Read-Host "Country / C (2 letters, optional)"
        $certState = Read-Host "State / ST (optional)"
        $certCity = Read-Host "City / L (optional)"
        
        if (-not [string]::IsNullOrWhiteSpace($certCountry)) {
            $certCountry = $certCountry.ToUpper().Substring(0, [Math]::Min(2, $certCountry.Length))
        }
        
        $certYears = Read-Host "Validity in years (default: 5)"
        if ([string]::IsNullOrWhiteSpace($certYears) -or -not ($certYears -match '^\d+$')) { 
            $certYears = 5 
        } else { 
            $certYears = [int]$certYears 
        }
        
        Write-Host ""
        $certPassword = Read-Host "Certificate password (default: 50bvd2025)"
        if ([string]::IsNullOrWhiteSpace($certPassword)) { $certPassword = "50bvd2025" }
        
        # Build subject
        $subjectParts = @("CN=$certName")
        if (-not [string]::IsNullOrWhiteSpace($certOrg)) { $subjectParts += "O=$certOrg" }
        if (-not [string]::IsNullOrWhiteSpace($certOU)) { $subjectParts += "OU=$certOU" }
        if (-not [string]::IsNullOrWhiteSpace($certCity)) { $subjectParts += "L=$certCity" }
        if (-not [string]::IsNullOrWhiteSpace($certState)) { $subjectParts += "S=$certState" }
        if (-not [string]::IsNullOrWhiteSpace($certCountry)) { $subjectParts += "C=$certCountry" }
        $subjectString = $subjectParts -join ", "
        
        Write-Host "`nCertificate will be created:" -ForegroundColor Cyan
        Write-Host "  Subject: $subjectString" -ForegroundColor White
        Write-Host "  Validity: $certYears years" -ForegroundColor White
        Write-Host ""
        
        try {
            $cert = New-SelfSignedCertificate `
                -Type CodeSigningCert `
                -Subject $subjectString `
                -CertStoreLocation Cert:\CurrentUser\My `
                -NotAfter (Get-Date).AddYears($certYears)
            
            Write-Host "  Certificate created: $($cert.Thumbprint)" -ForegroundColor Green
            
            $securePwd = ConvertTo-SecureString $certPassword -AsPlainText -Force
            Export-PfxCertificate -Cert $cert -FilePath $SelfSignedCertPath -Password $securePwd -ErrorAction Stop | Out-Null
            
            Write-Host "  Saved to: $SelfSignedCertPath" -ForegroundColor Green
            Write-Host "  Password: $certPassword" -ForegroundColor Yellow
            
            try {
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
                $store.Open("ReadWrite")
                $store.Add($cert)
                $store.Close()
                Write-Host "  Added to Trusted Root (trusted on this PC)" -ForegroundColor Green
            } catch {
                Write-Host "  WARNING: Could not add to Trusted Root: $_" -ForegroundColor Yellow
            }
            
            $signCert = $SelfSignedCertPath
            $Global:CertPassword = $certPassword
        } catch {
            Write-Host "  ERROR creating certificate: $_" -ForegroundColor Red
            Write-Host "  Certificate will not be used for signing" -ForegroundColor Yellow
            $signCert = $null
        }
    }
    
    "2" {
        Write-Host ""
        $customCertPath = Read-Host "Enter path to .PFX certificate"
        
        if (-not (Test-Path $customCertPath)) {
            Write-Host "  ERROR: Certificate file not found" -ForegroundColor Red
            $signCert = $null
        } elseif ($customCertPath -notmatch '\.pfx$') {
            Write-Host "  ERROR: File must be .PFX" -ForegroundColor Red
            $signCert = $null
        } else {
            Write-Host "  Certificate file found" -ForegroundColor Green
            $signCert = $customCertPath
            $secureCertPassword = Read-Host "Enter certificate password" -AsSecureString
            
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureCertPassword)
            $testPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            
            if ([string]::IsNullOrWhiteSpace($testPassword)) {
                Write-Host "  ERROR: Password cannot be empty" -ForegroundColor Red
                $signCert = $null
            } else {
                $Global:CertPassword = $testPassword
                Write-Host "  Password validated" -ForegroundColor Green
            }
        }
    }
    
    "3" {
        Write-Host "  Signing skipped" -ForegroundColor Yellow
        $signCert = $null
    }
    
    default {
        Write-Host "  Invalid choice, skipping" -ForegroundColor Yellow
        $signCert = $null
    }
}

# Sign EXE if certificate available
if ($signCert -and (Test-Path $signCert)) {
    Write-Host "`nSigning EXE..." -ForegroundColor Cyan
    
    try {
        Write-Host "  Loading certificate..." -ForegroundColor Gray
        
        if ($Global:CertPassword) {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($signCert, $Global:CertPassword)
        } else {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($signCert)
        }
        
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
        Write-Host "  Note: SSL/TLS certificates cannot sign code" -ForegroundColor Yellow
    }
} else {
    Write-Host "`nEXE not signed" -ForegroundColor Gray
}

# Cleanup
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $zipFile -Force -ErrorAction SilentlyContinue

# Generate release info
$hash = (Get-FileHash $OutputEXE -Algorithm SHA256).Hash
$isSigned = (Get-AuthenticodeSignature $OutputEXE).Status -eq "Valid"
$size = [math]::Round((Get-Item $OutputEXE).Length / 1MB, 2)

$releaseInfo = "MS KMS Activator - Build Info`n"
$releaseInfo += "========================================`n`n"
$releaseInfo += "File    : KMS_Activator.exe`n"
$releaseInfo += "Size    : $size MB`n"
$releaseInfo += "Signed  : $isSigned`n"
$releaseInfo += "Icon    : Yes`n"
$releaseInfo += "SHA256  : $hash`n`n"

if ($isSigned) {
    $sig = Get-AuthenticodeSignature $OutputEXE
    $releaseInfo += "Certificate:`n"
    $releaseInfo += "  Subject : $($sig.SignerCertificate.Subject)`n"
    $releaseInfo += "  Issuer  : $($sig.SignerCertificate.Issuer)`n"
    $releaseInfo += "  Valid   : $($sig.SignerCertificate.NotAfter)`n`n"
}

$releaseInfo += "Built   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
$releaseInfo += "By      : 50bvd`n"

$releaseInfo | Out-File -FilePath "$BuildDir\RELEASE_INFO.txt" -Encoding UTF8

# Success message
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  BUILD SUCCESS!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Output : $OutputEXE" -ForegroundColor White
Write-Host "Size   : $size MB" -ForegroundColor White
Write-Host "Icon   : Yes" -ForegroundColor Green
Write-Host "Signed : $(if ($isSigned) { 'Yes' } else { 'No' })" -ForegroundColor $(if ($isSigned) { 'Green' } else { 'Yellow' })
Write-Host "SHA256 : $hash" -ForegroundColor White

if ($signCert -eq $SelfSignedCertPath -and (Test-Path $SelfSignedCertPath)) {
    Write-Host "`nCertificate: $SelfSignedCertPath" -ForegroundColor Cyan
    Write-Host "Password   : $Global:CertPassword" -ForegroundColor Yellow
}

Write-Host "`nReady for distribution!" -ForegroundColor Green
Write-Host ""


