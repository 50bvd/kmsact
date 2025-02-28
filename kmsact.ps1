#Require admin rights
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "& {Set-ExecutionPolicy Bypass -Scope Process -Force; & '" + $myinvocation.mycommand.definition + "'}"
    Start-Process powershell.exe -Verb runAs -ArgumentList $arguments
    exit
}

# Global Configuration
$Global:Config = @{
    KMSServer = "50bvd.com"
    WindowsEditions = @{
        Pro = @{
            Name = "Pro"
            EditionId = "Professional"
            GenericKey = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"  # Windows 10/11 Pro Generic Key
            KMSKey = "W269N-WFGWX-YVC9B-4J6C9-T83GX"        # Pro KMS Client Key
        }
        Education = @{
            Name = "Education"
            EditionId = "Education"
            GenericKey = "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"    # Education Generic Key
            KMSKey = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"        # Education KMS Client Key
        }
        Enterprise = @{
            Name = "Enterprise"
            EditionId = "Enterprise"
            GenericKey = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"    # Enterprise Generic Key
            KMSKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"        # Enterprise KMS Client Key
        }
        Home = @{
            Name = "Home"
            EditionId = "Core"
            KMSKey = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"       # Home KMS Client Key
        }
    }
    Office = @{
        Key = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"              # Office 2021 LTSC Volume Key
        ODTUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe"
        Paths = @(
            "C:\Program Files\Microsoft Office\Office16",
            "C:\Program Files (x86)\Microsoft Office\Office16",
            "C:\Program Files\Microsoft Office\Office15",
            "C:\Program Files (x86)\Microsoft Office\Office15"
        )
        MainPath = "C:\Program Files\Microsoft Office\Office16"
    }
}

function Show-AsciiArt {
    Clear-Host
    Write-Host @"
  ___                       ___  
 (o o)                     (o o) 
(  V  ) MS KMS Activation (  V  )
--m-m-----------------------m-m--
    https://github.com/50bvd
"@ -ForegroundColor Green
}

# Monitors a process with a spinner animation and timeout
function Watch-Process {
    param(
        [Parameter(Mandatory=$true)]
        [System.Diagnostics.Process]$Process,
        [string]$ActivityName,
        [int]$TimeoutSeconds = 300
    )
    
    $spinChars = '/', '-', '\', '|'
    $currentSpinIdx = 0
    $startTime = Get-Date
    $processEnded = $false

    Write-Host "`n$ActivityName..." -ForegroundColor Yellow

    while (-not $processEnded) {
        Write-Host "`r$($spinChars[$currentSpinIdx])" -NoNewline -ForegroundColor Yellow
        $currentSpinIdx = ($currentSpinIdx + 1) % $spinChars.Length
        
        if ($Process.HasExited -or $null -eq (Get-Process -Id $Process.Id -ErrorAction SilentlyContinue)) {
            $processEnded = $true
        }
        else {
            Start-Sleep -Milliseconds 100
        }
        
        $elapsedTime = ((Get-Date) - $startTime).TotalSeconds
        if ($elapsedTime -gt $TimeoutSeconds) {
            Write-Host "`r[-] $ActivityName timed out after $TimeoutSeconds seconds." -ForegroundColor Red
            try { $Process.Kill() } catch {}
            return $false
        }
    }
    
    $elapsedTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
    
    if ($Process.ExitCode -eq 0) {
        Write-Host "`r[+] $ActivityName completed in $elapsedTime seconds." -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "`r[-] $ActivityName failed with exit code: $($Process.ExitCode)" -ForegroundColor Red
        return $false
    }
}

# Handles file downloads with progress bar
function Watch-Download {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$OutFile,
        [string]$ActivityName = "Downloading file"
    )
    
    try {
        $uri = New-Object System.Uri($Url)
        $request = [System.Net.HttpWebRequest]::Create($uri)
        $response = $request.GetResponse()
        $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
        $responseStream = $response.GetResponseStream()
        $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $OutFile, Create
        $buffer = New-Object byte[] 10KB
        $count = $responseStream.Read($buffer, 0, $buffer.length)
        $downloadedBytes = $count
        $previousProgress = 0
        
        Write-Host "`n$ActivityName..." -ForegroundColor Yellow
        
        while ($count -gt 0) {
            $targetStream.Write($buffer, 0, $count)
            $count = $responseStream.Read($buffer, 0, $buffer.length)
            $downloadedBytes += $count
            $currentProgress = [System.Math]::Floor(($downloadedBytes/1024) / $totalLength * 100)
            
            if ($currentProgress -gt $previousProgress) {
                $previousProgress = $currentProgress
                $progressBar = "[" + ("=" * [math]::Floor($currentProgress/2)) + ">" + (" " * (50 - [math]::Floor($currentProgress/2))) + "]"
                Write-Host "`r$progressBar $currentProgress% " -NoNewline -ForegroundColor Yellow
            }
        }
        
        Write-Host "`r[+] $ActivityName completed successfully." -ForegroundColor Green
        $targetStream.Flush()
        $targetStream.Close()
        $targetStream.Dispose()
        $responseStream.Dispose()
        return $true
    }
    catch {
        Write-Host "`r[-] $ActivityName failed: $_" -ForegroundColor Red
        return $false
    }
}

# Retrieves Windows version and edition details
function Get-WindowsInfo {
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $windowsInfo = @{
            ProductName       = "Unknown Windows Version"
            EditionID         = ""
            CurrentBuild      = ""
            UBR               = ""
            WindowsFamily     = "Unknown"
            Architecture      = if ([System.Environment]::Is64BitOperatingSystem) { "64-bit" } else { "32-bit" }
            FullName          = ""
            CurrentEdition    = $null
            AvailableEditions = @()
            KMSKey            = $null
            IsHome            = $false
        }
        $regValues = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($regValues) {
            $windowsInfo.ProductName = $regValues.ProductName
            $windowsInfo.EditionID = $regValues.EditionID
            $windowsInfo.CurrentBuild = $regValues.CurrentBuild
            $windowsInfo.UBR = $regValues.UBR
            $windowsInfo.FullName = $os.Caption
            # Detect Home edition
            $windowsInfo.IsHome = $regValues.EditionID -match "Core"
            if ([System.Environment]::OSVersion.Version.Build -ge 22000) {
                $windowsInfo.WindowsFamily = "Windows 11"
            }
            else {
                $windowsInfo.WindowsFamily = "Windows 10"
            }
            
            # Populate available editions from global config
            $windowsInfo.AvailableEditions = $Global:Config.WindowsEditions.Values | Where-Object { $_.Name -ne "Home" }
            
            foreach ($edition in $Global:Config.WindowsEditions.Values) {
                if ($edition.EditionId -eq $windowsInfo.EditionID) {
                    $windowsInfo.CurrentEdition = $edition
                    $windowsInfo.KMSKey = $edition.KMSKey
                    break
                }
            }
        }
        return $windowsInfo
    }
    catch {
        return $null
    }
}

# Activates Windows using KMS
function Enable-Windows {
    $windowsInfo = Get-WindowsInfo    
    Clear-Host
    Show-AsciiArt
    
    Write-Host "Activating $($windowsInfo.CurrentEdition.Name)..." -ForegroundColor Yellow

    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /ipk $($windowsInfo.KMSKey)" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Installing product key" | Out-Null
    
    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /skms $($Global:Config.KMSServer)" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Setting KMS server" | Out-Null

    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /ato" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Activating Windows" | Out-Null

    Start-Sleep -Seconds 5
    Write-Host "`nWindows has been successfully activated!" -ForegroundColor Green
}

# Changes Windows edition (e.g., Home → Pro)
function Set-WindowsEdition {
    $windowsInfo = Get-WindowsInfo
    Clear-Host
    Show-AsciiArt
    Write-Host ""
    Write-Host "Current Edition: $($windowsInfo.EditionID)" -ForegroundColor Yellow
    Write-Host ""

    # Upgrade Home to Pro if needed
    if ($windowsInfo.IsHome) {
        Write-Host "You are using Windows Home. To change edition, you must upgrade to Windows Pro." -ForegroundColor Yellow
        Write-Host ""
        $choice = Read-Host "Do you want to proceed with the upgrade to Windows Pro? (Y/N)"
        if ($choice -eq 'Y' -or $choice -eq 'y') {
            Write-Host "Upgrading to Windows Pro..." -ForegroundColor Yellow
            Write-Host ""
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c sc config licensemanager start= auto & net start licensemanager" -PassThru -WindowStyle Hidden
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c sc config wuauserv start= auto & net start wuauserv" -PassThru -WindowStyle Hidden
            Watch-Process -Process $process -ActivityName "Reconfiguring SCM service" | Out-Null
            Write-Host ""
            cmd.exe /c "changepk /ProductKey $($Global:Config.WindowsEditions.Pro.GenericKey)"
            exit
        }
        else {
            Write-Host "Operation canceled." -ForegroundColor Yellow
            return
        }
    }
    
    # Edition selection menu
    Write-Host "Select the desired edition:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   1. Windows Pro"
    Write-Host ""
    Write-Host "   2. Windows Education"
    Write-Host ""
    Write-Host "   3. Windows Enterprise"
    Write-Host ""
    Write-Host "   B. Back"
    Write-Host ""
    
    $choice = Read-Host "Enter your choice (1, 2, 3 or B)"
    
    if ($choice -eq 'B' -or $choice -eq 'b') {
        return
    }
    
    $editionMap = @{
        "1" = $Global:Config.WindowsEditions.Pro
        "2" = $Global:Config.WindowsEditions.Education
        "3" = $Global:Config.WindowsEditions.Enterprise
    }
    
    if (-not $editionMap.ContainsKey($choice)) {
        Write-Host "Invalid choice, operation canceled." -ForegroundColor Red
        return
    }
    
    $selectedEdition = $editionMap[$choice]
    Write-Host "Changing edition to Windows $($selectedEdition.Name)..." -ForegroundColor Yellow
    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /ipk $($selectedEdition.KMSKey)" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Installing product key" | Out-Null

    Write-Host "Product key changed." -ForegroundColor Green
    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /skms $($Global:Config.KMSServer)" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Setting KMS server" | Out-Null
    
    Write-Host "KMS server configured." -ForegroundColor Green
    $process = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo $env:windir\system32\slmgr.vbs /ato" -PassThru -WindowStyle Hidden
    Watch-Process -Process $process -ActivityName "Activating Windows" | Out-Null
    
    Write-Host "Windows edition changed successfully." -ForegroundColor Green
    return
}

# Activates Microsoft Office
function Enable-Office {
    $officeVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).VersionToReport
    if ($officeVersion) {
        Clear-Host
        Show-AsciiArt
        Write-Host ""
        Write-Host "Detected Office version: $officeVersion" -ForegroundColor Cyan
        
        $officePath = $null
        foreach ($path in $Global:Config.Office.Paths) {
            if (Test-Path $path) {
                $officePath = $path
                break
            }
        }
        
        if ($officePath) {
            Set-Location $officePath
            
            $keyProcess = Start-Process -FilePath "cscript.exe" -ArgumentList "ospp.vbs", "/inpkey:$($Global:Config.Office.Key)" -PassThru -WindowStyle Hidden
            Watch-Process -Process $keyProcess -ActivityName "Installing Office product key" | Out-Null
    
            $hostProcess = Start-Process -FilePath "cscript.exe" -ArgumentList "ospp.vbs", "/sethst:$($Global:Config.KMSServer)" -PassThru -WindowStyle Hidden
            Watch-Process -Process $hostProcess -ActivityName "Setting KMS host for Office" | Out-Null
    
            $actProcess = Start-Process -FilePath "cscript.exe" -ArgumentList "ospp.vbs", "/act" -PassThru -WindowStyle Hidden
            Watch-Process -Process $actProcess -ActivityName "Activating Office" | Out-Null

            Write-Host "`nOffice has been successfully activated!" -ForegroundColor Green

        }
        else {
            Write-Host "Office installation not found." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Office installation not detected." -ForegroundColor Red
    }
}

# Schedules monthly activation renewal
function New-ActivationSchedule {
    $taskName = "MS KMS Activation Renewal"
    $taskDescription = "Renews MS KMS Activation monthly"
    
    # Dynamically build Office activation command if installed
    $officeCmd = ""
    foreach ($path in $Global:Config.Office.Paths) {
        if (Test-Path $path) {
            $officeCmd = "; if (Test-Path '$path') { Set-Location '$path'; cscript ospp.vbs /act }"
            break
        }
    }
    
    $action = New-ScheduledTaskAction -Execute "cscript.exe" -Command "//nologo $env:windir\system32\slmgr.vbs /ato" "//nologo $officeCmd\ospp.vbs /act"
    $trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Monday -At 9am
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable
    
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Clear-Host
        Show-AsciiArt
        Write-Host ""
        Write-Host "Updating existing scheduled task..." -ForegroundColor Yellow
        Write-Host ""
        Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings
    }
    else {
        Clear-Host
        Show-AsciiArt
        Write-Host ""
        Write-Host "Creating new scheduled task..." -ForegroundColor Yellow
        Write-Host ""
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description $taskDescription -User "SYSTEM"
    }
    
    Write-Host "Scheduled task created/updated successfully." -ForegroundColor Green
    Write-Host "Task Name: $taskName" -ForegroundColor Cyan
}

# Installs Office LTSC 2021 with predefined settings
function Install-OfficeLTSC {
    Clear-Host
    Show-AsciiArt
    
    $tempFolder = New-Item -ItemType Directory -Path $env:TEMP -Name "OfficeLTSC_Install_$(Get-Random)"
    Set-Location $tempFolder

    try {
        $odtFile = Join-Path $tempFolder "ODT.exe"
        $downloadSuccess = Watch-Download -Url $Global:Config.Office.ODTUrl -OutFile $odtFile -ActivityName "Downloading Office Deployment Tool"
        
        if (-not $downloadSuccess) {
            throw "Failed to download Office Deployment Tool"
        }
        
        Write-Host ""
        Write-Host "Extracting Office Deployment Tool..." -ForegroundColor Yellow
        $extractProcess = Start-Process -FilePath $odtFile -ArgumentList "/extract:$tempFolder /quiet" -PassThru -WindowStyle Hidden -Wait
        if ($extractProcess.ExitCode -eq 0) {
            Write-Host "[+] Office Deployment Tool extracted successfully." -ForegroundColor Green
        }
        else {
            throw "ODT extraction failed with exit code: $($extractProcess.ExitCode)"
        }
        
        Start-Sleep -Seconds 2
        
        $setupPath = Join-Path $tempFolder "setup.exe"
        if (!(Test-Path $setupPath)) {
            throw "Setup.exe not found after extraction"
        }

        Write-Host "Integrating configuration file..." -ForegroundColor Yellow
        Write-Host ""
        $xmlContent = @'
<Configuration ID="66c343d1-8a0c-4347-bc31-863bad4959fc">
<Info Description=""/>
<Add OfficeClientEdition="64" Channel="PerpetualVL2021" MigrateArch="TRUE">
<Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
<Language ID="MatchOS"/>
<Language ID="MatchPreviousMSI"/>
</Product>
<Product ID="VisioPro2021Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">
<Language ID="MatchOS"/>
<Language ID="MatchPreviousMSI"/>
</Product>
<Product ID="ProjectPro2021Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">
<Language ID="MatchOS"/>
<Language ID="MatchPreviousMSI"/>
</Product>
<Product ID="AccessRuntimeRetail">
<Language ID="MatchOS"/>
<Language ID="MatchPreviousMSI"/>
</Product>
</Add>
<Property Name="SharedComputerLicensing" Value="0"/>
<Property Name="FORCEAPPSHUTDOWN" Value="FALSE"/>
<Property Name="DeviceBasedLicensing" Value="0"/>
<Property Name="SCLCacheOverride" Value="0"/>
<Property Name="AUTOACTIVATE" Value="1"/>
<Updates Enabled="TRUE"/>
<RemoveMSI/>
<AppSettings>
<Setup Name="Company" Value="Cracked by 50bvd"/>
</AppSettings>
<Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
'@
        $xmlFile = Join-Path $tempFolder "Office.xml"
        $xmlContent | Out-File -FilePath $xmlFile -Encoding UTF8

        $installProcess = Start-Process -FilePath $setupPath -ArgumentList "/configure", "`"$xmlFile`"" -PassThru -WindowStyle Hidden
        Watch-Process -Process $installProcess -ActivityName "Installing Office LTSC 2021" -TimeoutSeconds 600 | Out-Null
        Write-Host ""
        if ($installProcess.ExitCode -eq 0) {
            Write-Host "Office LTSC 2021 has been successfully installed." -ForegroundColor Green
            Write-Host ""
            Write-Host "Starting Office activation..." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
            Enable-Office
        }
        else {
            Write-Host "Office installation failed with error code: $($installProcess.ExitCode)" -ForegroundColor Red
            Write-Host ""
        }
    }
    catch {
        Write-Host "An error occurred during Office installation: $_" -ForegroundColor Red
        Write-Host ""
    }
    finally {
        Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
        Write-Host ""
        Set-Location $env:TEMP
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Miscellaneous tools menu
function Show-MiscMenu {
    Clear-Host
    Show-AsciiArt
    Write-Host ""
    Write-Host "Miscellaneous Menu:" -ForegroundColor Green
    Write-Host ""
    Write-Host "1. Uninstall activation keys" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "2. Manually set KMS server" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "3. Uninstall Office completely" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Q. Back to main menu" -ForegroundColor Yellow
    Write-Host ""
    $choice = Read-Host "Enter your choice"
    
    switch ($choice) {
        1 { Uninstall-ProductKey }
        2 { Select-KMSServer }
        3 { Uninstall-Office }
        q { return }
        default { Write-Host "Invalid choice, please try again." -ForegroundColor Red; pause; Show-MiscMenu }
    }
    pause
    Show-MiscMenu
}

# Removes Windows and Office activation keys
function Uninstall-ProductKey {
    Clear-Host
    Show-AsciiArt
    Write-Host ""
    Write-Host "Uninstalling Windows activation key..." -ForegroundColor Yellow
    $result = cscript //nologo $env:windir\system32\slmgr.vbs /upk
    Write-Host ""
    Write-Host $result -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Clearing product key from registry..." -ForegroundColor Yellow
    $result = cscript //nologo $env:windir\system32\slmgr.vbs /cpky
    Write-Host ""
    Write-Host $result -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Uninstalling Office activation keys..." -ForegroundColor Yellow
    
    $officeFound = $false
    foreach ($path in $Global:Config.Office.Paths) {
        if (Test-Path $path) {
            $officeFound = $true
            Set-Location $path
            Write-Host "Checking Office licenses at $path..." -ForegroundColor Yellow
            cscript ospp.vbs /dstatus
            Write-Host "Attempting to remove Office licenses..." -ForegroundColor Yellow
            $result = cscript ospp.vbs /rearm
            Write-Host $result -ForegroundColor Cyan
            Write-Host "Verifying final status..." -ForegroundColor Yellow
            $finalStatus = cscript ospp.vbs /dstatus
            Write-Host $finalStatus -ForegroundColor Cyan
        }
    }
    
    if (-not $officeFound) {
        Write-Host ""
        Write-Host "No Office installation detected." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Activation keys uninstallation completed." -ForegroundColor Green
    Write-Host "A system restart may be required for all changes to take effect." -ForegroundColor Yellow
}

# Updates KMS server address
function Select-KMSServer {
    Clear-Host
    Show-AsciiArt
    Write-Host ""
    Write-Host "Current KMS server: $($Global:Config.KMSServer)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Enter a new KMS server address (or press Enter to keep current):" -ForegroundColor Yellow
    $newServer = Read-Host
    if ($newServer -and $newServer -ne $Global:Config.KMSServer) {
        $Global:Config.KMSServer = $newServer
        Write-Host "Updating Windows KMS server..." -ForegroundColor Yellow
        cscript //nologo $env:windir\system32\slmgr.vbs /skms $Global:Config.KMSServer
        Write-Host ""
        
        $officeFound = $false
        foreach ($path in $Global:Config.Office.Paths) {
            if (Test-Path $path) {
                $officeFound = $true
                Write-Host "Updating Office KMS server for $path..." -ForegroundColor Yellow
                Set-Location $path
                cscript ospp.vbs /sethst:$Global:Config.KMSServer
                Write-Host ""
            }
        }
        
        if (-not $officeFound) {
            Write-Host "No Office installation detected." -ForegroundColor Yellow
        }
        Write-Host "KMS server updated to: $($Global:Config.KMSServer)" -ForegroundColor Green
    }
}

# Uninstalls Office completely
function Uninstall-Office {
    Clear-Host
    Show-AsciiArt
    Write-Host ""
    Write-Host "Starting Office uninstallation process..." -ForegroundColor Yellow
    Write-Host ""
    
    $tempFolder = New-Item -ItemType Directory -Path $env:TEMP -Name "OfficeUninstall_$(Get-Random)"
    Set-Location $tempFolder

    try {
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe"
        $odtFile = Join-Path $tempFolder "ODT.exe"
        Watch-Download -Url $odtUrl -OutFile $odtFile -ActivityName "Downloading Office Deployment Tool" | Out-Null
        Write-Host ""
        $configPath = Join-Path $tempFolder "configuration.xml"
        $configXml = @"
<Configuration>
  <Remove All="True" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
        $configXml | Out-File -FilePath $configPath -Encoding UTF8
        Write-Host "Extracting Office Deployment Tool..." -ForegroundColor Yellow
        $extractProcess = Start-Process -FilePath $odtFile -ArgumentList "/extract:$tempFolder /quiet" -PassThru -WindowStyle Hidden -Wait
        if ($extractProcess.ExitCode -eq 0) {
            Write-Host "[+] Office Deployment Tool extracted successfully." -ForegroundColor Green
        }
        else {
            throw "ODT extraction failed with exit code: $($extractProcess.ExitCode)"
        }
        Start-Sleep -Seconds 2
        $setupPath = Join-Path $tempFolder "setup.exe"
        if (!(Test-Path $setupPath)) {
            throw "Setup.exe not found after extraction"
        }
        Write-Host "Starting Office removal..." -ForegroundColor Yellow
        $uninstallProcess = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -PassThru -WindowStyle Hidden
        Watch-Process -Process $uninstallProcess -ActivityName "Removing Office" | Out-Null
        Write-Host ""
        if ($uninstallProcess.ExitCode -eq 0) {
            Write-Host "Office has been successfully uninstalled." -ForegroundColor Green
            Write-Host ""
        }
        else {
            Write-Host "Office uninstallation failed with error code: $($uninstallProcess.ExitCode)" -ForegroundColor Red
            Write-Host ""
        }
    }
    catch {
        Write-Host "An error occurred during Office uninstallation: $_" -ForegroundColor Red
        Write-Host ""
    }
    finally {
        Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
        Write-Host ""
        Set-Location $env:TEMP
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Main menu
function Show-Menu {
    $windowsInfo = Get-WindowsInfo
    if ($null -eq $windowsInfo) {
        Write-Host "Unsupported Windows version. Only Windows 10 and 11 are supported." -ForegroundColor Red
        pause
        exit
    }

    Show-AsciiArt
    Write-Host ""
    Write-Host "Detected System:" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "- Version: $($windowsInfo.WindowsFamily)" -ForegroundColor Cyan
    Write-Host "- Full Name: $($windowsInfo.FullName)" -ForegroundColor Cyan
    Write-Host "- Edition: $($windowsInfo.EditionID)" -ForegroundColor Cyan
    Write-Host "- Build: $($windowsInfo.CurrentBuild).$($windowsInfo.UBR)" -ForegroundColor Cyan
    Write-Host "- Architecture: $($windowsInfo.Architecture)" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "- 1. Activate Windows" -ForegroundColor Yellow
    Write-Host ""
    if (Test-Path "C:\Program Files\Microsoft Office\Office16") {
        Write-Host "- 2. Activate Office" -ForegroundColor Yellow
        Write-Host ""
    }
    else {
        Write-Host "- 2. Activate Office [NOT INSTALLED]" -ForegroundColor Yellow
        Write-Host ""
    }
    
    Write-Host "- 3. Change Windows Edition" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "- 4. Schedule Activation Task" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "- 5. Install & Activate Office LTSC 2021" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "- 6. Miscellaneous Tools" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "- Q. Exit" -ForegroundColor Yellow
    Write-Host ""
    
    $choice = Read-Host "Enter your choice"
    switch ($choice) {
        "1" { Enable-Windows }
        "2" { 
            if (Test-Path "C:\Program Files\Microsoft Office\Office16") {
                Enable-Office 
            }
            else {
                Write-Host "Office is not installed. Would you like to install it? (Y/N)" -ForegroundColor Yellow
                $response = Read-Host
                if ($response -eq 'Y' -or $response -eq 'y') {
                    Install-OfficeLTSC
                }
            }
        }
        "3" { Set-WindowsEdition }
        "4" { New-ActivationSchedule }
        "5" { Install-OfficeLTSC }
        "6" { Show-MiscMenu }
        "q" { exit }
        default { Write-Host "Invalid choice. Please try again." -ForegroundColor Red }
    }
    
    pause
    Show-Menu
}

Show-Menu

Write-Host "Press any key to exit..."
pause
exit
