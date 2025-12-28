# Activation Core Module - Clean version with WPF MessageBox only

function Invoke-CommandWithRealTimeOutput {
    param(
        [string]$FilePath,
        [string[]]$ArgumentList,
        [string]$WorkingDirectory = $null
    )
    
    $cmdLine = $ArgumentList -join " "
    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "cmd.exe"
    $psi.Arguments = "/c `"chcp 65001 >nul && cd /d `"$WorkingDirectory`" && $FilePath $cmdLine`""
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    
    $outputHandler = {
        if ($EventArgs.Data) {
            $line = $EventArgs.Data
            try {
                $Global:LogBox.Dispatcher.Invoke([action]{
                    $Global:LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] [INFO] $line`n")
                    $Global:LogBox.ScrollToEnd()
                }, "Normal")
                [System.Windows.Forms.Application]::DoEvents()
            } catch {}
        }
    }
    
    $errorHandler = {
        if ($EventArgs.Data) {
            $line = $EventArgs.Data
            try {
                $Global:LogBox.Dispatcher.Invoke([action]{
                    $Global:LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] [WARN] $line`n")
                    $Global:LogBox.ScrollToEnd()
                }, "Normal")
                [System.Windows.Forms.Application]::DoEvents()
            } catch {}
        }
    }
    
    Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action $outputHandler | Out-Null
    Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action $errorHandler | Out-Null
    
    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    
    while (-not $process.HasExited) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    
    Get-EventSubscriber | Where-Object { $_.SourceObject -eq $process } | Unregister-Event
    
    return $exitCode
}

function Enable-Windows {
    $windowsInfo = Get-WindowsInfo
    
    Write-LogHeader "Windows Activation"
    Update-Status "Activating Windows..."
    
    Write-LogStep "Windows Edition: $($windowsInfo.CurrentEdition.Name)" "INFO"
    Write-LogStep "Product Key: $($windowsInfo.KMSKey)" "INFO"
    Write-LogStep "KMS Server: $($Global:Config.KMSServer)" "INFO"
    
    try {
        Write-LogStep "Installing product key..." "INFO"
        $slmgrPath = "$env:windir\system32\slmgr.vbs"
        
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/ipk", $windowsInfo.KMSKey) -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep "Product key installed successfully" "SUCCESS"
        } else {
            Write-LogStep "Product key installation returned code: $exitCode" "WARNING"
        }
        
        Start-Sleep -Seconds 1
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-LogStep "Configuring KMS server..." "INFO"
        
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/skms", $Global:Config.KMSServer) -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep "KMS server configured successfully" "SUCCESS"
        } else {
            Write-LogStep "KMS server configuration returned code: $exitCode" "WARNING"
        }
        
        Start-Sleep -Seconds 1
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-LogStep "Activating Windows..." "INFO"
        
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/ato") -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep "Activation command completed" "SUCCESS"
        } else {
            Write-LogStep "Activation returned code: $exitCode" "WARNING"
        }
        
        Start-Sleep -Seconds 1
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-LogHeader "Activation Complete"
        Write-LogStep "Windows has been activated!" "SUCCESS"
        
        Show-StyledMessageBox -Message (Get-String "windowsActivated") -Title (Get-String "success") -Buttons "OK" -Icon "Information"
    }
    catch {
        Write-LogStep "Exception: $($_.Exception.Message)" "ERROR"
        Write-LogStep "Stack trace: $($_.ScriptStackTrace)" "ERROR"
        
        Show-StyledMessageBox -Message "Activation error: $_" -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

function Enable-Office {
    Write-LogHeader "Office Activation"
    Update-Status "Activating Office..."
    
    try {
        Write-LogStep "Checking Office installation..." "INFO"
        $officeVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).VersionToReport
        
        if ($officeVersion) {
            Write-LogStep "Office version detected: $officeVersion" "SUCCESS"
            
            $officePath = $null
            Write-LogStep "Searching for Office installation..." "INFO"
            
            foreach ($path in $Global:Config.Office.Paths) {
                if (Test-Path $path) {
                    $officePath = $path
                    Write-LogStep "Office found at: $path" "SUCCESS"
                    break
                }
            }
            
            if ($officePath) {
                $osppPath = Join-Path $officePath "ospp.vbs"
                if (Test-Path $osppPath) {
                    Write-LogStep "ospp.vbs found" "SUCCESS"
                } else {
                    throw "ospp.vbs not found in $officePath"
                }
                
                Write-LogStep "Installing Office product key..." "INFO"
                
                $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$osppPath`"", "/inpkey:$($Global:Config.Office.Key)") -WorkingDirectory $officePath
                
                if ($exitCode -eq 0) {
                    Write-LogStep "Product key installed successfully" "SUCCESS"
                } else {
                    Write-LogStep "Product key installation returned code: $exitCode" "WARNING"
                }
                
                Start-Sleep -Seconds 1
                [System.Windows.Forms.Application]::DoEvents()
                
                Write-LogStep "Configuring KMS host..." "INFO"
                
                $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$osppPath`"", "/sethst:$($Global:Config.KMSServer)") -WorkingDirectory $officePath
                
                if ($exitCode -eq 0) {
                    Write-LogStep "KMS host configured successfully" "SUCCESS"
                } else {
                    Write-LogStep "KMS host configuration returned code: $exitCode" "WARNING"
                }
                
                Start-Sleep -Seconds 1
                [System.Windows.Forms.Application]::DoEvents()
                
                Write-LogStep "Activating Office..." "INFO"
                
                $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$osppPath`"", "/act") -WorkingDirectory $officePath
                
                if ($exitCode -eq 0) {
                    Write-LogStep "Activation command completed" "SUCCESS"
                } else {
                    Write-LogStep "Activation returned code: $exitCode" "WARNING"
                }
                
                Write-LogHeader "Activation Complete"
                Write-LogStep "Office has been activated!" "SUCCESS"
                
                Show-StyledMessageBox -Message (Get-String "officeActivated") -Title (Get-String "success") -Buttons "OK" -Icon "Information"
            }
            else {
                Write-LogStep "Office installation path not found" "ERROR"
                Write-LogStep "Searched paths:" "ERROR"
                foreach ($path in $Global:Config.Office.Paths) {
                    Write-LogStep "  - $path" "ERROR"
                }
                
                Show-StyledMessageBox -Message (Get-String "officeNotInstalled") -Title (Get-String "error") -Buttons "OK" -Icon "Error"
            }
        }
        else {
            Write-LogStep "Office not detected on system" "ERROR"
            Write-LogStep "Registry key not found: HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" "ERROR"
            
            Show-StyledMessageBox -Message (Get-String "officeNotInstalled") -Title (Get-String "error") -Buttons "OK" -Icon "Error"
        }
    }
    catch {
        Write-LogStep "Exception: $($_.Exception.Message)" "ERROR"
        Write-LogStep "Stack trace: $($_.ScriptStackTrace)" "ERROR"
        
        Show-StyledMessageBox -Message "Activation error: $_" -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

function Uninstall-ProductKey {
    Write-Host "DEBUG: Uninstall-ProductKey called" -ForegroundColor Magenta
    Write-LogHeader "Uninstalling Product Keys"
    Update-Status "Uninstalling keys..."
    
    try {
        Write-LogStep "Uninstalling Windows activation key..." "INFO"
        $slmgrPath = "$env:windir\system32\slmgr.vbs"
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/upk") -WorkingDirectory $env:windir
        
        Write-LogStep "Clearing product key from registry..." "INFO"
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/cpky") -WorkingDirectory $env:windir
        
        Write-LogStep "Uninstalling Office activation keys..." "INFO"
        $officeFound = $false
        
        foreach ($path in $Global:Config.Office.Paths) {
            if (Test-Path $path) {
                $officeFound = $true
                $osppPath = Join-Path $path "ospp.vbs"
                
                Write-LogStep "Rearming Office licenses at: $path" "INFO"
                $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("`"$osppPath`"", "/rearm") -WorkingDirectory $path
            }
        }
        
        if (-not $officeFound) {
            Write-LogStep "No Office installation detected" "WARNING"
        }
        
        Write-LogStep "Product keys uninstallation completed" "SUCCESS"
        Write-LogStep "A system restart may be required" "WARNING"
        
        Show-StyledMessageBox -Message (Get-String "productKeysUninstalled") -Title (Get-String "success") -Buttons "OK" -Icon "Information"
    }
    catch {
        Write-LogStep "Exception: $($_.Exception.Message)" "ERROR"
        
        Show-StyledMessageBox -Message "Uninstall error: $_" -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

function New-ActivationSchedule {
    Write-LogHeader "Scheduling Auto-Renewal"
    Update-Status "Creating scheduled task..."
    
    try {
        $taskName = "MS KMS Activation Renewal"
        $taskDescription = "Renews MS KMS Activation monthly"
        
        Write-LogStep "Task name: $taskName" "INFO"
        Write-LogStep "Schedule: Every 4 weeks on Monday at 9:00 AM" "INFO"
        
        $officeCmd = ""
        foreach ($path in $Global:Config.Office.Paths) {
            if (Test-Path $path) {
                $officeCmd = "; if (Test-Path '$path') { Set-Location '$path'; cscript ospp.vbs /act }"
                Write-LogStep "Office found - adding to scheduled task" "INFO"
                break
            }
        }
        
        $actionScript = "cscript.exe //nologo $env:windir\system32\slmgr.vbs /ato $officeCmd"
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command `"$actionScript`""
        
        $trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Monday -At 9am
        
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable
        
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($existingTask) {
            Write-LogStep "Updating existing scheduled task..." "INFO"
            Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
            Write-LogStep "Task updated successfully" "SUCCESS"
        }
        else {
            Write-LogStep "Creating new scheduled task..." "INFO"
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description $taskDescription -User "SYSTEM" | Out-Null
            Write-LogStep "Task created successfully" "SUCCESS"
        }
        
        Show-StyledMessageBox -Message (Get-String "autoRenewalScheduled") -Title (Get-String "success") -Buttons "OK" -Icon "Information"
    }
    catch {
        Write-LogStep "Exception: $($_.Exception.Message)" "ERROR"
        
        Show-StyledMessageBox -Message "Failed to create scheduled task: $_" -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

Write-Host "ActivationCore Module loaded" -ForegroundColor Green
