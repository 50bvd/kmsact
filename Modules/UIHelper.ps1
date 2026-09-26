# UI Helper Functions

# Get Windows Information
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
            
            # Determine Windows family
            if ([System.Environment]::OSVersion.Version.Build -ge 22000) {
                $windowsInfo.WindowsFamily = "Windows 11"
            }
            else {
                $windowsInfo.WindowsFamily = "Windows 10"
            }
            
            # Populate available editions from global config
            if ($Global:Config -and $Global:Config.WindowsEditions) {
                $windowsInfo.AvailableEditions = $Global:Config.WindowsEditions.Values | Where-Object { $_.Name -ne "Home" }
            } else {
                $windowsInfo.AvailableEditions = @()
            }
            
            # Find current edition and KMS key
            if ($Global:Config -and $Global:Config.WindowsEditions) {
                foreach ($edition in $Global:Config.WindowsEditions.Values) {
                    if ($edition.EditionId -eq $windowsInfo.EditionID) {
                        $windowsInfo.CurrentEdition = $edition
                        $windowsInfo.KMSKey = $edition.KMSKey
                        break
                    }
                }
            }
        }
        
        return $windowsInfo
    }
    catch {
        Write-Host "Error getting Windows info: $_" -ForegroundColor Red
        return $null
    }
}

# Write colored log
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewLine,
        [switch]$NoTimestamp
    )
    
    if ($NoTimestamp) {
        $formattedMessage = $Message
    } else {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $formattedMessage = "[$timestamp] $Message"
    }
    
    if ($Global:LogBox) {
        # Marshal to the UI thread SYNCHRONOUSLY. Write-Log runs on the UI thread,
        # so Invoke executes the append inline and the line shows immediately with
        # no blocking. Do NOT switch this to an async BeginInvoke closure: the
        # deferred block would run after Write-Log has returned, when
        # $formattedMessage is no longer in scope, and would append nothing -
        # leaving the console blank.
        $utf8Message = if ($NoNewLine) { $formattedMessage } else { "$formattedMessage`n" }
        try {
            $Global:LogBox.Dispatcher.Invoke([action]{
                $Global:LogBox.AppendText($utf8Message)
                $Global:LogBox.ScrollToEnd()
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
        } catch {}
    }
    
    # Also write to console
    $colorMap = @{
        "Red" = [System.ConsoleColor]::Red
        "Green" = [System.ConsoleColor]::Green
        "Yellow" = [System.ConsoleColor]::Yellow
        "Cyan" = [System.ConsoleColor]::Cyan
        "Magenta" = [System.ConsoleColor]::Magenta
        "Gray" = [System.ConsoleColor]::Gray
        "White" = [System.ConsoleColor]::White
    }
    
    $consoleColor = if ($colorMap.ContainsKey($Color)) { $colorMap[$Color] } else { [System.ConsoleColor]::White }
    if ($NoNewLine) {
        Write-Host $formattedMessage -ForegroundColor $consoleColor -NoNewline
    } else {
        Write-Host $formattedMessage -ForegroundColor $consoleColor
    }
}

# Write section header
function Write-LogHeader {
    param([string]$Title)
    
    Write-Log "" "White"
    Write-Log "===========================================" "Cyan"
    Write-Log "  $Title" "Cyan"
    Write-Log "===========================================" "Cyan"
    Write-Log "" "White"
}

# Write step
function Write-LogStep {
    param(
        [string]$Step,
        [string]$Status = "INFO"
    )
    
    $color = switch ($Status) {
        "SUCCESS" { "Green" }
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "INFO" { "Cyan" }
        default { "White" }
    }
    
    $prefix = switch ($Status) {
        "SUCCESS" { "[OK]" }
        "ERROR" { "[ERROR]" }
        "WARNING" { "[WARN]" }
        "INFO" { "[INFO]" }
        default { "[*]" }
    }
    
    Write-Log "$prefix $Step" $color
}

# Update status
function Update-Status {
    param([string]$Status)
    
    if ($Global:StatusLabel) {
        $Global:StatusLabel.Text = $Status
    }
}

# ASCII Art
function Get-AsciiArt {
    return @"
  ___                       ___ 
 (o o)                     (o o)
(  V  ) MS KMS Activation (  V  )
--m-m-----------------------m-m--
    https://github.com/50bvd
"@
}

# Global config is now initialized in GlobalInit.ps1
# Removed from here to avoid duplication in monolithic build

# Initialize remaining globals if not already set
if (-not $Global:CurrentTheme) { $Global:CurrentTheme = "Dark" }
if (-not $Global:LogBox) { $Global:LogBox = $null }
if (-not $Global:StatusLabel) { $Global:StatusLabel = $null }
if (-not $Global:MainWindow) { $Global:MainWindow = $null }

# Restart application
function Restart-Application {
    param([string]$Message = "Restart required to apply changes. Restart now?")
    
    $result = Show-CustomMessageBox -Message $Message -Title "Restart Required" -Type "YesNo"
    
    if ($result -eq "Yes") {
        # Close current window
        if ($Global:MainWindow) {
            $Global:MainWindow.Dispatcher.Invoke([action]{
                $Global:MainWindow.Close()
            }, "Normal")
        }
        
        # Check if running from compiled EXE or script
        $parentProcess = (Get-WmiObject Win32_Process -Filter "ProcessId=$PID").ParentProcessId
        $parentName = (Get-Process -Id $parentProcess -ErrorAction SilentlyContinue).Name
        
        if ($parentName -eq "KMS_Activator" -or $parentName -match "KMS_Activator") {
            # Running from compiled EXE - restart the EXE
            $exePath = (Get-Process -Id $parentProcess).Path
            Start-Process -FilePath $exePath
        } else {
            # Running as script - restart PowerShell script
            $scriptPath = $Global:ScriptRoot
            $mainScript = Join-Path $scriptPath "KMS_Activator_GUI.ps1"
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File `"$mainScript`""
        }
        
        exit
    }
}

# Apply theme to any window
function Apply-ThemeToWindow {
    param(
        [System.Windows.Window]$Window,
        [string]$Theme = $Global:CurrentTheme
    )
    
    if (-not $Window) { return }
    
    try {
        if ($Theme -eq "Dark") {
            # Dark theme
            $Window.Background = "#1E1E1E"
            
            # Helper function to get all children
            function Get-VisualChildren {
                param($parent)
                $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent)
                for ($i = 0; $i -lt $count; $i++) {
                    $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                    $child
                    Get-VisualChildren $child
                }
            }
            
            # Update all TextBlocks
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object {
                if ($_.FontSize -ge 14 -or $_.FontWeight -eq "Bold") {
                    $_.Foreground = "#FFFFFF"
                } else {
                    $_.Foreground = "#B0B0B0"
                }
            }
            
            # Update all Borders/Cards
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.Border] } | ForEach-Object {
                if ($_.Background -and $_.Background.ToString() -eq "#FFFFFFFF") {
                    $_.Background = "#2D2D2D"
                }
                if ($_.BorderBrush) {
                    $_.BorderBrush = "#3F3F3F"
                }
            }
            
            # Update all Buttons
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.Button] } | ForEach-Object {
                # Skip blue primary buttons
                if ($_.Background.ToString() -ne "#FF0078D4") {
                    $_.Background = "#3F3F3F"
                    $_.BorderBrush = "#4F4F4F"
                    
                    # Add hover effect using MouseEnter/MouseLeave - ONLY border color
                    $_.Add_MouseEnter({
                        if ($this.Background.ToString() -ne "#FF0078D4") {
                            $this.BorderBrush = "#0078D4"  # Blue border only
                        }
                    })
                    $_.Add_MouseLeave({
                        if ($this.Background.ToString() -ne "#FF0078D4") {
                            $this.BorderBrush = "#4F4F4F"  # Restore gray border
                        }
                    })
                }
                
                # Update button text
                Get-VisualChildren $_ | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object {
                    if ($_.Parent.Background.ToString() -eq "#FF0078D4") {
                        $_.Foreground = "#FFFFFF"  # White text on blue button
                    } else {
                        $_.Foreground = "#E0E0E0"  # Light gray on dark button
                    }
                }
            }
            
            # Update ComboBoxes and their popups
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.ComboBox] } | ForEach-Object {
                $comboBox = $_
                $comboBox.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                $comboBox.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0"))
                $comboBox.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#4F4F4F"))
                
                # Force template parts
                $comboBox.ApplyTemplate()
                $toggleButton = $comboBox.Template.FindName("toggleButton", $comboBox)
                if ($toggleButton) {
                    $toggleButton.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                }
                
                # Handle dropdown opened event to style popup
                $comboBox.Add_DropDownOpened({
                    Start-Sleep -Milliseconds 50
                    
                    # Find and style the popup
                    $popup = $comboBox.Template.FindName("PART_Popup", $comboBox)
                    if ($popup -and $popup.Child) {
                        # Style popup background
                        Get-VisualChildren $popup.Child | ForEach-Object {
                            if ($_ -is [System.Windows.Controls.Border]) {
                                $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                                $_.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#4F4F4F"))
                            }
                            # Style items in dropdown
                            if ($_ -is [System.Windows.Controls.ScrollViewer]) {
                                $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                            }
                        }
                        
                        # Style dropdown items
                        for ($i = 0; $i -lt $comboBox.Items.Count; $i++) {
                            $container = $comboBox.ItemContainerGenerator.ContainerFromIndex($i)
                            if ($container) {
                                $container.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                                $container.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0"))
                                $container.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#3F3F3F"))
                                
                                # Hover effect
                                $container.Add_MouseEnter({ $this.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#3F3F3F")) })
                                $container.Add_MouseLeave({ $this.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D")) })
                            }
                        }
                    }
                })
            }
            
            # Update TextBoxes
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.TextBox] } | ForEach-Object {
                $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#2D2D2D"))
                $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0"))
                $_.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#4F4F4F"))
                $_.CaretBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0"))
            }
            
            # Update CheckBoxes
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.CheckBox] } | ForEach-Object {
                $_.Foreground = "#E0E0E0"
            }
            
            # Update Separators
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.Separator] } | ForEach-Object {
                $_.Background = "#3F3F3F"
            }
        }
        else {
            # Light theme - restore defaults
            $Window.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
            
            # Helper function to get all children
            function Get-VisualChildren {
                param($parent)
                $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent)
                for ($i = 0; $i -lt $count; $i++) {
                    $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
                    $child
                    Get-VisualChildren $child
                }
            }
            
            # Update all TextBlocks to dark - FORCE with SolidColorBrush
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object {
                $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
                $_.SetValue([System.Windows.Controls.TextBlock]::ForegroundProperty, [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333")))
            }
            
            # Update all Borders/Cards to white
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.Border] } | ForEach-Object {
                if ($_.Background -and $_.Background.ToString() -ne "Transparent") {
                    $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
                }
                if ($_.BorderBrush) {
                    $_.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0"))
                }
            }
            
            # Update all Buttons (except blue primary)
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.Button] } | ForEach-Object {
                # Skip blue primary buttons
                if ($_.Background.ToString() -ne "#FF0078D4") {
                    $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
                    $_.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#DDDDDD"))
                    $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
                    
                    # Update button text - FORCE black
                    Get-VisualChildren $_ | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object {
                        $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
                        $_.SetValue([System.Windows.Controls.TextBlock]::ForegroundProperty, [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333")))
                    }
                } else {
                    # Blue buttons keep white text
                    Get-VisualChildren $_ | Where-Object { $_ -is [System.Windows.Controls.TextBlock] } | ForEach-Object {
                        $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
                    }
                }
            }
            
            # Update TextBoxes
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.TextBox] } | ForEach-Object {
                $_.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
                $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
                $_.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#CCCCCC"))
                $_.CaretBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
            }
            
            # Update CheckBoxes
            Get-VisualChildren $Window | Where-Object { $_ -is [System.Windows.Controls.CheckBox] } | ForEach-Object {
                $_.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#333333"))
            }
        }
    }
    catch {
        Write-Host "Error applying theme to window: $_" -ForegroundColor Yellow
    }
}

# Detect Windows theme (Dark/Light)
function Get-WindowsTheme {
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $appsUseLightTheme = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        
        if ($null -eq $appsUseLightTheme -or $appsUseLightTheme.AppsUseLightTheme -eq 0) {
            Write-Host "Detected Windows theme: Dark" -ForegroundColor Cyan
            return "Dark"
        }
        else {
            Write-Host "Detected Windows theme: Light" -ForegroundColor Cyan
            return "Light"
        }
    }
    catch {
        Write-Host "Could not detect theme, defaulting to Dark" -ForegroundColor Yellow
        return "Dark"
    }
}

# Modern light/dark palettes for the main window. Because every themed element
# in MainWindow.xaml references its color via {DynamicResource ...}, swapping
# these resource values re-themes the whole window instantly - no visual-tree
# walking and no per-element hardcoded colors.
function Set-AppThemeResources {
    param(
        [System.Windows.Window]$Window,
        [string]$Theme = $Global:CurrentTheme
    )

    if (-not $Window) { return }

    if ($Theme -eq "Dark") {
        $palette = @{
            AppWindowBrush        = "#1B1B1F"
            AppCardBrush          = "#26262B"
            AppCardBorderBrush    = "#37373D"
            AppButtonBrush        = "#2F2F35"
            AppButtonBorderBrush  = "#3C3C43"
            AppTextPrimaryBrush   = "#F3F3F3"
            AppTextSecondaryBrush = "#A9A9B2"
            AppTextButtonBrush    = "#ECECEC"
            AppAccentBrush        = "#4CA9FF"
            AppSeparatorBrush     = "#37373D"
        }
    }
    else {
        $palette = @{
            AppWindowBrush        = "#F5F6F8"
            AppCardBrush          = "#FFFFFF"
            AppCardBorderBrush    = "#ECEFF3"
            AppButtonBrush        = "#FFFFFF"
            AppButtonBorderBrush  = "#E3E6EA"
            AppTextPrimaryBrush   = "#1F2328"
            AppTextSecondaryBrush = "#5B6068"
            AppTextButtonBrush    = "#2B2F36"
            AppAccentBrush        = "#0078D4"
            AppSeparatorBrush     = "#E4E7EC"
        }
    }

    foreach ($key in $palette.Keys) {
        try {
            $color = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$key])
            $Window.Resources[$key] = [System.Windows.Media.SolidColorBrush]::new($color)
        }
        catch {
            Write-Host "Theme resource '$key' failed: $_" -ForegroundColor Yellow
        }
    }
}

# Apply theme to the main window by swapping DynamicResource brushes.
function Apply-Theme {
    param([string]$Theme = $Global:CurrentTheme)

    if (-not $Global:MainWindow) { return }

    try {
        $Global:MainWindow.Dispatcher.Invoke([action]{
            Set-AppThemeResources -Window $Global:MainWindow -Theme $Theme
        }, "Normal")
    }
    catch {
        Write-Host "Error applying theme: $_" -ForegroundColor Red
    }
}

# Save config to file
function Save-Config {
    try {
        # Save in AppData for persistence (not in temp folder)
        $configDir = Join-Path $env:APPDATA "KMSActivator"
        if (-not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }
        
        $configPath = Join-Path $configDir "config.json"
        $configData = @{
            KMSServer = $Global:Config.KMSServer
            Language = $Global:CurrentLanguage
            Theme = $Global:CurrentTheme
        }
        $configData | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8
        Write-LogStep "Configuration saved to AppData" "SUCCESS"
        return $true
    }
    catch {
        Write-LogStep "Failed to save config: $_" "ERROR"
        return $false
    }
}

# Load config from file
function Load-Config {
    try {
        # Load from AppData (persistent location)
        $configPath = Join-Path $env:APPDATA "KMSActivator\config.json"
        if (Test-Path $configPath) {
            $loaded = Get-Content $configPath -Raw | ConvertFrom-Json
            if ($loaded.KMSServer) {
                $Global:Config.KMSServer = $loaded.KMSServer
            }
            if ($loaded.Language) {
                # Only use saved language if it exists in available languages
                if ($Global:Languages.ContainsKey($loaded.Language)) {
                    $Global:CurrentLanguage = $loaded.Language
                    Write-Host "Language loaded from config: $($Global:CurrentLanguage)" -ForegroundColor Green
                } else {
                    Write-Host "Saved language not available, using detected language" -ForegroundColor Yellow
                }
            }
            if ($loaded.Theme) {
                $Global:CurrentTheme = $loaded.Theme
                Write-Host "Theme loaded from config: $($Global:CurrentTheme)" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "Failed to load config: $_" -ForegroundColor Yellow
    }
}
