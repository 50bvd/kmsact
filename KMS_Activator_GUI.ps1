#Requires -RunAsAdministrator

# Force console encoding
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Get script root directory
$Global:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load modules in order
. "$Global:ScriptRoot\Resources\GlobalInit.ps1"
. "$Global:ScriptRoot\Resources\Languages.ps1"
. "$Global:ScriptRoot\Resources\Themes.ps1"
. "$Global:ScriptRoot\Resources\OfficeConfig.ps1"
. "$Global:ScriptRoot\Modules\UIHelper.ps1"
. "$Global:ScriptRoot\Modules\MessageBoxHelper.ps1"
. "$Global:ScriptRoot\Modules\ActivationCore.ps1"
. "$Global:ScriptRoot\Modules\EditionChanger.ps1"
. "$Global:ScriptRoot\Modules\OfficeInstaller.ps1"

# Initialize language
$Global:CurrentLanguage = Get-SystemLanguage
Write-Host "Detected system language: $Global:CurrentLanguage" -ForegroundColor Cyan

# Initialize theme
$Global:CurrentTheme = Get-WindowsTheme

# Load saved config
Load-Config

Write-Host "Final language: $Global:CurrentLanguage" -ForegroundColor Cyan
Write-Host "Final theme: $Global:CurrentTheme" -ForegroundColor Cyan

# Load XAML
function Load-XAML {
    param([string]$XamlFile)
    
    try {
        # In compiled EXE, use embedded XAML
        if ($Global:EmbeddedMainWindowXAML) {
            $xamlContent = $Global:EmbeddedMainWindowXAML
        } else {
            # Running as script, load from file
            $xamlPath = Join-Path $Global:ScriptRoot "UI\$XamlFile"
            if (-not (Test-Path $xamlPath)) {
                throw "XAML file not found: $xamlPath"
            }
            $xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8
        }
        
        $xamlContent = $xamlContent -replace 'MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"', ''
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlContent)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        return $window
    }
    catch {
        Write-Host "ERROR loading XAML: $_" -ForegroundColor Red
        throw $_
    }
}

# Settings Window
function Show-SettingsWindow {
    Write-Host "=== DEBUG: Show-SettingsWindow CALLED ===" -ForegroundColor Magenta
    
    # Adapt colors based on current theme
    $bgColor = if ($Global:CurrentTheme -eq "Dark") { "#1E1E1E" } else { "White" }
    $titleColor = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $labelColor = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $defaultLabelColor = if ($Global:CurrentTheme -eq "Dark") { "#B0B0B0" } else { "Gray" }
    
    # ComboBox/TextBox colors
    $inputBg = if ($Global:CurrentTheme -eq "Dark") { "#2D2D2D" } else { "#F5F5F5" }
    $inputFg = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $inputBorder = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#CCCCCC" }
    $inputHoverBg = if ($Global:CurrentTheme -eq "Dark") { "#3F3F3F" } else { "#E5E5E5" }
    
    # Button colors
    $secondaryBtnBg = if ($Global:CurrentTheme -eq "Dark") { "#2D2D2D" } else { "White" }
    $secondaryBtnFg = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $secondaryBtnBorder = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#CCCCCC" }
    $secondaryBtnHover = if ($Global:CurrentTheme -eq "Dark") { "#3F3F3F" } else { "#F5F5F5" }
    
    $iconPath = Join-Path $Global:ScriptRoot "assets\key.ico"
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Settings" Height="600" Width="600"
        WindowStartupLocation="CenterScreen"
        Background="$bgColor">
    
    <Window.Resources>
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="25,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#004275"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="$secondaryBtnBg"/>
            <Setter Property="Foreground" Value="$secondaryBtnFg"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$secondaryBtnBorder"/>
            <Setter Property="Padding" Value="25,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="$secondaryBtnHover"/>
                                <Setter Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="$inputHoverBg"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="StyledComboBox" TargetType="ComboBox">
            <Setter Property="Background" Value="$inputBg"/>
            <Setter Property="Foreground" Value="$inputFg"/>
            <Setter Property="BorderBrush" Value="$inputBorder"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Border Name="Border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <Grid>
                                <ToggleButton Name="ToggleButton" Grid.Column="2" 
                                            Focusable="false"
                                            IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                            ClickMode="Press">
                                    <ToggleButton.Template>
                                        <ControlTemplate TargetType="ToggleButton">
                                            <Grid Background="Transparent">
                                                <ContentPresenter />
                                            </Grid>
                                        </ControlTemplate>
                                    </ToggleButton.Template>
                                </ToggleButton>
                                <Border Background="{TemplateBinding Background}" Margin="10,0,30,0" IsHitTestVisible="False">
                                    <TextBlock Name="ContentSite" 
                                               Text="{TemplateBinding SelectionBoxItem}"
                                               VerticalAlignment="Center"
                                               HorizontalAlignment="Left"
                                               Foreground="{TemplateBinding Foreground}" />
                                </Border>
                                <TextBox x:Name="PART_EditableTextBox" Visibility="Hidden" IsReadOnly="{TemplateBinding IsReadOnly}"/>
                                <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                                     AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                    <Border Name="DropDownBorder" Background="$inputBg" BorderBrush="$inputBorder" 
                                            BorderThickness="1" CornerRadius="6" MaxHeight="200">
                                        <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                        </ScrollViewer>
                                    </Border>
                                </Popup>
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ComboBoxItem">
                        <Setter Property="Background" Value="$inputBg"/>
                        <Setter Property="Foreground" Value="$inputFg"/>
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ComboBoxItem">
                                    <Border Name="Border" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                        <ContentPresenter />
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="$inputHoverBg"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#0078D4"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="35">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Name="TitleText" Grid.Row="0" FontSize="22" FontWeight="Bold" Foreground="$titleColor" Margin="0,0,0,25"/>
        
        <StackPanel Grid.Row="1">
            <TextBlock Name="LangLabel" FontSize="15" FontWeight="Bold" Foreground="$labelColor" Margin="0,0,0,10"/>
            <ComboBox Name="LanguageCombo" Style="{StaticResource StyledComboBox}" Height="45" Margin="0,0,0,30" FontSize="14" Padding="12"/>
            
            <TextBlock Name="ThemeLabel" FontSize="15" FontWeight="Bold" Foreground="$labelColor" Margin="0,0,0,10"/>
            <ComboBox Name="ThemeCombo" Style="{StaticResource StyledComboBox}" Height="45" Margin="0,0,0,30" FontSize="14" Padding="12"/>
            
            <TextBlock Name="ServerLabel" FontSize="15" FontWeight="Bold" Foreground="$labelColor" Margin="0,0,0,10"/>
            <Border Background="$inputBg" BorderBrush="$inputBorder" BorderThickness="1" CornerRadius="6" Height="45" Margin="0,0,0,10">
                <TextBox Name="ServerText" Background="Transparent" Foreground="$inputFg" BorderThickness="0" 
                         FontSize="14" Padding="12" VerticalContentAlignment="Center"
                         CaretBrush="$inputFg" SelectionBrush="#0078D4" SelectionTextBrush="White"/>
            </Border>
            <TextBlock Name="DefaultLabel" FontSize="12" Foreground="$defaultLabelColor"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,30,0,0">
            <Button Name="CancelBtn" Style="{StaticResource SecondaryButton}" Margin="0,0,12,0"/>
            <Button Name="SaveBtn" Style="{StaticResource PrimaryButton}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        Write-Host "DEBUG: Parsing XAML..." -ForegroundColor Yellow
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $settingsWindow = [Windows.Markup.XamlReader]::Load($reader)
        
        Write-Host "DEBUG: Setting window icon..." -ForegroundColor Yellow
        if (Test-Path $iconPath) {
            $settingsWindow.Icon = $iconPath
        }
        
        Write-Host "DEBUG: Getting UI elements..." -ForegroundColor Yellow
        $langCombo = $settingsWindow.FindName("LanguageCombo")
        $themeCombo = $settingsWindow.FindName("ThemeCombo")
        $serverText = $settingsWindow.FindName("ServerText")
        $saveBtn = $settingsWindow.FindName("SaveBtn")
        $cancelBtn = $settingsWindow.FindName("CancelBtn")
        $titleText = $settingsWindow.FindName("TitleText")
        $langLabel = $settingsWindow.FindName("LangLabel")
        $themeLabel = $settingsWindow.FindName("ThemeLabel")
        $serverLabel = $settingsWindow.FindName("ServerLabel")
        $defaultLabel = $settingsWindow.FindName("DefaultLabel")
        
        Write-Host "DEBUG: Setting texts..." -ForegroundColor Yellow
        $titleText.Text = Get-String "settings"
        $langLabel.Text = Get-String "language"
        $themeLabel.Text = Get-String "theme"
        $serverLabel.Text = Get-String "kmsServer"
        $defaultLabel.Text = Get-String "kmsServerDefault"
        $saveBtn.Content = Get-String "save"
        $cancelBtn.Content = Get-String "cancel"
        
        Write-Host "DEBUG: Populating themes..." -ForegroundColor Yellow
        $lightItem = New-Object System.Windows.Controls.ComboBoxItem
        $lightItem.Content = Get-String "light"
        $lightItem.Tag = "Light"
        $themeCombo.Items.Add($lightItem) | Out-Null
        
        $darkItem = New-Object System.Windows.Controls.ComboBoxItem
        $darkItem.Content = Get-String "dark"
        $darkItem.Tag = "Dark"
        $themeCombo.Items.Add($darkItem) | Out-Null
        
        if ($Global:CurrentTheme -eq "Dark") {
            $themeCombo.SelectedItem = $darkItem
        } else {
            $themeCombo.SelectedItem = $lightItem
        }
        
        Write-Host "DEBUG: Populating languages..." -ForegroundColor Yellow
        $langCombo.Items.Clear()
        foreach ($langCode in $Global:Languages.Keys | Sort-Object) {
            $langName = $Global:Languages[$langCode].name
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $langName
            $item.Tag = $langCode
            $langCombo.Items.Add($item) | Out-Null
            
            if ($langCode -eq $Global:CurrentLanguage) {
                $langCombo.SelectedItem = $item
            }
        }
        
        $serverText.Text = $Global:Config.KmsServer
        
        Write-Host "DEBUG: Wiring buttons..." -ForegroundColor Yellow
        $saveBtn.Add_Click({
            $selectedLang = $langCombo.SelectedItem.Tag
            $selectedTheme = $themeCombo.SelectedItem.Tag
            $newKmsServer = $serverText.Text.Trim()
            $languageChanged = $false
            
            if ([string]::IsNullOrEmpty($newKmsServer)) {
                Show-StyledMessageBox -Message (Get-String "kmsServerEmpty") -Title (Get-String "error") -Buttons "OK" -Icon "Error"
                return
            }
            
            if ($selectedLang -ne $Global:CurrentLanguage) {
                $languageChanged = $true
            }
            
            $Global:CurrentLanguage = $selectedLang
            $Global:CurrentTheme = $selectedTheme
            $Global:Config.Language = $selectedLang
            $Global:Config.Theme = $selectedTheme
            $Global:Config.KmsServer = $newKmsServer
            
            Save-Config
            Apply-Theme -Theme $selectedTheme
            
            if ($languageChanged) {
                $Global:MainWindow.FindName("SystemInfoText").Text = Get-String "systemInfo"
                $Global:MainWindow.FindName("QuickActionsText").Text = Get-String "quickActions"
                $Global:MainWindow.FindName("ToolsText").Text = Get-String "tools"
                $Global:MainWindow.FindName("ConsoleOutputText").Text = Get-String "consoleOutput"
                $Global:MainWindow.FindName("ActivateWindowsBtn").Content = Get-String "activateWindows"
                $Global:MainWindow.FindName("ActivateOfficeBtn").Content = Get-String "activateOffice"
                $Global:MainWindow.FindName("ChangeEditionBtn").Content = Get-String "changeEdition"
                $Global:MainWindow.FindName("InstallOfficeBtn").Content = Get-String "installOffice"
                $Global:MainWindow.FindName("ScheduleTaskBtn").Content = Get-String "scheduleTasks"
                $Global:MainWindow.FindName("UninstallKeysBtn").Content = Get-String "uninstallKeys"
                $Global:MainWindow.FindName("ClearLogBtn").Content = Get-String "clearLog"
            }
            
            $settingsWindow.Close()
        })
        
        $cancelBtn.Add_Click({ $settingsWindow.Close() })
        
        Write-Host "DEBUG: Showing dialog..." -ForegroundColor Yellow
        $settingsWindow.ShowDialog() | Out-Null
        Write-Host "DEBUG: Dialog closed" -ForegroundColor Yellow
    }
    catch {
        Write-Host "SETTINGS ERROR: $_" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

# Create Main GUI
function Show-GUI {
    try {
        $window = Load-XAML "MainWindow.xaml"
        $Global:MainWindow = $window
        
        $iconPath = Join-Path $Global:ScriptRoot "icon.ico"
        if (Test-Path $iconPath) {
            try {
                $window.Icon = $iconPath
                Write-Host "Window icon set" -ForegroundColor Green
            } catch {
                Write-Host "Could not set icon" -ForegroundColor Yellow
            }
        }
        
        $Global:LogBox = $window.FindName("LogTextBox")
        $headerBar = $window.FindName("HeaderBar")
        $systemInfoText = $window.FindName("SystemInfoText")
        $quickActionsText = $window.FindName("QuickActionsText")
        $toolsText = $window.FindName("ToolsText")
        $consoleOutputText = $window.FindName("ConsoleOutputText")
        $versionLabel = $window.FindName("VersionLabel")
        $editionLabel = $window.FindName("EditionLabel")
        $buildLabel = $window.FindName("BuildLabel")
        $archLabel = $window.FindName("ArchLabel")
        $settingsBtn = $window.FindName("SettingsBtn")
        $activateWindowsBtn = $window.FindName("ActivateWindowsBtn")
        $activateOfficeBtn = $window.FindName("ActivateOfficeBtn")
        $changeEditionBtn = $window.FindName("ChangeEditionBtn")
        $installOfficeBtn = $window.FindName("InstallOfficeBtn")
        $scheduleTaskBtn = $window.FindName("ScheduleTaskBtn")
        $uninstallKeysBtn = $window.FindName("UninstallKeysBtn")
        $clearLogBtn = $window.FindName("ClearLogBtn")
        
        $windowsInfo = Get-WindowsInfo
        if ($null -eq $windowsInfo) {
            Show-StyledMessageBox -Message "Unsupported Windows version" -Title "Error" -Buttons "OK" -Icon "Error"
            exit
        }
        
        $versionLabel.Text = $windowsInfo.WindowsFamily
        $editionLabel.Text = $windowsInfo.EditionID
        $buildLabel.Text = "$($windowsInfo.CurrentBuild).$($windowsInfo.UBR)"
        $archLabel.Text = $windowsInfo.Architecture
        
        $systemInfoText.Text = Get-String "systemInfo"
        $quickActionsText.Text = Get-String "quickActions"
        $toolsText.Text = Get-String "tools"
        $consoleOutputText.Text = Get-String "consoleOutput"
        
        $activateWindowsBtn.Content = Get-String "activateWindows"
        $activateOfficeBtn.Content = Get-String "activateOffice"
        $changeEditionBtn.Content = Get-String "changeEdition"
        $installOfficeBtn.Content = Get-String "installOffice"
        $scheduleTaskBtn.Content = Get-String "scheduleTasks"
        $uninstallKeysBtn.Content = Get-String "uninstallKeys"
        $clearLogBtn.Content = Get-String "clearLog"
        
        Write-Log (Get-AsciiArt) "Green" -NoTimestamp
        Write-Log ""
        Write-LogHeader (Get-String "started")
        Write-LogStep "System: $($windowsInfo.WindowsFamily) $($windowsInfo.EditionID)" "INFO"
        Write-LogStep "Build: $($windowsInfo.CurrentBuild).$($windowsInfo.UBR)" "INFO"
        Write-LogStep "Architecture: $($windowsInfo.Architecture)" "INFO"
        Write-LogStep (Get-String "ready") "SUCCESS"
        Write-Log ""
        
        $window.Add_Loaded({
            Start-Sleep -Milliseconds 500
            
            if ($headerBar) {
                $headerColor = if ($Global:CurrentTheme -eq "Dark") { "#1E1E1E" } else { "#0078D4" }
                $headerBar.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($headerColor)
            }
            
            Apply-Theme -Theme $Global:CurrentTheme
        })
        
        # BUTTON EVENTS - WITH DEBUG
        Write-Host "=== Wiring button events ===" -ForegroundColor Cyan
        
        if ($settingsBtn) {
            Write-Host "DEBUG: Settings button found, wiring..." -ForegroundColor Green
            $settingsBtn.Add_Click({ 
                Write-Host "=== SETTINGS BUTTON CLICKED ===" -ForegroundColor Magenta
                Show-SettingsWindow 
            })
        } else {
            Write-Host "ERROR: Settings button NOT FOUND!" -ForegroundColor Red
        }
        
        $activateWindowsBtn.Add_Click({ Enable-Windows })
        $activateOfficeBtn.Add_Click({ Enable-Office })
        $changeEditionBtn.Add_Click({ Show-EditionChanger })
        $installOfficeBtn.Add_Click({ Show-OfficeInstaller })
        $scheduleTaskBtn.Add_Click({ New-ActivationSchedule })
        
        $uninstallKeysBtn.Add_Click({
            $result = Show-StyledMessageBox -Message (Get-String "uninstallKeysConfirm") -Title (Get-String "confirm") -Buttons "YesNo" -Icon "Warning"
            if ($result -eq "Yes") {
                Uninstall-ProductKey
            }
        })
        
        $clearLogBtn.Add_Click({
            $Global:LogBox.Clear()
            Write-Log (Get-AsciiArt) "Green" -NoTimestamp
            Write-Log ""
            Write-LogHeader (Get-String "logCleared")
            Write-LogStep (Get-String "ready") "SUCCESS"
            Write-Log ""
        })
        
        $window.ShowDialog() | Out-Null
    }
    catch {
        Write-Host "CRITICAL ERROR: $_" -ForegroundColor Red
        Show-StyledMessageBox -Message "Fatal error: $_" -Title "Error" -Buttons "OK" -Icon "Error"
    }
}

Show-GUI
