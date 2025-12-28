# Office Installer - FINAL VERSION with correct buttons and icon

function Show-OfficeInstaller {
    Add-Type -AssemblyName PresentationFramework
    
    # Check for existing Office
    $officeInstalled = Test-OfficeInstallation
    
    if ($officeInstalled) {
        $result = Show-StyledMessageBox -Message (Get-String "uninstallAndContinue") -Title (Get-String "existingOfficeDetected") -Buttons "YesNo" -Icon "Question"
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Uninstall-ExistingOffice
        }
        else {
            return
        }
    }
    
    # Adapt colors based on theme
    $bgColor = if ($Global:CurrentTheme -eq "Dark") { "#1E1E1E" } else { "White" }
    $titleColor = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $textColor = if ($Global:CurrentTheme -eq "Dark") { "#B0B0B0" } else { "#666666" }
    $checkFg = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $cardBg = if ($Global:CurrentTheme -eq "Dark") { "#2D2D2D" } else { "#F5F5F5" }
    $sectionTitle = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $btnBorder = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#CCCCCC" }
    
    # Icon paths
    $assetsPath = Join-Path $Global:ScriptRoot "assets"
    $windowIcon = Join-Path $assetsPath "key.ico"
    $officeIcon = Join-Path $assetsPath "office.png"
    $visioIcon = Join-Path $assetsPath "visio.png"
    $accesIcon = Join-Path $assetsPath "acces.png"
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(Get-String 'installOfficeLTSC')" 
        Height="600" 
        Width="650"
        WindowStartupLocation="CenterScreen"
        Background="$bgColor">
    
    <Window.Resources>
        <Style x:Key="CheckBoxStyle" TargetType="CheckBox">
            <Setter Property="Foreground" Value="$checkFg"/>
            <Setter Property="Margin" Value="0,8,0,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="25,12"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
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
            <Setter Property="Background" Value="$bgColor"/>
            <Setter Property="Foreground" Value="$checkFg"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$btnBorder"/>
            <Setter Property="Padding" Value="25,12"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="35">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="$(Get-String 'installOfficeLTSC')" FontSize="22" FontWeight="Bold" Foreground="$titleColor" Margin="0,0,0,10"/>
        <TextBlock Grid.Row="1" Text="$(Get-String 'selectComponents')" FontSize="13" Foreground="$textColor" Margin="0,0,0,25"/>
        
        <StackPanel Grid.Row="2">
            <!-- Office Suite -->
            <Border Background="$cardBg" CornerRadius="8" Padding="20" Margin="0,0,0,15">
                <StackPanel>
                    <TextBlock Text="$(Get-String 'officeSuite')" FontSize="15" FontWeight="Bold" Margin="0,0,0,15" Foreground="$sectionTitle"/>
                    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        
                        <Image Name="OfficeIcon" Grid.Column="0" Width="36" Height="36" Margin="0,0,15,0" VerticalAlignment="Center"/>
                        
                        <StackPanel Grid.Column="1">
                            <CheckBox Name="ProPlusCheck" Content="Office ProPlus 2024" Style="{StaticResource CheckBoxStyle}" IsChecked="True"/>
                            <TextBlock Text="Word, Excel, PowerPoint, Outlook, Access, Publisher" FontSize="11" Foreground="$textColor" Margin="25,0,0,0"/>
                        </StackPanel>
                    </Grid>
                </StackPanel>
            </Border>
            
            <!-- Additional Apps -->
            <Border Background="$cardBg" CornerRadius="8" Padding="20">
                <StackPanel>
                    <TextBlock Text="$(Get-String 'additionalApps')" FontSize="15" FontWeight="Bold" Margin="0,0,0,15" Foreground="$sectionTitle"/>
                    
                    <!-- Visio -->
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        
                        <Image Name="VisioIcon" Grid.Column="0" Width="36" Height="36" Margin="0,0,15,0" VerticalAlignment="Center"/>
                        <CheckBox Name="VisioCheck" Grid.Column="1" Content="Visio Professional 2024" Style="{StaticResource CheckBoxStyle}"/>
                    </Grid>
                    
                    <!-- Project (Access icon) -->
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        
                        <Image Name="ProjectIcon" Grid.Column="0" Width="36" Height="36" Margin="0,0,15,0" VerticalAlignment="Center"/>
                        <CheckBox Name="ProjectCheck" Grid.Column="1" Content="Project Professional 2024" Style="{StaticResource CheckBoxStyle}"/>
                    </Grid>
                </StackPanel>
            </Border>
        </StackPanel>
        
        <!-- Buttons - ORDER REVERSED TO MATCH SETTINGS -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,25,0,0">
            <Button Name="CancelBtn" Content="$(Get-String 'cancel')" Style="{StaticResource SecondaryButton}" Margin="0,0,12,0"/>
            <Button Name="InstallBtn" Content="$(Get-String 'install')" Style="{StaticResource PrimaryButton}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Set window icon
        if (Test-Path $windowIcon) {
            try {
                $window.Icon = $windowIcon
            } catch {
                # Fallback silently
            }
        }
        
        # Load icons
        $officeIconImg = $window.FindName("OfficeIcon")
        $visioIconImg = $window.FindName("VisioIcon")
        $projectIconImg = $window.FindName("ProjectIcon")
        
        if (Test-Path $officeIcon) {
            $officeIconImg.Source = $officeIcon
        }
        if (Test-Path $visioIcon) {
            $visioIconImg.Source = $visioIcon
        }
        if (Test-Path $accesIcon) {
            $projectIconImg.Source = $accesIcon
        }
        
        $proPlusCheck = $window.FindName("ProPlusCheck")
        $visioCheck = $window.FindName("VisioCheck")
        $projectCheck = $window.FindName("ProjectCheck")
        $installBtn = $window.FindName("InstallBtn")
        $cancelBtn = $window.FindName("CancelBtn")
        
        $installBtn.Add_Click({
            $components = @()
            
            if ($proPlusCheck.IsChecked) {
                $components += "ProPlus2024Volume"
            }
            if ($visioCheck.IsChecked) {
                $components += "VisioPro2024Volume"
            }
            if ($projectCheck.IsChecked) {
                $components += "ProjectPro2024Volume"
            }
            
            if ($components.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    (Get-String "selectOneComponent"),
                    (Get-String "noSelection"),
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            $window.Close()
            Install-Office2024 -Components $components
        })
        
        $cancelBtn.Add_Click({
            $window.Close()
        })
        
        $window.ShowDialog() | Out-Null
    }
    catch {
        Write-Host "OfficeInstaller error: $_" -ForegroundColor Red
    }
}

function Test-OfficeInstallation {
    $officePaths = @(
        "C:\Program Files\Microsoft Office",
        "C:\Program Files (x86)\Microsoft Office"
    )
    
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            return $true
        }
    }
    
    return $false
}

function Uninstall-ExistingOffice {
    try {
        Write-Log "Downloading Office Deployment Tool..." "Yellow"
        
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"
        $odtPath = "$env:TEMP\ODT.exe"
        $extractPath = "$env:TEMP\ODT"
        
        if (Test-Path $extractPath) {
            Remove-Item $extractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        Write-Log "Download complete" "Green"
        
        Write-Log "Extracting ODT..." "Yellow"
        Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$extractPath" -Wait
        
        Write-Log "Preparing uninstaller..." "Yellow"
        
        $uninstallXml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
        $uninstallXmlPath = "$extractPath\uninstall.xml"
        $uninstallXml | Out-File -FilePath $uninstallXmlPath -Encoding UTF8
        
        Write-Log "Launching Office uninstaller..." "Yellow"
        Write-Log "Please complete the uninstallation in the Office window" "Cyan"
        
        $setupPath = "$extractPath\setup.exe"
        $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$uninstallXmlPath`"" -PassThru -WindowStyle Hidden
        
        $lastUpdate = Get-Date
        $uninstallStartTime = Get-Date
        
        while (-not $process.HasExited) {
            [System.Windows.Forms.Application]::DoEvents()
            
            $now = Get-Date
            if (($now - $lastUpdate).TotalSeconds -ge 30) {
                $elapsed = [math]::Round(($now - $uninstallStartTime).TotalMinutes, 1)
                Write-Log "Uninstallation in progress... ($elapsed minutes elapsed)" "Cyan"
                $lastUpdate = $now
            }
            
            Start-Sleep -Milliseconds 500
        }
        
        Write-Log "Uninstallation completed" "Green"
        
        $stillInstalled = Test-OfficeInstallation
        if ($stillInstalled) {
            throw "Office may still be partially installed"
        }
        
        Write-Log "Office successfully uninstalled!" "Green"
    }
    catch {
        Write-Log "Failed to uninstall Office: $_" "Red"
        [System.Windows.MessageBox]::Show(
            (Get-String "failedToUninstall") + ": $_",
            (Get-String "error"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Install-Office2024 {
    param([string[]]$Components)
    
    try {
        Write-Log "Downloading Office Deployment Tool..." "Yellow"
        
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"
        $odtPath = "$env:TEMP\ODT.exe"
        $extractPath = "$env:TEMP\ODT"
        
        if (Test-Path $extractPath) {
            Remove-Item $extractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        Write-Log "Download complete" "Green"
        
        Write-Log "Extracting ODT..." "Yellow"
        Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$extractPath" -Wait
        
        Write-Log "Preparing installer..." "Yellow"
        
        $productXml = ""
        foreach ($component in $Components) {
            $productXml += "`n    <Product ID=`"$component`">`n      <Language ID=`"en-us`" />`n    </Product>"
        }
        
        $installXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">$productXml
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
        
        $installXmlPath = "$extractPath\install.xml"
        $installXml | Out-File -FilePath $installXmlPath -Encoding UTF8
        
        Write-Log "Launching Office installer..." "Yellow"
        Write-Log "Please complete the installation in the Office window (15-30 min)" "Cyan"
        
        $setupPath = "$extractPath\setup.exe"
        $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$installXmlPath`"" -PassThru -WindowStyle Hidden
        
        $lastUpdate = Get-Date
        $installStartTime = Get-Date
        
        while (-not $process.HasExited) {
            [System.Windows.Forms.Application]::DoEvents()
            
            $now = Get-Date
            if (($now - $lastUpdate).TotalSeconds -ge 60) {
                $elapsed = [math]::Round(($now - $installStartTime).TotalMinutes, 1)
                Write-Log "Installation in progress... ($elapsed minutes elapsed)" "Cyan"
                $lastUpdate = $now
            }
            
            Start-Sleep -Milliseconds 500
        }
        
        Write-Log "Installation completed" "Green"
        
        Write-Log "Verifying installation..." "Yellow"
        $installed = Test-OfficeInstallation
        
        if ($installed) {
            Write-Log "Office installed successfully!" "Green"
            
            $result = Show-StyledMessageBox -Message (Get-String "officeActivateNow") -Title (Get-String "success") -Buttons "YesNo" -Icon "Question"
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "Starting Office activation..." "Cyan"
                Enable-Office
            }
            else {
                Write-Log "Activation skipped - you can activate Office later from the main menu" "Yellow"
            }
        }
        else {
            throw "Installation verification failed"
        }
    }
    catch {
        Write-Log "Office installation failed: $_" "Red"
        Show-StyledMessageBox -Message "Failed to install Office: $_" -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
}

Write-Host "OfficeInstaller Module loaded" -ForegroundColor Green
