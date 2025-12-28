# Edition Changer - Complete and Working

function Show-EditionChanger {
    Add-Type -AssemblyName PresentationFramework
    
    # Check if Core/Home FIRST
    $windowsInfo = Get-WindowsInfo
    $isCore = $windowsInfo.EditionID -match "Core|Home"
    
    if ($isCore) {
        Write-LogStep "Detected Core/Home edition: $($windowsInfo.EditionID)" "WARNING"
        
        $result = Show-StyledMessageBox -Message (Get-String "coreHomeDetected") -Title (Get-String "warning") -Buttons "YesNo" -Icon "Warning"
        
        if ($result -eq "Yes") {
            Upgrade-CoreToPro
        }
        return
    }
    
    # If not Core/Home, show edition selector
    $bgColor = if ($Global:CurrentTheme -eq "Dark") { "#1E1E1E" } else { "White" }
    $titleColor = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $textColor = if ($Global:CurrentTheme -eq "Dark") { "#B0B0B0" } else { "#666666" }
    $btnBg = if ($Global:CurrentTheme -eq "Dark") { "#3F3F3F" } else { "White" }
    $btnFg = if ($Global:CurrentTheme -eq "Dark") { "#E0E0E0" } else { "#333333" }
    $btnBorder = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#DDDDDD" }
    $hoverBg = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#F5F5F5" }
    
    $iconPath = Join-Path $Global:ScriptRoot "assets\key.ico"
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(Get-String 'changeWindowsEdition')" 
        SizeToContent="Height"
        MinHeight="500" MaxHeight="800"
        Width="650"
        WindowStartupLocation="CenterScreen"
        Background="$bgColor">
    
    <Window.Resources>
        <Style x:Key="EditionButton" TargetType="Button">
            <Setter Property="Background" Value="$btnBg"/>
            <Setter Property="Foreground" Value="$btnFg"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$btnBorder"/>
            <Setter Property="Padding" Value="15,12"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="$hoverBg"/>
                                <Setter Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="$btnBg"/>
            <Setter Property="Foreground" Value="$btnFg"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$btnBorder"/>
            <Setter Property="Padding" Value="20,10"/>
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
                                <Setter Property="Background" Value="$hoverBg"/>
                                <Setter Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="30">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Name="TitleText" FontSize="20" FontWeight="Bold" Foreground="$titleColor" Margin="0,0,0,10"/>
        <StackPanel Grid.Row="1" Margin="0,0,0,20">
            <TextBlock Name="CurrentEditionText" FontSize="13" Foreground="$textColor" Margin="0,0,0,5"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Name="EditionPanel">
                <Button Name="ProBtn" Style="{StaticResource EditionButton}">
                    <StackPanel Orientation="Horizontal">
                        <Border Background="#0078D4" CornerRadius="4" Width="32" Height="32" Margin="0,0,10,0">
                            <TextBlock Text="P" FontSize="16" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel>
                            <TextBlock Name="ProTitle" FontWeight="SemiBold" Foreground="$titleColor"/>
                            <TextBlock Name="ProDesc" FontSize="11" Foreground="$textColor"/>
                        </StackPanel>
                    </StackPanel>
                </Button>
                
                <Button Name="EducationBtn" Style="{StaticResource EditionButton}">
                    <StackPanel Orientation="Horizontal">
                        <Border Background="#107C10" CornerRadius="4" Width="32" Height="32" Margin="0,0,10,0">
                            <TextBlock Text="E" FontSize="16" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel>
                            <TextBlock Name="EducationTitle" FontWeight="SemiBold" Foreground="$titleColor"/>
                            <TextBlock Name="EducationDesc" FontSize="11" Foreground="$textColor"/>
                        </StackPanel>
                    </StackPanel>
                </Button>
                
                <Button Name="EnterpriseBtn" Style="{StaticResource EditionButton}">
                    <StackPanel Orientation="Horizontal">
                        <Border Background="#FFB900" CornerRadius="4" Width="32" Height="32" Margin="0,0,10,0">
                            <TextBlock Text="E" FontSize="16" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel>
                            <TextBlock Name="EnterpriseTitle" FontWeight="SemiBold" Foreground="$titleColor"/>
                            <TextBlock Name="EnterpriseDesc" FontSize="11" Foreground="$textColor"/>
                        </StackPanel>
                    </StackPanel>
                </Button>
            </StackPanel>
        
        <Button Grid.Row="3" Name="CancelBtn" Style="{StaticResource SecondaryButton}" Width="100" HorizontalAlignment="Right" Margin="0,20,0,0"/>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        if (Test-Path $iconPath) {
            $window.Icon = $iconPath
        }
        
        $titleText = $window.FindName("TitleText")
        $currentEditionText = $window.FindName("CurrentEditionText")
        $proTitle = $window.FindName("ProTitle")
        $proDesc = $window.FindName("ProDesc")
        $educationTitle = $window.FindName("EducationTitle")
        $educationDesc = $window.FindName("EducationDesc")
        $enterpriseTitle = $window.FindName("EnterpriseTitle")
        $enterpriseDesc = $window.FindName("EnterpriseDesc")
        $proBtn = $window.FindName("ProBtn")
        $educationBtn = $window.FindName("EducationBtn")
        $enterpriseBtn = $window.FindName("EnterpriseBtn")
        $cancelBtn = $window.FindName("CancelBtn")
        
        $titleText.Text = Get-String "changeWindowsEdition"
        $proTitle.Text = Get-String "windowsPro"
        $proDesc.Text = Get-String "proDescription"
        $educationTitle.Text = Get-String "windowsEducation"
        $educationDesc.Text = Get-String "educationDescription"
        $enterpriseTitle.Text = Get-String "windowsEnterprise"
        $enterpriseDesc.Text = Get-String "enterpriseDescription"
        $cancelBtn.Content = Get-String "cancel"
        
        $currentEditionText.Text = (Get-String "currentEdition") + ": $($windowsInfo.EditionID)"
        
        $proBtn.Add_Click({
            $window.Close()
            Change-ToEdition -Edition "Pro"
        })
        
        $educationBtn.Add_Click({
            $window.Close()
            Change-ToEdition -Edition "Education"
        })
        
        $enterpriseBtn.Add_Click({
            $window.Close()
            Change-ToEdition -Edition "Enterprise"
        })
        
        $cancelBtn.Add_Click({ $window.Close() })
        
        $window.ShowDialog() | Out-Null
    }
    catch {
        Write-Host "EditionChanger error: $_" -ForegroundColor Red
    }
}

function Upgrade-CoreToPro {
    Write-LogHeader (Get-String "upgradingToPro")
    Update-Status "Upgrading edition..."
    
    try {
        Write-LogStep (Get-String "usingDismUpgrade") "INFO"
        
        Write-LogStep "Configuring License Manager service..." "INFO"
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c sc config licensemanager start= auto & net start licensemanager" -PassThru -WindowStyle Hidden -Wait
        
        Write-LogStep "Configuring Windows Update service..." "INFO"
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c sc config wuauserv start= auto & net start wuauserv" -PassThru -WindowStyle Hidden -Wait
        
        Start-Sleep -Seconds 2
        
        $genericProKey = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"
        
        Write-LogStep "Applying Windows Pro generic key..." "INFO"
        Write-LogStep "Running: changepk /ProductKey $genericProKey" "INFO"
        
        $changePkPath = "$env:SystemRoot\System32\changepk.exe"
        
        if (Test-Path $changePkPath) {
            $process = Start-Process -FilePath $changePkPath -ArgumentList "/ProductKey $genericProKey" -PassThru -Wait
            
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                Write-LogStep (Get-String "upgradeSuccessful") "SUCCESS"
                Write-LogStep (Get-String "systemWillRestart") "WARNING"
                
                Show-StyledMessageBox -Message (Get-String "upgradeCompleteMessage") -Title (Get-String "success") -Buttons "OK" -Icon "Information"
                
                Start-Process -FilePath "shutdown.exe" -ArgumentList "/r /t 60 /c `"Windows Pro upgrade - restart required`"" -WindowStyle Hidden
            }
            else {
                throw "changepk.exe failed with exit code: $($process.ExitCode)"
            }
        }
        else {
            throw "changepk.exe not found at $changePkPath"
        }
    }
    catch {
        Write-LogStep ((Get-String "upgradeFailed") + ": $_") "ERROR"
        
        $errorMsg = (Get-String "upgradeFailedMessage") -f $_
        Show-StyledMessageBox -Message $errorMsg -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

function Change-ToEdition {
    param([string]$Edition)
    
    Write-LogHeader "Changing Windows Edition"
    Update-Status "Changing edition..."
    
    try {
        $upgradeKeys = @{
            Pro = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
            Education = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
            Enterprise = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
        }
        
        $key = $upgradeKeys[$Edition]
        
        if (-not $key) {
            throw (Get-String "invalidEdition")
        }
        
        Write-LogStep (Get-String "installingProductKey") "INFO"
        $slmgrPath = "$env:windir\system32\slmgr.vbs"
        
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/ipk", $key) -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep (Get-String "productKeyInstalled") "SUCCESS"
        }
        
        Start-Sleep -Seconds 2
        
        Write-LogStep (Get-String "settingKmsServer") "INFO"
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/skms", $Global:Config.KmsServer) -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep (Get-String "kmsServerConfigured") "SUCCESS"
        }
        
        Start-Sleep -Seconds 2
        
        $activatingMsg = (Get-String "activatingWindows") -f $Edition
        Write-LogStep $activatingMsg "INFO"
        
        $exitCode = Invoke-CommandWithRealTimeOutput -FilePath "cscript.exe" -ArgumentList @("//nologo", "`"$slmgrPath`"", "/ato") -WorkingDirectory $env:windir
        
        if ($exitCode -eq 0) {
            Write-LogStep (Get-String "activationSuccessful") "SUCCESS"
            
            $successMsg = (Get-String "editionChangeSuccess") -f $Edition
            Show-StyledMessageBox -Message $successMsg -Title (Get-String "success") -Buttons "OK" -Icon "Information"
        }
    }
    catch {
        Write-LogStep ((Get-String "editionChangeFailed") + ": $_") "ERROR"
        
        Show-StyledMessageBox -Message ((Get-String "editionChangeFailed") + ": $_") -Title (Get-String "error") -Buttons "OK" -Icon "Error"
    }
    finally {
        Update-Status "Ready"
    }
}

Write-Host "EditionChanger Module loaded" -ForegroundColor Green
