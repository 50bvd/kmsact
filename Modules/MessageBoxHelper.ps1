# MessageBox Helper - Simple WPF MessageBox (NO EMOJIS)

function Show-StyledMessageBox {
    param(
        [string]$Message,
        [string]$Title,
        [string]$Buttons = "OK",
        [string]$Icon = "Information"
    )
    
    Add-Type -AssemblyName PresentationFramework
    
    $bgColor = if ($Global:CurrentTheme -eq "Dark") { "#1E1E1E" } else { "White" }
    $titleColor = if ($Global:CurrentTheme -eq "Dark") { "#FFFFFF" } else { "#333333" }
    $textColor = if ($Global:CurrentTheme -eq "Dark") { "#E0E0E0" } else { "#333333" }
    $btnBg = if ($Global:CurrentTheme -eq "Dark") { "#2D2D2D" } else { "White" }
    $btnFg = if ($Global:CurrentTheme -eq "Dark") { "#E0E0E0" } else { "#333333" }
    $btnBorder = if ($Global:CurrentTheme -eq "Dark") { "#4F4F4F" } else { "#CCCCCC" }
    $btnHover = if ($Global:CurrentTheme -eq "Dark") { "#3F3F3F" } else { "#F5F5F5" }
    
    $iconPath = Join-Path $Global:ScriptRoot "assets\key.ico"
    
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" 
        SizeToContent="WidthAndHeight"
        MinWidth="450" MaxWidth="600"
        MinHeight="200" MaxHeight="500"
        WindowStartupLocation="CenterScreen"
        Background="$bgColor"
        ResizeMode="NoResize">
    
    <Window.Resources>
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="25,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="MinWidth" Value="100"/>
            <Setter Property="Margin" Value="5,0"/>
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
            <Setter Property="Background" Value="$btnBg"/>
            <Setter Property="Foreground" Value="$btnFg"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$btnBorder"/>
            <Setter Property="Padding" Value="25,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="MinWidth" Value="100"/>
            <Setter Property="Margin" Value="5,0"/>
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
                                <Setter Property="Background" Value="$btnHover"/>
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
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Name="MessageText" Text="$Message" FontSize="14" Foreground="$textColor" TextWrapping="Wrap" MaxWidth="500" Margin="0,0,0,20"/>
        
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="YesBtn" Style="{StaticResource PrimaryButton}" Visibility="Collapsed"/>
            <Button Name="NoBtn" Style="{StaticResource SecondaryButton}" Visibility="Collapsed"/>
            <Button Name="OkBtn" Style="{StaticResource PrimaryButton}" Visibility="Collapsed"/>
            <Button Name="CancelBtn" Style="{StaticResource SecondaryButton}" Visibility="Collapsed"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        if (Test-Path $iconPath) {
            $window.Icon = $iconPath
        }
        
        $yesBtn = $window.FindName("YesBtn")
        $noBtn = $window.FindName("NoBtn")
        $okBtn = $window.FindName("OkBtn")
        $cancelBtn = $window.FindName("CancelBtn")
        
        $window.Tag = $null
        
        if ($Buttons -eq "YesNo") {
            $yesBtn.Content = Get-String "yes"
            $noBtn.Content = Get-String "no"
            $yesBtn.Visibility = "Visible"
            $noBtn.Visibility = "Visible"
            
            $yesBtn.Add_Click({
                $window.Tag = "Yes"
                $window.Close()
            })
            
            $noBtn.Add_Click({
                $window.Tag = "No"
                $window.Close()
            })
        }
        elseif ($Buttons -eq "OKCancel") {
            $okBtn.Content = Get-String "ok"
            $cancelBtn.Content = Get-String "cancel"
            $okBtn.Visibility = "Visible"
            $cancelBtn.Visibility = "Visible"
            
            $okBtn.Add_Click({
                $window.Tag = "OK"
                $window.Close()
            })
            
            $cancelBtn.Add_Click({
                $window.Tag = "Cancel"
                $window.Close()
            })
        }
        else {
            $okBtn.Content = Get-String "ok"
            $okBtn.Visibility = "Visible"
            
            $okBtn.Add_Click({
                $window.Tag = "OK"
                $window.Close()
            })
        }
        
        $window.ShowDialog() | Out-Null
        return $window.Tag
    }
    catch {
        Write-Host "MessageBox error: $_" -ForegroundColor Red
        return $null
    }
}

Write-Host "MessageBoxHelper Module loaded" -ForegroundColor Green
