# Theme definitions for KMS Activator

$Global:Themes = @{
    Dark = @{
        Name = "Dark"
        Window = "#1E1E1E"
        Panel = "#252526"
        PanelLight = "#2D2D30"
        Header = "#0078D4"
        Accent = "#0078D4"
        AccentHover = "#3E3E42"
        AccentPressed = "#007ACC"
        Border = "#3E3E3E"
        Text = "#FFFFFF"
        TextSecondary = "#B0B0B0"
        TextTertiary = "#808080"
        Console = "#0C0C0C"
        ConsoleText = "#00FF00"
        ProgressBar = "#0078D4"
        Success = "#4EC9B0"
        Error = "#F48771"
        Warning = "#CE9178"
    }
    
    Light = @{
        Name = "Light"
        Window = "#FFFFFF"
        Panel = "#F3F3F3"
        PanelLight = "#E8E8E8"
        Header = "#0078D4"
        Accent = "#0078D4"
        AccentHover = "#005A9E"
        AccentPressed = "#004578"
        Border = "#D0D0D0"
        Text = "#000000"
        TextSecondary = "#4A4A4A"
        TextTertiary = "#707070"
        Console = "#FFFFFF"
        ConsoleText = "#008000"
        ProgressBar = "#0078D4"
        Success = "#107C10"
        Error = "#E81123"
        Warning = "#FF8C00"
    }
}

# Current theme
$Global:CurrentTheme = "Dark"

function Get-WindowsTheme {
    try {
        # Check Windows theme from registry
        $themePath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $appsUseLightTheme = Get-ItemProperty -Path $themePath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        
        if ($null -ne $appsUseLightTheme) {
            # 0 = Dark theme, 1 = Light theme
            return if ($appsUseLightTheme.AppsUseLightTheme -eq 0) { "Dark" } else { "Light" }
        }
        
        # Fallback to Light if can't detect
        return "Light"
    }
    catch {
        return "Light"
    }
}

function Get-ThemeColor {
    param([string]$ColorKey)
    return $Global:Themes[$Global:CurrentTheme][$ColorKey]
}

function Set-Theme {
    param([string]$ThemeName)
    
    if ($Global:Themes.ContainsKey($ThemeName)) {
        $Global:CurrentTheme = $ThemeName
    }
}
