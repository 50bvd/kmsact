# Language Loader - Load translations from JSON files

$Global:Languages = @{}
$Global:CurrentLanguage = "en"

# Function to load language from JSON
function Load-Language {
    param([string]$LangCode)
    
    $jsonPath = Join-Path $PSScriptRoot "..\locales\$LangCode.json"
    
    if (Test-Path $jsonPath) {
        try {
            $langData = Get-Content $jsonPath -Raw | ConvertFrom-Json
            
            # Convert JSON object to hashtable
            $langHash = @{}
            $langData.PSObject.Properties | ForEach-Object {
                $langHash[$_.Name] = $_.Value
            }
            
            $Global:Languages[$LangCode] = $langHash
            return $true
        }
        catch {
            Write-Host "Error loading language $LangCode : $_" -ForegroundColor Red
            return $false
        }
    }
    return $false
}

# Load all available languages
$localesDir = Join-Path $PSScriptRoot "..\locales"
if (Test-Path $localesDir) {
    Get-ChildItem $localesDir -Filter "*.json" | ForEach-Object {
        $langCode = $_.BaseName
        Load-Language $langCode | Out-Null
    }
}

# Fallback to embedded languages if JSON files not found
if ($Global:Languages.Count -eq 0) {
    Write-Host "Warning: No JSON locales found, using embedded defaults" -ForegroundColor Yellow
    
    # Minimal English fallback
    $Global:Languages.en = @{
        name = "English"
        activateWindows = "Activate Windows"
        activateOffice = "Activate Office"
        changeEdition = "Change Edition"
        installOffice = "Install Office"
        scheduleTasks = "Schedule Auto-Renewal"
        uninstallKeys = "Uninstall Keys"
        clearLog = "Clear Console"
        settings = "Settings"
        ready = "Ready"
        yes = "Yes"
        no = "No"
        ok = "OK"
        cancel = "Cancel"
    }
}

# Auto-detect system language
function Get-SystemLanguage {
    $culture = [System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName
    
    # Map to available languages
    if ($Global:Languages.ContainsKey($culture)) {
        return $culture
    }
    
    # Default to English
    return "en"
}

# Get localized string
function Get-String {
    param([string]$Key)
    
    if ($Global:Languages.ContainsKey($Global:CurrentLanguage)) {
        $lang = $Global:Languages[$Global:CurrentLanguage]
        if ($lang.ContainsKey($Key)) {
            return $lang[$Key]
        }
    }
    
    # Fallback to English
    if ($Global:Languages.ContainsKey("en") -and $Global:Languages.en.ContainsKey($Key)) {
        return $Global:Languages.en[$Key]
    }
    
    # Return key if not found
    return $Key
}

# Initialize with system language
$Global:CurrentLanguage = Get-SystemLanguage

Write-Host "Loaded $($Global:Languages.Count) languages" -ForegroundColor Green
Write-Host "Current language: $Global:CurrentLanguage" -ForegroundColor Cyan
