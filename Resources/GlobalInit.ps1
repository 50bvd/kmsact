# Global Initialization - Load FIRST in monolithic build

# Initialize Global Config BEFORE any other module
$Global:Config = @{
    KMSServer = "kms.50bvd.com"
    WindowsEditions = @{
        Pro = @{
            Name = "Pro"
            EditionId = "Professional"
            GenericKey = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"
            KMSKey = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
        }
        Education = @{
            Name = "Education"
            EditionId = "Education"
            GenericKey = "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
            KMSKey = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
        }
        Enterprise = @{
            Name = "Enterprise"
            EditionId = "Enterprise"
            GenericKey = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
            KMSKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
        }
        Home = @{
            Name = "Home"
            EditionId = "Core"
            KMSKey = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
        }
    }
    Office = @{
        Key = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
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

# Initialize other globals
$Global:CurrentTheme = "Dark"
$Global:CurrentLanguage = "English"
$Global:LogBox = $null
$Global:StatusLabel = $null
$Global:MainWindow = $null

Write-Host "[INIT] Global variables initialized" -ForegroundColor Green
