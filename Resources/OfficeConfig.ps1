# Office Configurations

$Global:OfficeVersions = @{
    '2016' = @{
        Name = "Office 2016 Professional Plus"
        ProductID = "ProPlus2016Volume"
        Channel = "PerpetualVL2016"
        Key = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    }
    '2019' = @{
        Name = "Office 2019 Professional Plus"
        ProductID = "ProPlus2019Volume"
        Channel = "PerpetualVL2019"
        Key = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
    }
    '2021' = @{
        Name = "Office 2021 LTSC Professional Plus"
        ProductID = "ProPlus2021Volume"
        Channel = "PerpetualVL2021"
        Key = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    }
    '2024' = @{
        Name = "Office 2024 LTSC Professional Plus"
        ProductID = "ProPlus2024Volume"
        Channel = "PerpetualVL2024"
        Key = "2TDPW-NDQ7G-FMG99-DXQ7M-TX3T2"
    }
    '365' = @{
        Name = "Microsoft 365 Apps"
        ProductID = "O365ProPlusRetail"
        Channel = "Current"
        Key = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    }
}

$Global:OfficeApps = @{
    Word = @{ ID = "Word"; ExcludeID = "" }
    Excel = @{ ID = "Excel"; ExcludeID = "" }
    PowerPoint = @{ ID = "PowerPoint"; ExcludeID = "" }
    Outlook = @{ ID = "Outlook"; ExcludeID = "" }
    Access = @{ ID = "Access"; ExcludeID = "Access" }
    Publisher = @{ ID = "Publisher"; ExcludeID = "Publisher" }
    OneNote = @{ ID = "OneNote"; ExcludeID = "OneNote" }
    Teams = @{ ID = "Teams"; ExcludeID = "Teams" }
    Groove = @{ ID = "Groove"; ExcludeID = "Groove" }
}

$Global:OfficePremium = @{
    Project = @{
        ProductID2016 = "ProjectPro2016Volume"
        ProductID2019 = "ProjectPro2019Volume"
        ProductID2021 = "ProjectPro2021Volume"
        Key2016 = "YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
        Key2019 = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
        Key2021 = "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8"
    }
    Visio = @{
        ProductID2016 = "VisioPro2016Volume"
        ProductID2019 = "VisioPro2019Volume"
        ProductID2021 = "VisioPro2021Volume"
        Key2016 = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
        Key2019 = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
        Key2021 = "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4"
    }
}

function Get-OfficeXML {
    param(
        [string]$Version,
        [hashtable]$SelectedApps,
        [bool]$IncludeProject = $false,
        [bool]$IncludeVisio = $false
    )
    
    $officeConfig = $Global:OfficeVersions[$Version]
    $excludeApps = @()
    
    # Build exclusion list
    foreach ($app in $Global:OfficeApps.Keys) {
        if (-not $SelectedApps[$app] -and $Global:OfficeApps[$app].ExcludeID) {
            $excludeApps += "      <ExcludeApp ID=`"$($Global:OfficeApps[$app].ExcludeID)`" />"
        }
    }
    
    $excludeString = if ($excludeApps.Count -gt 0) {
        "`n" + ($excludeApps -join "`n")
    } else { "" }
    
    # Configuration XML - Simple LTSC format (no PIDKEY in XML)
    $xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$($officeConfig.Channel)">
    <Product ID="$($officeConfig.ProductID)">      <Language ID="MatchOS" />$excludeString
    </Product>
"@

    # Add Project if selected
    if ($IncludeProject) {
        $projectConfig = $Global:OfficePremium.Project
        $projectID = "ProductID$Version"
        $xml += @"

    <Product ID="$($projectConfig[$projectID])">
      <Language ID="MatchOS" />
    </Product>
"@
    }
    
    # Add Visio if selected
    if ($IncludeVisio) {
        $visioConfig = $Global:OfficePremium.Visio
        $visioID = "ProductID$Version"
        $xml += @"

    <Product ID="$($visioConfig[$visioID])">
      <Language ID="MatchOS" />
    </Product>
"@
    }
    
    $xml += @"

  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%" />
</Configuration>
"@
    
    return $xml
}
