# Update checker - queries the GitHub Releases API and compares the newest
# published release against the running version ($Global:AppVersion).

function ConvertTo-AppVersion {
    param([string]$Tag)

    if ([string]::IsNullOrWhiteSpace($Tag)) { return $null }
    $t = $Tag.Trim().TrimStart('v', 'V')
    $parts = $t -split '-', 2
    $base = $parts[0]
    $pre = if ($parts.Count -gt 1) { $parts[1] } else { '' }
    try { $ver = [version]$base } catch { $ver = $null }
    return [pscustomobject]@{ Version = $ver; Pre = $pre; Raw = $Tag }
}

# Returns > 0 when $A is newer than $B, < 0 when older, 0 when equal.
# A final release outranks a pre-release that shares the same base version.
function Compare-AppVersion {
    param($A, $B)

    if ($null -eq $A -or $null -eq $A.Version) { return -1 }
    if ($null -eq $B -or $null -eq $B.Version) { return 1 }

    $c = $A.Version.CompareTo($B.Version)
    if ($c -ne 0) { return $c }

    if ($A.Pre -eq '' -and $B.Pre -ne '') { return 1 }
    if ($A.Pre -ne '' -and $B.Pre -eq '') { return -1 }
    return [string]::Compare($A.Pre, $B.Pre, $true)
}

# Checks GitHub for a newer release. -Silent suppresses the "up to date" and
# error dialogs (used for an unobtrusive check at startup).
function Test-ForUpdate {
    param([switch]$Silent)

    Write-LogHeader "Update Check"
    Update-Status "Checking for updates..."

    try {
        $current = ConvertTo-AppVersion $Global:AppVersion
        Write-LogStep "Current version: v$($Global:AppVersion)" "INFO"

        $url = "https://api.github.com/repos/$($Global:AppRepo)/releases?per_page=20"
        Write-LogStep "Querying: $url" "INFO"

        $headers = @{
            'User-Agent' = 'KMS-Activator-UpdateCheck'
            'Accept'     = 'application/vnd.github+json'
        }

        # TLS 1.2 for older Windows PowerShell hosts.
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

        $releases = @(Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 15 -ErrorAction Stop |
            Where-Object { -not $_.draft })

        if ($releases.Count -eq 0) {
            Write-LogStep "No releases published yet" "WARNING"
            if (-not $Silent) {
                Show-StyledMessageBox -Message (Get-String "noUpdate") -Title (Get-String "info") -Buttons "OK" -Icon "Information"
            }
            return
        }

        $best = $null
        foreach ($r in $releases) {
            $rv = ConvertTo-AppVersion $r.tag_name
            if ($null -eq $best -or (Compare-AppVersion $rv $best.Ver) -gt 0) {
                $best = [pscustomobject]@{ Ver = $rv; Rel = $r }
            }
        }

        if ((Compare-AppVersion $best.Ver $current) -gt 0) {
            $tag = $best.Rel.tag_name
            $link = $best.Rel.html_url
            Write-LogStep "Update available: $tag" "SUCCESS"
            $msg = ((Get-String "updateAvailable") -f $tag) + "`n`n$link"
            $answer = Show-StyledMessageBox -Message $msg -Title (Get-String "info") -Buttons "YesNo" -Icon "Question"
            if ($answer -eq "Yes") {
                try { Start-Process $link } catch { Write-LogStep "Could not open browser: $($_.Exception.Message)" "WARNING" }
            }
        }
        else {
            Write-LogStep "Already up to date (v$($Global:AppVersion))" "SUCCESS"
            if (-not $Silent) {
                Show-StyledMessageBox -Message (Get-String "noUpdate") -Title (Get-String "info") -Buttons "OK" -Icon "Information"
            }
        }
    }
    catch {
        Write-LogStep "Update check failed: $($_.Exception.Message)" "WARNING"
        if (-not $Silent) {
            Show-StyledMessageBox -Message (Get-String "updateCheckFailed") -Title (Get-String "warning") -Buttons "OK" -Icon "Warning"
        }
    }
    finally {
        Update-Status "Ready"
    }
}

Write-Host "Updater Module loaded" -ForegroundColor Green
