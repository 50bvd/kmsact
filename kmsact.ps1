#adm exec verif
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "& {Set-ExecutionPolicy Bypass -Scope Process -Force; & '" + $myinvocation.mycommand.definition + "'}"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}
#var
$WindowsKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
$OfficeKey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
$KMSserver = "50bvd.com"
$Answer = Read-Host "Do you want to install Windows key, Office key or both? (W/O/B)"
#wonderfull ascii art
Write-Host "
  ___                       ___  
 (o o)                     (o o) 
(  V  ) MS KMS Activation (  V  )
--m-m-----------------------m-m--
    https://github.com/50bvd
" -ForegroundColor Green
#rearm office gvlk
function Initialize-OfficeLicense {
    Push-Location
    Set-Location "C:\Program Files\Microsoft Office\Office16"
    cmd.exe /c "for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:""..\root\Licenses16\%x"""
    Pop-Location
}
#menu
if ($Answer -eq "W") { #windows act
    slmgr /ipk $WindowsKey
    slmgr /skms $KMSserver
    slmgr /ato
} elseif ($Answer -eq "O") { #office act
    Initialize-OfficeLicense
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /inpkey:$OfficeKey
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /sethst:$KMSserver
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /act
} elseif ($Answer -eq "B") { #full act
    slmgr /ipk $WindowsKey
    slmgr /skms $KMSserver
    slmgr /ato

    Initialize-OfficeLicense
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /inpkey:$OfficeKey
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /sethst:$KMSserver
    cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /act
} else {
    Write-Host "Invalid answer, please type W, O or B."
}
