if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}

$WindowsKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
$OfficeKey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
$KMSserver = "kms.digiboy.ir"

Write-Host "
  ___                       ___  
 (o o)                     (o o) 
(  V  ) MS KMS Activation (  V  )
--m-m-----------------------m-m--
    https://github.com/50bvd
" -ForegroundColor Green

$Answer = Read-Host "Do you want to install Windows key, Office key or both? (W/O/B)"

if ($Answer -eq "W") {
    slmgr /ipk $WindowsKey
    slmgr /skms $KMSserver
    slmgr /ato
} elseif ($Answer -eq "O") {
    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /inpkey:$OfficeKey
    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /sethst:$KMSserver
    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /act
} elseif ($Answer -eq "B") {
    slmgr /ipk $WindowsKey
    slmgr /skms $KMSserver
    slmgr /ato

    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /inpkey:$OfficeKey
    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /sethst:$KMSserver
    cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /act
} else {
    Write-Host "Invalid answer, please type W, O or B."
}
