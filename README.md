# ![pwsh](/icon/powershell.png) Windows/Office KMS Activator
## This script will activate your Windows and Office product using [KMS activation](https://learn.microsoft.com/en-us/windows-server/get-started/kms-client-activation-keys)

# To execute the script
1. Download kmsact.ps1
2. right click and execute with PowerShell
3. wait a few seconds before activation

_See [50bvd.com Windows/Office activation script](https://50bvd.com/posts/Activation-de-Windows-&-Office/#14-exectution-du-script)_

# Note
You have to restart the script after 180 days because the liscences are valid for 180 days. *you can create a schedule to run the `slmgr /rearm` every 180 days*

Type on your terminal `slmgr /dlv` 

You will have the information about your license as on this example :arrow_heading_down:

[![slmgrdlv](/icon/slmgrdlv.png)](https://learn.microsoft.com/en-us/windows-server/get-started/kms-client-activation-keys)

If you dont have a KMS server, see this [link](https://gist.github.com/Zibri/69d55039e5354100d2f8a053cbe64c5a#online-kms-host-address), it lists several KMS servers

Otherwise, you can use my KMS server that I leave at your disposal. -> kms.50bvd.com :white_check_mark: Actually UP
