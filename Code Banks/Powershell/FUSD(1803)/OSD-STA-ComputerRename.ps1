$SN = (gwmi win32_BIOS).serialnumber
$HostName = $env:computername

# Serial Number Permute
$SN = $SN.Replace(' ','')
$String = $SN.Length
If ($String -gt 10)
    {
        $NewHostName = 'STA-' + ($SN.Substring($SN.Length - 11))
    }
Else 
    {
        $NewHostName = 'STA-' + ($SN)
    }
Write-Host -fore Red "This device has a gerenic computer name: '$HostName'..."
Write-Host -fore Green "It will get renamed to: '$NewHostName' and will continue with domain join after rebooting."
Rename-Computer $NewHostName
