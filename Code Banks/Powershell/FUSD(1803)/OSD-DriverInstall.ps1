﻿#ElseIF ($Model -eq ""){$DriverLocation = ""}

$Model = (gwmi win32_computersystem).model
$DriverLocation

New-Item c:\windows\temp\ -Name Drivers\$Model -Type Directory
New-Item c:\windows\temp\ -Name DriversLog -Type Directory

    IF($Model -eq "Virtual Machine"){Write-Host "VM, No drivers needed"} 
    # Alienware
        ElseIF ($Model -eq "ASM201"){$DriverLocation = "Y:\Win10x64\Alienware\ASM201"}
    # ASUS
        ElseIf ($Model -eq "Z240IC-H170-GTX950"){$DriverLocation = "Y:\Win10x64\ASUS\Z240IC-H170-GTX950"}
    # HP
    # Microsoft
        #ElseIf ($Model -eq "Surface Book"){$DriverLocation = ""}

# Copy down drivers from the network storage    
Copy-item $DriverLocation "C:\windows\Temp\Drivers\$model" -Recurse

# Install Drivers
DISM.exe /Image:%OSDisk%\ /Add-Driver /disDriver:C:\Windows\Temp\Drivers /Recurse /ForceUnsigned /logpath:C:\Windows\Temp\Driverslog\Dism.log

# Clean up
Remove-Item c:\windows\temp\Drivers\$Model\* -Recurse