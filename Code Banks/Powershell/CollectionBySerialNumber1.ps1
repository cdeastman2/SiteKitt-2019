cls
# This script will populate a collection based on a list of PC
# serial numbers.
#
# AuthorKenneth Merenda
# Date:November 8, 2016
#
###################################################################
# CUSTOMIZE THIS SECTION
#
# Enter the hostname of the primary site server
$SCCMServer = "SC-config2.staff.fusd.local"
#
#
# The 3-character SCCM Site Code
$SCCMSiteCode = "FS1"
#
#
# Full path to the SCCM PowerShell Commandlets file
# Example:$SCCMPowerShellCmdLets = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
$SCCMPowerShellCmdLets = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"

#
#
# Enter the full path to the .csv file
# Example:csvFile = "\\rhs--rnd\Master_Profiles\Sites\Yosemite_Middle\Group_Profiles\A5\SerialNumbers.csv"

#$MyInvocation.MyCommand.Path
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition


$csvFile = $scriptPath + "\YMS-CART-N21.csv"  #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
write-host "file name" $csvFile
#
#
# Enter the name of the target collection to be used or created
$collectionName = "YMS-CART-N21"              #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#
#
# Enter the name of the limiting collection that will be used to create the target collection, if the target collection does not already exist.
$limitingCollection = "Yosemite Student Computers"
#
#
# END CUSTOMIZATIONS
###################################################################
#Import the SCCM PowerShell Commandlets
Import-Module $SCCMPowerShellCmdLets
# Connect to the SCCM Site
Set-Location $SCCMSiteCode':'

# Determine if target collection already exists
$collection = Get-CMCollection -Name $collectionName
# If target collection does not exist, create it
if (!$collection) {
Write-Host "Creating Collection"
New-CMDeviceCollection -Name $collectionName -LimitingCollectionName $limitingCollection
$collection = Get-CMCollection -Name $collectionName
}



# Import serial numbers from CSV
$SerialList = Import-Csv -Path $csvFile

# Loop through serial numbers
foreach ($SerialNumber in $SerialList) {
# save serial number to a variable
$deviceSerialNumber = $SerialNumber.SerialNumber

# create collection membership rule
Write-Host "Creating a membership rule for $deviceSerialNumber"
Add-CMDeviceCollectionQueryMembershipRule -Collection $collection -RuleName "Serial Number is $deviceSerialNumber" -QueryExpression "select * from SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM_PRODUCT on SMS_G_System_COMPUTER_SYSTEM_PRODUCT.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM_PRODUCT.IdentifyingNumber = ""$deviceSerialNumber"""
}
