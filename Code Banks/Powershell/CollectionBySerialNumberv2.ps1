cls


# Set SCCM Variables
$SCCMServer            = "SC-config2.staff.fusd.local"
$SCCMSiteCode          = "FS1"
$collectionName        = "YMS-Cart-B3"                        #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$limitingCollection    = "Yosemite-Active-Student-Computers"  #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



$SCCMPowerShellCmdLets = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
Import-Module $SCCMPowerShellCmdLets

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$csvFile = $scriptPath + "\YMS-Cart-B3.csv"                     #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
write-host "file name" $csvFile




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
