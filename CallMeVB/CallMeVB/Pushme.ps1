
#$MyInvocation.MyCommand.Path
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$csvFile = $scriptPath + "\MHS-LAB-209.csv"  #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#write-host "file name" $csvFile
$collectionName = "MHS-LAB-209"              #<<<<<<<<<<<<<<<<<<<<<<<<<<  Change Me !!!!   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$SerialList = Import-Csv -Path $csvFile
$SerialList