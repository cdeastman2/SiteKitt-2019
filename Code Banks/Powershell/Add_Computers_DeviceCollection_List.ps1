$computers = Get-Content "C:\users\admin-serrato\Desktop\ftm.txt"
$DeviceCollection = 'FT Miller Room 27'

foreach($computer in $computers)
{ Add-CMDeviceCollectionDirectMembershipRule  -CollectionName $DeviceCollection -ResourceId $(get-cmdevice -Name $computer).ResourceID
}