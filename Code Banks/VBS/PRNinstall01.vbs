'Batch Re-Namer Omega
'Clifford Dinel Eastman ii	
'FUSD
'This is the last version of this application that will be written in VBScript
'Have fun, use at your own risk ;>
	set wshShell= createObject("wscript.shell")	
	
	
	DIM srtOldComputerName, strPrefixx, strSerialNumber , strnewComputerName ,strWhereToJoinMe
	Dim  arrSystemInfo, arrNamingDataSet
	strComputer = "."
	MyOS = CheckOS(strComputer)
	
	
	
	
	
	'Install Default printer or univercal print privers
	
	


	dim oShell, strPath
	set oShell= Wscript.CreateObject("WScript.Shell")
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colInstalledPrinters =  objWMIService.ExecQuery("Select * from Win32_Printer Where Network = FALSE")

	'For Each objPrinter in colInstalledPrinters
    'objPrinter.Delete_
	'Next
	
	
	
	
	
	if MyOS(1) = "64-bit" then
			'wshShell.popup"64-bit",2 'For Debug
			strPath = """\\rhs--rnd\Master Profiles\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x64-5.5.0.12834\Install.exe"""
		else
			'wshShell.popup"32-bit",2 'For Debug
			strPath = """\\rhs--rnd\Master Profiles\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x32-5.5.0.12834\Install.exe"""
	end if
	
	rem wshShell.popup strPath
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-01 /n""DHS-MOF-PRN-01""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-02 /n""DHS-MOF-PRN-02""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-03 /n""DHS-MOF-PRN-03""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-04 /n""DHS-MOF-PRN-04""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-05 /n""DHS-MOF-PRN-05""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-06 /n""DHS-MOF-PRN-06""", 1, true
	oShell.Run strPath & " /q /h /dst /smDHS-MOF-PRN-07 /n""DHS-MOF-PRN-07""", 1, true
	
	
	
	'oShell.Run strPath & " /q /h /dst /sm172.26.107.215 /n""LeavenWorth LMC 4100 Printer""", 1, true
	'oShell.Run strPath & " /q /h /dst /sm172.26.110.118 /n""LeavenWorth LMC LJ 2100 Printer""", 1, true
	
	
Function CheckOS(strComputer)
	set wshShell= createObject("wscript.shell") 'for Pop up windows 
	WshShell.Popup "Getting system OS information ",2 ' Script launch verifactioin 

		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
			For Each objItem in colItems
				
				ThisOS = objItem.Name
				ThisArchitecture = objItem.OSArchitecture
				CheckOS = Array(ThisOS,ThisArchitecture)
			Next
	WshShell.Popup "Getting system  OS information ",2 		
end function