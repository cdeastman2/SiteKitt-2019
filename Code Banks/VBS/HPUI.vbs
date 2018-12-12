'Clifford Dinel Eastman ii	

	set wshShell= createObject("wscript.shell")	
	'wshShell.popup"FUSD Technology Services Site Printer Install"',5
	
	Dim  arrSystemInfo, arrNamingDataSet
	strComputer = "."
	
	MyOS = CheckOS(strComputer)
	'Install Default printer or univercal print privers
	
	DIM oShell, strPath
	set oShell= Wscript.CreateObject("WScript.Shell")
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colInstalledPrinters =  objWMIService.ExecQuery("Select * from Win32_Printer Where Network = FALSE")

	rem For Each objPrinter in colInstalledPrinters
    	rem objPrinter.Delete_
	rem Next


Sub Install_Site_Ricoh()	
	if MyOS(1) = "64-bit" then
			'wshShell.popup"64-bit",2 'For Debug
			strPath = """\\iscsi-ex02\Installations\Site Profiles\MHSR\Leavenworth_Elementary\Printers\Leavenworth_Ricoh_Pro_8100s_TCP_IP_RICOH_PCL6_UniversalDriver_V4.4_64Bit_1.1.0.exe"""
			oShell.Run strPath
		else
			'wshShell.popup"32-bit",2 'For Debug
			strPath = """\\iscsi-ex02\Installations\Site Profiles\MHSR\Leavenworth_Elementary\Printers\Leavenworth_Ricoh_Pro_8100s_TCP_IP_RICOH_PCL6_UniversalDriver_V4.4_32Bit_1.1.0.exe"""
			oShell.Run strPath
	end if
End Sub


	
	if MyOS(1) = "64-bit" then
			'wshShell.popup"64-bit",2 'For Debug
			strPath = """\\iscsi-ex02\Installations\Site Profiles\PRINTERS\Universal\HP Universal Print Driver\pcl6-x64-5.6.0.14430\Install.exe"""
		else
			'wshShell.popup"32-bit",2 'For Debug
			strPath = """\\iscsi-ex02\Installations\Site Profiles\PRINTERS\Universal\HP Universal Print Driver\pcl5-x32-5.5.0.12834\Install.exe"""
	end if
	
	
	
	
	rem wshShell.popup strPath
	oShell.Run strPath & " /q /h /dst /sm172.26.57.51 /n""SMS LMC Printer #01""", 1, true
	oShell.Run strPath & " /q /h /dst /sm172.26.55.61 /n""SMS LMC Printer #02""", 1, true
	oShell.Run strPath & " /q /h /dst /sm172.26.56.85 /n""SMS LMC Printer #03""", 1, true
	
	

	
Function CheckOS(strComputer)
	set wshShell= createObject("wscript.shell") 'for Pop up windows 
	'WshShell.Popup "Getting system  OS information ",2 ' Script launch verifactioin 

		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
			For Each objItem in colItems
				
				ThisOS = objItem.Name
				ThisArchitecture = objItem.OSArchitecture
				CheckOS = Array(ThisOS,ThisArchitecture)
			Next
	'WshShell.Popup "Getting system  OS information ",2 		
end function