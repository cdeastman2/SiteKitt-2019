Dim CSVfilepath,Accfilepath
Dim DataFile,dataset
Dim On_the_return
Dim strComputer
Dim TheDataPoints
Dim Error_Info
	

set wshShell= createObject("wscript.shell")
	WshShell.Popup "Starting Script ",2

	
	strComputer = "."
	Dataset = getSCCMImportData(strComputer)
'
	CSVfilepath = "grouptest.csv"
	Call Do_CSV_Data
'	
	Accfilepath = "SightKitt2018.accdb"
	
	On_the_return = DoAccesslog(Dataset)
	 WshShell.Popup "Done ",5
		
		
		
		
		
		
		
		
		
		
'Function and Sub 



Sub Do_CSV_Data()
		
		WshShell.Popup "SerialNumber "  & Dataset (0) ,5
		DataFile = Dataset(0)
		On_the_return = WrightData(DataFile,CSVfilepath)
		wshShell.Run("notepad " & CSVfilepath)

ENd Sub 
	
Function  GetSerialNumber(ThisComputer)
	Set objWMIService = GetObject("winmgmts:\\" & ThisComputer & "\root\CIMV2")
	Set colBIOS = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
	
	For Each objBIOS in colBIOS
		GetSerialNumber = objBIOS.SerialNumber
	Next
end Function

Function WrightData(Data2wright,File2write2)
	Dim objFSO 'file system object
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	If (objFSO.FileExists(File2write2)) Then
			Rem WScript.Echo("File exists!")
			Set objTextStream = objFSO.OpenTextFile(File2write2,8, True)
			objTextStream.WriteLine(Data2wright & "," )
	Else
			Rem WScript.Echo("File does not exist!")
			Set objTextStream = objFSO.OpenTextFile(File2write2,8, True)
			objTextStream.WriteLine("SerialNumber" & "," )
			objTextStream.WriteLine(Data2wright & "," )
	End If
	
    

	

End Function

Function getSCCMImportData(strComputer)
	On Error resume next
	Dim DataPoints(10)
	Set objWMIService 		= GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	set colComsysPro 		= objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
	Set colComsys 			= objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	Set colBIOS				= objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
	Set colItems 			= objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter",,48) 
	
	
	For Each objBIOS in colBIOS
		DataPoints(0) = trim(objBIOS.SerialNumber)
		'WshShell.Popup DataPoints(0),6 'Debug
	Next
	
	For Each Product in colComsysPro
		DataPoints(1) = product.UUID
		'WshShell.Popup DataPoints(1),6 'Debug
	next 
	
	For Each objComsys in colComSys
	
		DataPoints(2)		 	= objComsys.DNSHostName & "." & objComsys.Domain 
		'WshShell.Popup DataPoints(2),6 'Debug
		
		DataPoints(3) 			= objComsys.model
		'WshShell.Popup DataPoints(3),5 'Degug
					
		DataPoints(4) 			= objComsys.Manufacturer
		'WshShell.Popup DataPoints(4),6 'Debug
		
		DataPoints(6) 			= objComsys.SystemFamily
		'WshShell.Popup DataPoints(6),6 'Debug
		
		DataPoints(7) 			= objComsys.SystemType
		'WshShell.Popup DataPoints(7),6 'Debug
		
		DataPoints(8) 			= objComsys.Description
		'WshShell.Popup DataPoints(8),6 'Debug
		
	Next
	
	NicCout = 0
	For Each objItem in colItems
		if objitem.NetConnectionID <> vbnull Then '
			if objitem.Manufacturer <> "Microsoft" Then 
				if NicCout = 0 then
					if objItem.MACAddress <> "00:50:B6:06:B1:FE" then
					DataPoints(5) = DataPoints(5) + objItem.MACAddress 
						NicCout =  NicCout + 1
					end if	
				else
					DataPoints(5) = DataPoints(5) + "," + objItem.MACAddress
					
				end if	
			End If
		End If	
	
	Next
	'WshShell.Popup DataPoints(5),6 'Debug
	
	
	
	getSCCMImportData 	= DataPoints
	Set objWMIService 	= Nothing
	Set colComsysPro	= Nothing
	Set colComsys 		= Nothing
	Set colBIOS 		= Nothing
	On Error GoTo 0
	
	
End Function

Function DoAccesslog(TheDataPoints)
	On Error resume next
	
	DIM datasource,connString
	Dim strSerialNumber,strUUID,strComputerName,strManufacturer,strmodel,strMACS,Main_SCCM_Group,strWarranty
	
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet  = CreateObject("ADODB.Recordset")
		
	strSerialNumber	= TheDataPoints(0)
	strUUID			= TheDataPoints(1)
	strComputerName	= TheDataPoints(2)
	strManufacturer = TheDataPoints(3)
	strmodel		= TheDataPoints(4)	
	strMACS			= TheDataPoints(5)
	Main_SCCM_Group = TheDataPoints(6)
	
	If strWarranty 	= "" then 
		strWarranty		= "11/29/2000"
	end if
	LastDate 		= date()
	
	
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; data source = SightKitt2018.accdb"
	objRecordSet.Open "insert into CPUInvo ( SerialNumber,UUID,ComputerName,Manufacturer,model,MACS,Main_SCCM_Group ) values(" & chr(39) & strSerialNumber & Chr(39) & "," & Chr(39) & strUUID & Chr(39) &"," & Chr(39) & strComputerName & Chr(39) & "," & Chr(39) & strManufacturer & Chr(39) & "," & Chr(39) & strmodel & Chr(39) & "," & Chr(39) & strMACS & Chr(39) & "," & Chr(39) & Main_SCCM_Group & Chr(39) & ")", objConnection, adOpenStatic, adLockOptimistic
	
	
	Rem objRecordSet.Open "insert into CPUInvo (SerialNumber,UUID,ComputerName,Manufacturer,model,MACS,Main_SCCM_Groupe,)" &_
	Rem " values (" & Chr(39) & strSerialNumber & Chr(39) & "," & Chr(39) & strUUID & Chr(39) & "," & Chr(39) & strComputerName & Chr(39) & "," & Chr(39) & strManufacturer & Chr(39) & "," & Chr(39) & strmodel & Chr(39) & "," & Chr(39) & strMACS & Chr(39) & "," & Chr(39) & Main_SCCM_Group & Chr(39) & ")", objConnection, adOpenStatic, adLockOptimistic
	
	
	
	Rem objRecordSet.Open "insert into CPUInvo (SerialNumber,UUID,ComputerName,Manufacturer,model,MACS,Main_SCCM_Group) values ( '"& strSerialNumber & "' , '" & strUUID &"','" & strComputerName & "', '" & strManufacturer &" ' , '" & strmodel & "' , '" & strMACS & "', '" & Main_SCCM_Group & "' )", objConnection, adOpenStatic, adLockOptimistic
	
	
	
	
	
	
	
	If Err.Number = 0 Then
		Rem WScript.Echo "It worked!" 
		DoAccesslog = true
	Else
		WScript.Echo "Error:"
	WScript.Echo Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
		DoAccesslog = Err.Number
		Err.Clear
	End If
	
	On Error GoTo 0
	
	
end function

Function DoAccesslogUpDate(strComputer)
	'On Error resume next
	
	DIM datasource,connString
	Dim strSerialNumber,strUUID,strComputerName,strManufacturer,strmodel,strMACS,Main_SCCM_Group,strWarranty

	Const adOpenStatic = 3
	Const adLockOptimistic = 3

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")
	
	strSerialNumber	= TheDataPoints(0)
	strUUID			= TheDataPoints(1)
	strComputerName	= TheDataPoints(2)
	strManufacturer = TheDataPoints(3)
	strmodel		= TheDataPoints(4)	
	strMACS			= TheDataPoints(5)
	Main_SCCM_Group = TheDataPoints(6)
	
	LastDate 		= date()
	
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; " & "data source = Q:\labKit.accdb"
	
	
	'                 "UPDATE Computers Set Department = 'None' Where Department = 'Human Resources'",
	objRecordSet.Open "UPDATE CPUInvo Set "&_
	"ComputerName = "  & Chr(39) & strComputerName & Chr(39) &_ 
	",Lastdate = "  & Chr(39) & LastDate & Chr(39) &_ 
	",Main_SCCM_Group = "  & Chr(39) & Main_SCCM_Group & Chr(39) &_ 
	",Log = 'UPDATE' "&_
	"Where SerialNumber = "& Chr(39) & strSerialNumber & Chr(39) & "", objConnection, adOpenStatic, adLockOptimistic






End function
	
Sub getdata()	
	TheDataPoints = getSCCMImportData(theComputer)
	Error_Info = DoAccesslog(TheDataPoints)
	
	
	If Error_Info <> 0 then 
	
		if Error_Info = -2147467259 then
			wshShell.popup "Updating Records",5
			Error_Info = DoAccesslogUpDate(TheDataPoints)
		End if
		wshshell.popup"Script is Done ",3
	end if
end sub
