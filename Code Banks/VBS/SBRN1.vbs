'Batch Re-Namer Omega
'Clifford Dinel Eastman ii	
'FUSD
'This is the last version of this application that will be written in VBScript
'Have fun, use at your own risk ;>
	set wshShell= createObject("wscript.shell")	
	wshshell.popup"SGP.Batch Re-Namer Beta version 1.0",3
	
	DIM srtOldComputerName, strPrefixx, strSerialNumber , strnewComputerName ,strWhereToJoinMe
	Dim  arrSystemInfo, arrNamingDataSet
	strComputer = "."
	myiniFilePath 				= "SGP.ini"
		
		
'Constructing New system name		
	Namingdata				= NewNameDataFromINI(myiniFilePath)	'Section 002
	newComputerName			= NameConstructor(Namingdata(0),Namingdata(1),Namingdata(2),Namingdata(3))	'Section 002
	'wshshell.popup Namingdata(3)
	strWhereToJoinMe		=  0			
    
'These subrutens are critical if either one fails the script will halt
	cALL UnJoinMe(strComputer)'Section 007
	cALL RestHostName(newComputerName)'Section 007
	
		
'Update data source	
	INILineToUpdaye			= "LastStationNamed=" & Namingdata(3)
	NextStation = Namingdata(3) + 1
	NewValue 				= "LastStationNamed=" & NextStation
	return = UpDataINI(myiniFilePath,INILineToUpdaye,NewValue)
	
	
	
	
'writing joinme.vbs to local driver 
	
	call WritJoiner
	call wrtFunctionJoinUS
	call wrtSubReboot
	call wrtAutologOffRemove
	call WriteRunOnceScriptler(strWhereToJoinMe)
	
	
'Closing procedure Section 004
	call AutoLogOn	
	Call RunMeOnce	
	Call Reboot		
	
	
	
	

'Section 002
Function NewNameDataFromINI(myFilePath)
	
	reDIM NewNameDataFromINIarry(3)
	Dim	GroupLocation		'School Short name Edison High Schooll = 		EHS	
	Dim GroupName 			'Room or Cart name Room A1 = RA1  or Cart01		EHS-RA1
	Dim StationDesignation 	'Station or system might use a "S" 				EHS-RA1-S
	Dim StationNumber 		'Station number in group 1 ,2 					EHS-RA1-S01
	Dim StationCounter		'Holds current station number
	Dim StationName			'holds final name value 
	
	'Read data form data source and seed variables


	'myFilePath 				= "SGP.ini"
	mySection 				= "NamingSchema"
	myKey					= "GroupLocation"
	NewNameDataFromINIarry(0)				= ReadIni( myFilePath, mySection, myKey )
	'wshshell.popup NewNameDataFromINIarry(0)
	
	mySection 				= "NamingSchema"
	myKey					= "GroupName"
	NewNameDataFromINIarry(1)				= ReadIni( myFilePath, mySection, myKey )
	'wshshell.popup NewNameDataFromINIarry(1)
	
	mySection 				= "NamingSchema"
	myKey					= "StationDesignation"
	NewNameDataFromINIarry(2) 				= ReadIni( myFilePath, mySection, myKey )
	'wshshell.popup NewNameDataFromINIarry(2)
	
	mySection 				= "StationCounters"
	myKey					= "LastStationNamed"
	NewNameDataFromINIarry(3)				= ReadIni( myFilePath, mySection, myKey )
	'wshshell.popup NewNameDataFromINIarry(3)
	
	
	NewNameDataFromINI = NewNameDataFromINIarry
End Function
Function NameConstructor(GroupLocation,GroupName,StationDesignation,StationCounter)	
	'Construct New computer name	
	StationNumber = DDCheck(StationCounter)
	NameConstructor = GroupLocation & "-" & GroupName & "-"& StationDesignation & StationNumber
	wshShell.popup "New Name is " & NameConstructor,5
End Function
Function DDCheck(theNumber) 
	'this function inserts a leading zero on one digit numbers for consistent ordering
	if theNumber > 0 and theNumber <= 9 then
			DDCheck = "0" & TheNumber
		else
			DDCheck = Thenumber
	end if
End Function
'002



'Section 003
Function ReadIni( myFilePath, mySection, myKey )
	
	'wshshell.popup "Function ReadIni"
	'wshshell.popup myFilePath
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )
			
			'wshshell.popup strLine
            
			' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function
Function UpDataINI(myFilePath,INILineToUpdate,NewValue)
	Const ForReading = 1
	Const ForWriting = 2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(myFilePath, ForReading)

	Do Until objTextFile.AtEndOfStream
		strNextLine = objTextFile.Readline

		If strNextLine = INILineToUpdate Then
			strNextLine = NewValue
		End If
		
		strNewFile = strNewFile & strNextLine & vbCrLf
	Loop

	objTextFile.Close

	Set objTextFile = objFSO.OpenTextFile(myFilePath, ForWriting)

	objTextFile.WriteLine strNewFile
	objTextFile.Close
End Function
'End Section 003



'Section 004
Sub AutoLogOn()
	'AutoAdminLogon = "1" Enabled  (To disable Auto Logon set to value to zero)
	'DefaultUserName = "xxx"
	'DefaultPassword = "xxxx0xxxx"
	'DefaultDomainName = "xxx.xxx".  Only needed if computer has joined a domai
	
	Set objShell = CreateObject("WScript.Shell")
	
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultDomainName", 1, "REG_SZ"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultDomainName", ".\", "REG_SZ"

	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName", 1, "REG_SZ"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName","Administrator", "REG_SZ"

	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DisableCAD", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DisableCAD", 1, "REG_DWORD"

	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon", 1, "REG_SZ"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon", 1, "REG_SZ"

	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\ForceAutoLogon", 1, "REG_SZ"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\ForceAutoLogon", 1, "REG_SZ"

	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword", 1, "REG_SZ"
	objShell.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword", "F2u0s0d9", "REG_SZ"
	
End sub
Sub RunMeOnce()
	const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
	strValueName = "String Value Name"
	strValue = "C:\Users\Public\Downloads\joinme.vbs"
	Return = objReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,,strValue)
	
End sub
Sub Reboot()
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate," & "(Shutdown)}")
	Set colOSes = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	
	For Each objOS In colOSes 
		intReturn = objOS.Win32Shutdown(6)
	Next

End Sub
'Section 004



'Section 005
Function WriteRunOnceScriptler(JoinWhere)
	
	DIM txtComstr
	set wshShell= createObject("wscript.shell")	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	
	if JoinWhere = 0 then
	
		WshShell.Popup "Joining Student Domain ",2
		Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
		'***************************************************************** Make Changes in line below ***************************************************************************************
		txtComstr = "Return = Joinus(" & chr(34) & "Students.fusd.local"& chr(34) &","& chr(34) & "@@!FDDSEWDXSAQQQWQ2112" & Chr(34) &","& chr(34) & "student.join@students"& Chr(34) &","& chr(34) & "ou=Computers,ou=Yosemite_Middle,ou=Schools,dc=students,dc=fusd,dc=local" & Chr(34) &")"
		objTextStream.WriteLine(txtComstr)
		objTextStream.Close
			
	else	
	
		WshShell.Popup "Joining Staff Domain",2
		Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
		'***************************************************************** Make Changes in line below ***************************************************************************************
		txtComstr = "Return = Joinus(" & chr(34) & "Staff.fusd.local"& chr(34) &","& chr(34) & "@@!FDDSEWDXSAQQQWQ2112" & Chr(34) &","& chr(34) & "Staff.join@Staff"& Chr(34) &","& chr(34) & "ou=Computers,ou=Yosemite_Middle,ou=Schools,dc=staff,dc=fusd,dc=local" & Chr(34) &")"
		objTextStream.WriteLine(txtComstr)
		objTextStream.Close
		
	end if	 
End Function
Sub WritJoiner()
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(4)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	
	Codelines(0) = "DeleteME()"
	Codelines(1) = "Sub DeleteME()"
	Codelines(2) = "	Set objFSO = CreateObject("& chr(34) &"Scripting.FileSystemObject"& chr(34) &")"
	Codelines(3) = "	objFSO.DeleteFile "& chr(34) &"C:\Users\Public\Downloads\Joinme.vbs"& chr(34) 
	Codelines(4) = "End Sub"
	For iloop = 0 to 4

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
	
end sub
Sub wrtFunctionJoinUS()
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(46)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	
	Codelines(0) = "Function JoinUS(TheDomain,thePassword,theuser,thesiteOU) "
	Codelines(1) = "	Const JOIN_DOMAIN = 1"
	Codelines(2) = "	Const ACCT_CREATE = 2"
	Codelines(3) = "	Const ACCT_DELETE = 4"
	Codelines(4) = "	Const WIN9X_UPGRADE = 16"
	Codelines(5) = "	Const DOMAIN_JOIN_IF_JOINED = 32"
	Codelines(6) = "	Const JOIN_UNSECURE = 64"
	Codelines(7) = "	Const MACHINE_PASSWORD_PASSED = 128"
	Codelines(8) = "	Const DEFERRED_SPN_SET = 256"
	Codelines(9) = "	Const INSTALL_INVOCATION = 262144"
	Codelines(10) = "	count = 0"
	Codelines(11) = "	restartRequired = 0" 
	Codelines(12) = "	strDomain = TheDomain"				
	Codelines(13) = "	strPassword = thePassword"			
	Codelines(14) = "	strUser = theuser"					
	Codelines(15) = "	Set objNetwork = CreateObject(""WScript.Network"")"
	Codelines(16) = "	strComputer = objNetwork.ComputerName"
	Codelines(17) = "	Set objComputer = GetObject("& chr(34) &"winmgmts:{impersonationLevel=Impersonate}!\\" & chr(34) &" & strComputer & "& chr(34) &"\root\cimv2:Win32_ComputerSystem.Name="& chr(39) & chr(34) &" & strComputer & "& chr(34) & chr(39) & chr(34) & ")"
	Codelines(18) = "	ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, strPassword, strDomain & " & chr(34) & "\"  & chr(34) & " & strUser," & thesiteOU & ",JOIN_DOMAIN + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED)"
	Codelines(19) = "	If ReturnValue <> 0 Then"
	Codelines(20) = "		If ReturnValue = 1355 Then"
	Codelines(21) = "			MsgBox " & chr(34) & "The specified domain either does not exist or could not be contacted"& chr(34) &",48,"& chr(34) & "Join To Domain"& Chr(34)
	Codelines(22) = "			ElseIf ReturnValue = 2691 Then"
	Codelines(23) = "			count = 1"
	Codelines(24) = "			MsgBox " & chr(34) & "This computer is already joined to a domain"& chr(34) &",48,"& chr(34) &"Join To Domain"& chr(34)		
	Codelines(25) = "			ElseIf ReturnValue = 1326 Then"
	Codelines(26) = "			MsgBox "& chr(34) & "Logon failure: unknown user name or bad password"& chr(34) &",48,"& chr(34) & "Join To Domain"	&Chr(34) 
	Codelines(27) = "			ElseIf ReturnValue = 2224 Then"
	Codelines(28) = "			MsgBox "& chr(34) & "The account already exists"& chr(34) &",48,"& chr(34) &"Join To Domain"&Chr(34)
	Codelines(29) = "'			SearchAndDeleteAdAccount(strComputer)"
	Codelines(30) = "			ElseIf ReturnValue = 5 Then"
	Codelines(31) = "			MsgBox "& chr(34) & "Access is denied"& chr(34) &",64,"& chr(34) & "Join To Domain"&Chr(34) 
	Codelines(32) = "			ElseIf ReturnValue = 2 Then"
	Codelines(33) = "			MsgBox "& chr(34) & "Could not find the OU" & chr(34) &",64,"  & chr(34) & "Join To Domain" &Chr(34)
	Codelines(34) = "			Elseif ReturnValue = 2202 then"
	Codelines(35) = "			msgbox " & chr(34) & "The Specified Username is invalid"& chr(34) &",48,"& chr(34) & "Join To Domain" & chr(34)
	Codelines(36) = "		else"
	Codelines(37) = "			msgbox "& chr(34) &"UnTrapeerror code ="& chr(34) &"& ReturnValue ,48,"& chr(34) &"Join To Domain"& chr(34) 
	Codelines(38) = "		End If"
	Codelines(39) = "	Else"
	Codelines(40) = "		count = 1"
	Codelines(41) = "		restartRequired = 1"
	Codelines(42) = "		'MsgBox "& chr(34) &"Successfully added to the domain"& chr(34) &",64,"& chr(34) &"Join To Domain"& chr(34) &" '<<<<<<<<<<< Debug point  <<<<<<------------<<<<<<<<<"& chr(34) 
	Codelines(43) = "		AutoLogOff"
	Codelines(44) = "		Reboot"
	Codelines(45) = "	End If"
	Codelines(46) = "End Function"
	
	For iloop = 0 to 46

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
End Sub
Sub wrtSubReboot()	
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(8)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	'ReBoot Code 
	Codelines(0) = "Sub Reboot"
	Codelines(1) = "	Set objWMIService = GetObject("& chr(34) &"winmgmts:{impersonationLevel=impersonate,"& chr(34) & " & "& chr(34) & "(Shutdown)}"& chr(34) &")"
	Codelines(2) = "	Set colOSes = objWMIService.ExecQuery("& chr(34) & "SELECT * FROM Win32_OperatingSystem"& chr(34) &")"
	Codelines(3) = ""	
	Codelines(4) = "	For Each objOS In colOSes "
	Codelines(5) = "		intReturn = objOS.Win32Shutdown(6)"
	Codelines(6) = "	Next"
	Codelines(7) = ""
	Codelines(8) = "End Sub"
	
	For iloop = 0 to 8

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
	
end sub
Sub wrtAutologOffRemove()
	'AutoLog Off
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(31)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	Codelines(0) = "Sub AutoLogOff"
	Codelines(1) = "	RegLocAutoLogon = "& chr(34) & "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\"& chr(34)
	Codelines(2) = "	keyDefaultDomainName = "& chr(34) &"DefaultDomainName"& chr(34)
	Codelines(3) = "	valDefaultDomainName = " & chr(34) & chr(34) 
	Codelines(4) = "	keyDefaultUserName = "& chr(34) & "DefaultUserName" & chr(34)
	Codelines(5) = "	valDefaultUserName = " & chr(34) & chr(34) 
	Codelines(6) = "	keyDisableCAD = "& chr(34) &"DisableCAD"& chr(34) 
	Codelines(7) = "	valDisableCAD = 0"
	Codelines(8) = "	keyAutoAdminLogon = "& chr(34) & "AutoAdminLogon"& chr(34)
	Codelines(9) = "	valAutoAdminLogon = "& chr(34)&"0"& chr(34)
	Codelines(10) = "	keyForceAutoLogon = "& chr(34) &"ForceAutoLogon"& chr(34) 
	Codelines(11) = "	valForceAutoLogon = "& chr(34)&"0"& chr(34)
	Codelines(12) = "	keyDefaultPassword = "& chr(34) &"DefaultPassword"& chr(34) 
	Codelines(13) = "	valDefaultPassword = "& chr(34) & chr(34) 
	Codelines(14) = ""
	Codelines(15) = "	Set objShell = CreateObject("& chr(34) &"WScript.Shell"& chr(34) &")"
	Codelines(16) = ""
	Codelines(17) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultDomainName, 0, "& chr(34) & "REG_SZ"& chr(34) 
	Codelines(18) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultDomainName, valDefaultDomainName, "& chr(34) & "REG_SZ"& chr(34) 
	Codelines(19) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultUserName, 0, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(20) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultUserName, valDefaultUserName, "& chr(34) & "REG_SZ"& chr(34) 
	Codelines(21) = "	objShell.RegWrite RegLocAutoLogon & keyDisableCAD, 0, "& chr(34) &"REG_DWORD"& chr(34) 
	Codelines(22) = "	objShell.RegWrite RegLocAutoLogon & keyDisableCAD, valDisableCAD, "& chr(34) &"REG_DWORD"& chr(34) 
	Codelines(23) = "	objShell.RegWrite RegLocAutoLogon & keyAutoAdminLogon, 0, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(24) = "	objShell.RegWrite RegLocAutoLogon & keyAutoAdminLogon, valAutoAdminLogon, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(25) = "	objShell.RegWrite RegLocAutoLogon & keyForceAutoLogon, 0, "& chr(34) & "REG_SZ"& chr(34) 
	Codelines(26) = "	objShell.RegWrite RegLocAutoLogon & keyForceAutoLogon, valForceAutoLogon, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(27) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultPassword, 0, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(28) = "	objShell.RegWrite RegLocAutoLogon & keyDefaultPassword, valDefaultPassword, "& chr(34) &"REG_SZ"& chr(34) 
	Codelines(29) = ""	
	Codelines(30) = ""	
	Codelines(31) = "END Sub"	
	
	For iloop = 0 to 31

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
	
End sub
sub wrtGetUserID()
	'get userid
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(2)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	Codelines(0) = "Function getUserName()"
	Codelines(1) = "getUserName = InputBox("& chr(34) &"Enter your UserName "& chr(34) &")"
	Codelines(2) = "End Function"
	For iloop = 0 to 2

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
	
end sub
sub wrtGetPassword()
	'Get Password
	Dim Iloop
	Dim objFSO
	Dim objTextStream
	Dim Codelines(3)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile("C:\Users\Public\Downloads\Joinme.vbs",8, True)
	
	Codelines(0) = "Function getPassword()"
	Codelines(1) = "getPassword = InputBox("& chr(34) &"Enter your Password "& chr(34) &")"
	Codelines(2) = "End Function"
	Codelines(3) ="" 
	For iloop = 0 to 3

		objTextStream.WriteLine(Codelines(iloop))
	Next 	
		objTextStream.Close
	
end sub
'005

'Section 007
Sub UnJoinMe(strComputer)
	
	Set objNetwork = CreateObject("WScript.Network")
	
	strComputer = objNetwork.ComputerName
	Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")
	ReturnValue = objComputer.unJoinDomainOrWorkGroup(null,null,0)
	
	If ReturnValue = 5 Then
		MsgBox "Access is denied",64,"unJoin from Domain programe terminated" 
		wscript.quit
	End If
	
End sub
Sub RestHostName(newComputerName)
	
	Set objWMIService = GetObject("Winmgmts:root\cimv2")

	For Each objComputer in objWMIService.InstancesOf("Win32_ComputerSystem")

       Return = objComputer.rename(newComputerName,Password,Username)
		If Return <> 0 Then
          WScript.Echo "Rename failed. Error = " & Err.Number
		  wscript.quit
        End If

	Next 
end Sub
'007