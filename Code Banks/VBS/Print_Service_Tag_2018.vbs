Option Explicit
'==========================================================
' LANG 					: VBScript
' NAME 					: Print_Service_Tage2.vbs
' AUTHOR				: Clifford Dinel Eastman ii	
' Mail					: CDEASTM4FUSD@gmail.com
' VERSION 				: 2.0.1
' DATE 					: 12/12/2016
' Description			: 
'						: 
' COMMENTS 				: 


' CONTRIBUTORS			:

' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
' LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
' CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
'==========================================================
'System wide Variable Declarations and Constants
'==========================================================
	Dim wshShell 	'Shell
	Dim strComputer 'The Computer being acted on 1
		
	Dim TheProsess	'The current  process being executed 
	Dim Error_Info	'Error reporting 
	Dim MyOS
	DIM verbos		'Display messages
	Dim ReturnOfFunk
	Dim Data_Check '
	
'==========================================================
' Program's Variable Declarations and Constants
'==========================================================
	'Tag Data points 
	DIM TagData(5)
	Dim txtIncident				'TagData(0)
	Dim	txtSite					'TagData(1)
	Dim txtLocation				'TagData(2)
	Dim txtInvoTrackingNumber	'TagData(3)
	Dim txtTech					'TagData(4)
	Dim txtFullname
	Dim txtNotes				'TagData(5)
		
	Dim Site_Name
	DIM Default_printer
	Dim Site_Ricoh
	Dim the_priter





'==========================================================	
' Program's Default settings
'==========================================================	
	
	verbos 				= false
	txtTech 				= ""
	txtFullname				= ""
	txtIncident 			= ""
	txtSite 				= ""
	txtLocation 			= ""
	txtInvoTrackingNumber	= ""
	txtNotes 				= ""
	strComputer				= "."
	



'==========================================================
' Main Body
'==========================================================
	set wshShell= Wscript.CreateObject("WScript.Shell")
	If verbos = true then 	
		wshshell.popup"Starting Program ",3 
	End if
	On Error resume next    	
	'GetOptions 'Get Options from user
'==========================================================
	
	txtSite = Get_site_name_by_gateway_hostname()
	If verbos = true then 
		wshshell.popup txtSite,5
	End if
	
	Default_printer = getDefaultPrinter()
	If verbos = true then 
		wshshell.popup Default_printer,5
	End if
	
	txtFullname = Get_My_name 
	txtTech = txtFullname(0) & " " & txtFullname(1)
	If verbos = true then 
		wshshell.popup txtTech,5
	End if
	
	
	
	
	
	
	'If txtTech ="Clifford D. Eastman ii" then 
	txtTech = Get_Tech_name(txtTech)
		If verbos = true then 	
			wshshell.popup "txtTech = " & txtTech ,3 
		End if
	'end if
	
	If txtIncident ="" then 
	txtIncident = Get_Incident_Number()
		If verbos = true then 	
			wshshell.popup "txtIncident = " & txtIncident ,3 
		End if
	end if
	
	'If txtSite ="" then 
	txtSite = Input_site(txtSite)
		If verbos = true then 	
			wshshell.popup "txtSite = " & txtSite ,3 
		End if
	'end if
	
	If txtLocation ="" then 
	txtLocation = Get_The_Location()
		If verbos = true then 	
			wshshell.popup "txtLocation = " & txtLocation ,3 
		End if
	end if
	
	If txtInvoTrackingNumber ="" then 
	txtInvoTrackingNumber = Get_Imvontory_Tracking_Number()
		If verbos = true then 	
			wshshell.popup "txtInvoTrackingNumber = " & txtInvoTrackingNumber ,3 
		End if
	end if
	
	
	txtNotes = Enter_any_Notes()
	
	

	ReturnOfFunk = Print_SerTag(txtTech,txtIncident,txtSite,txtLocation,txtInvoTrackingNumber,txtNotes)
'==========================================================	
	If verbos = true then 	
	 wshshell.popup" Done ",3
	End if
	'exit 
'==========================================================	

Function Get_Tech_name(txt_Tech_name)' Vertion 1.0 Aplha
	
	
	Data_check = False

	Do While Data_check = False	
		txt_Tech_name 				= Trim(Inputbox("Enter Your Name ","Techs Name",txt_Tech_name )) 
		If 	txt_Tech_name = "" Then
			txt_Tech_name 		= Inputbox(" Data enter error pleace re-enter Enter Your Name or Q to quit  ","Input Required")
			If UCase (txt_Tech_name) = "Q" Then 
				
				WScript.Quit
				
			End If
		Else 
					
			Data_check = True
			
		End If 	
	Loop 			
	Get_Tech_name = txt_Tech_name
'==========================================================	
End Function 
	
Function Get_Incident_Number()' Vertion 1.0 Aplha
	DIM txt_Incident
	Data_Check = False
	Do While Data_check = False	
		txt_Incident 			= Trim(Inputbox("Enter Incident # ","Input Required"))
		If 	txt_Incident = "" Then
			txt_Incident 		= Trim(Inputbox(" Data enter error pleace re-enter Incident # or Q to quit  ","Input Required"))
			If UCase (txt_Incident) = "Q" Then 
				
				WScript.Quit
				
			End If
		Else 
					
			Data_check = True
			
		End If 	
	Loop 	
	Get_Incident_Number = txt_Incident
End Function 	

Function Input_site(txt_Site_here)' Vertion 1.0 Aplha
	
	
	Data_check = False
	Do While Data_check = False	
		txt_Site_here 				= Trim(Inputbox("Enter Site ","Input Required",txt_Site_here)) 
		If 	txt_Site_here = "" Then
			txt_Site_here 		= Inputbox(" Data enter error pleace re-enter Site or Q to quit  ","Input Required")
			If UCase (txt_Site_here) = "Q" Then 
				
				WScript.Quit
				
			End If
		Else 
					
			Data_check = True
			
		End If 	
	Loop 		
	Input_site = txt_Site_here
End Function

Function Get_The_Location()	' Vertion 1.0 Aplha
	Dim txt_Location
	Data_Check = False
	Do While Data_check = False	
		txt_Location 			= Trim(Inputbox("Enter Room Location ","Input Required"))
		If 	txt_Location = "" Then
			txt_Location 		= Inputbox(" Data enter error pleace re-enter Room Location or Q to quit  ","Input Required")
			If UCase (txt_Location) = "Q" Then 
				
				WScript.Quit
				
			End If
		Else 
					
			Data_check = True
			
		End If 	
	Loop
	Get_The_Location = txt_Location
End Function 	

Function Get_Imvontory_Tracking_Number()' Vertion 1.0 Aplha
	Dim txt_InvoTrackingNumber
	Data_check = False
	Do While Data_check = False	
			txt_InvoTrackingNumber	= Trim(Inputbox("Enter Inventory tracking number ","Input Required")) 
		If 	txt_InvoTrackingNumber = "" Then
			txt_InvoTrackingNumber 		= Inputbox(" Data enter error pleace re-enter Inventory tracking number or Q to quit  ","Input Required")
			If UCase (txt_InvoTrackingNumber) = "Q" Then 
				
				WScript.Quit
				
			End If
		Else 
					
			Data_check = True
			
		End If 	
	Loop 
	Get_Imvontory_Tracking_Number = txt_InvoTrackingNumber
End Function

Function Enter_any_Notes()' Vertion 1.0 Aplha	
		txtNotes 			= Trim(Inputbox("Enter Any Notes  ","Input Required"))
		Enter_any_Notes = txtNotes
		
End Function 




Function getDefaultPrinter()' Vertion 1.0 Aplha
	' Export the default printer separately
	
	Dim objWMIService
	Dim colPrinters
	Dim objPrinter 
	Dim strText
	
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer Where Default = TRUE")

	For Each objPrinter in colPrinters
		strText = objPrinter.Name
	Next
	getDefaultPrinter = strText
End function


Function Set_default_Printer(ScriptedPrinter)' Vertion 1.0 Aplha
		Dim objWMIService
		Dim colInstalledPrinters
		Dim objPrinter
		
		
		

		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set colInstalledPrinters =  objWMIService.ExecQuery ("Select * from Win32_Printer Where Name = 'ScriptedPrinter'")
		For Each objPrinter in colInstalledPrinters
			objPrinter.SetDefaultPrinter()
		Next

end function

Function Get_My_name () ' Vertion 1.0 Aplha
	DIm objSysInfo
	Dim strUser
	Dim objUser
	Dim My_NAme_IS (1)
	
	
	Set objSysInfo = CreateObject("ADSystemInfo")
	strUser = objSysInfo.UserName
	Set objUser = GetObject("LDAP://" & strUser)
	My_NAme_IS(0) 	= objUser.givenName
	My_NAme_IS(1)	=objUser.SN
	Get_My_name = My_NAme_IS
ENd function






Function Print_SerTag(txtTech,txtIncident,txtSite,txtLocation,txtInvoTrackingNumber,txtNotes)' Vertion 1.0 Aplha
	Dim objWord,objDoc,objSelection ' Word objects"
	
	Set objWord 		= CreateObject("Word.Application")
	objWord.Caption 	= "Information Technology Service Request"
	objWord.Visible 	= True '<<<<<<<<< Debug point <<<<<<<<< 
	Set objDoc			= objWord.Documents.Add()
	Set objSelection 	= objWord.Selection

	objSelection.ParagraphFormat.Alignment = 1
	objSelection.Font.Name = "Arial"
	objSelection.Font.Size = "24"
	objSelection.TypeText "Information Technology Service Tag"
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()



	objSelection.ParagraphFormat.Alignment = 0
	objSelection.Font.Size = "14"
	objSelection.TypeText "Date: "
	objSelection.Font.Bold = True
	objSelection.TypeText "" & Date()
	objSelection.Font.Bold = false
	objSelection.TypeParagraph()

	objSelection.Font.Size = "14"
	objSelection.TypeText "Tech: "
	objSelection.Font.Bold = True
	objSelection.TypeText txtTech
	objSelection.Font.Bold = false
	objSelection.TypeParagraph()

	objSelection.Font.Size = "14"
	objSelection.TypeText "Incident#: "
	objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
	objSelection.Font.Bold = false
	objSelection.Font.Size = "14"
	objSelection.TypeText "*" & txtIncident & "*" 
	objSelection.Font.Bold = false
	objSelection.Font.Name = "Arial"
	objSelection.TypeParagraph()

	objSelection.Font.Name = "Arial"
	objSelection.Font.Size = "14"
	objSelection.TypeText "Site: "
	objSelection.Font.Bold = True
	objSelection.TypeText txtSite
	objSelection.Font.Bold = false
	objSelection.TypeParagraph()

	objSelection.Font.Size = "14"
	objSelection.TypeText "Location: "
	objSelection.Font.Bold = True
	objSelection.TypeText txtLocation
	objSelection.Font.Bold = false
	
	objSelection.TypeParagraph()



	objSelection.Font.Size = "14"
	objSelection.TypeText "Serial#/DPN: "
	objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
	objSelection.Font.Bold = false
	objSelection.Font.Size = "14"
	objSelection.TypeText "*" & txtInvoTrackingNumber &"*" 
	
	objSelection.Font.Bold = false
	objSelection.Font.Name = "Arial"
	objSelection.TypeParagraph()

	
	objSelection.ParagraphFormat.Alignment = 1
	objSelection.Font.Name = "Arial"
	objSelection.Font.Size = "10"
	objSelection.TypeText("Please refer to incident ")
	objSelection.Font.Bold = True
	objSelection.TypeText("#" & txtIncident)
	objSelection.Font.Bold = false
	objSelection.TypeText(" when E-Mailing budget information to ")
	objSelection.Font.Bold = True
	objSelection.TypeText("IT.helpdesk@Fresnounified.org")
	objSelection.Font.Bold = false
	objSelection.TypeParagraph()
	
	
	if txtNotes <> "" then
		objSelection.ParagraphFormat.Alignment = 0
		objSelection.Font.Bold = false
		objSelection.Font.Size = "14"
		objSelection.TypeText "========= Note ============================="
		objSelection.TypeParagraph()
		objSelection.Font.Bold = True
		objSelection.TypeText txtNotes
		objSelection.Font.Bold = false
		objSelection.TypeParagraph()
	
	end if
	

	'objDoc.PrintOut()
	'objDoc.SaveAs(txtIncident &".doc")
	'objWord.Quit
	
End Function
'==========================================================
'System Functions,Subroutines and Dependencies
'==========================================================
		
'====================='Get_SCCM_Import_Data v3.0'====================='
Function Get_SCCM_Import_Data(strComputer)
	On Error resume next

	Dim DataPoints(10)
	Dim	objWMIService
	Dim colComsysPro,Product
	DIm colComsys,objComsys
	Dim colBIOS,objBIOS
	Dim colNICItems,objItem,NicCout
	
	
	Set objWMIService 		= GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	set colComsysPro 		= objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
	Set colComsys 			= objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	Set colBIOS				= objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
	Set colNICItems 		= objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter",,48) 
	
	
	For Each objBIOS in colBIOS
		DataPoints(0) = trim(objBIOS.SerialNumber)
		
		If verbos = true then
			WshShell.Popup "objBIOS.SerialNumber = " & DataPoints(0),6 'Debug
		end if
		
	Next
	
	For Each Product in colComsysPro
		DataPoints(1) = product.UUID
		
		If Verbos = true then
			WshShell.Popup  "product.UUID = " & DataPoints(1),6 'Debug
		end if
		
	next 
	
	For Each objComsys in colComSys
	
		DataPoints(2)		 	= objComsys.DNSHostName 
		
		If verbos = true then 
			WshShell.Popup  "DNSHostName = " & DataPoints(2),6 'Debug
		end if
		DataPoints(3) 			= objComsys.model
		If verbos = true then 
			WshShell.Popup  "objComsys.model =" & DataPoints(3),5 'Degug
		end if			
		DataPoints(4) 			= objComsys.Manufacturer
		If verbos = true then 
			WshShell.Popup  "objComsys.Manufacturer = " & DataPoints(4),6 'Debug
		end if
		DataPoints(6) 			= objComsys.SystemFamily
		If verbos = true then 
			WshShell.Popup  "objComsys.SystemFamily = " & DataPoints(6),6 'Debug
		end if
		DataPoints(7) 			= objComsys.SystemType
		If verbos = true then 
			WshShell.Popup  "objComsys.SystemType = " & DataPoints(7),6 'Debug
		end if
		DataPoints(8) 			= objComsys.Description
		If verbos = true then 
			WshShell.Popup  "objComsys.Description = " & DataPoints(8),6 'Debug
			
		DataPoints(9) 			= objComsys.Domain 
		If verbos = true then 
			WshShell.Popup  "objComsys.Domain  = " & DataPoints(8),6 'Debug
		end if
			
			
		end if
	Next
	
	NicCout = 0
	For Each objItem in colNICItems
		if objitem.NetConnectionID <> vbnull Then '
			if objitem.Manufacturer <> "Microsoft" Then 
				if NicCout = 0 then
					'if objItem.MACAddress <> "00:50:B6:06:B1:FE" then
					DataPoints(5) = DataPoints(5) + objItem.MACAddress 
					NicCout =  NicCout + 1
					'end if	
				else
					DataPoints(5) = DataPoints(5) + "," + objItem.MACAddress
					
				end if	
			End If
		End If	
	
	Next

	if Verbos = true then
		WshShell.Popup  "objComsys.Description = " & DataPoints(5),6 'Debug
	end if
	
	
	get_SCCM_Import_Data 	= DataPoints
	Set objWMIService 	= Nothing
	Set colComsysPro	= Nothing
	Set colComsys 		= Nothing
	Set colBIOS 		= Nothing
	On Error GoTo 0
End Function
'=====================''=====================''====================='
'========================INIFile_reader.vbs V2.0==================================
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
		'wshshell.popup "My Key = " & myKey   & ",  My Str = " & strKey & ",   " '<<< Hard Stop
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
                                ReadIni = "*"
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
'==========================================================

Sub GetOptions()																			' GetOptions V 1.0
	Dim objArguments 		'Passed in commandline arguments
	Dim NumberOfArguments   'Number of Arguments
	Set objArguments = WScript.Arguments


'==========================================================
	If (objArguments.Count > 0) Then
		For NumberOfArguments = 0 To objArguments.Count - 1
			'WScript.Echo objArguments(NumberOfArguments)
			SetOptions objArguments(NumberOfArguments)
			
		Next
	End If
	
	If (objArguments.Count > 0) Then
		For NumberOfArguments = 0 To objArguments.Count - 1
			
		Next
	Else
		'WScript.Echo "For help type: \?"
	End If
		
End Sub 
Sub SetOptions(strOption)																	' Set and check Options V.1
	Dim strFlag
	Dim	strParameter
	Dim newArguments
	DIM bInvalidArgument
	DIM bDisplayHelp
	
	newArguments = Len(strOption)
	If (newArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Ucase(Left(strOption,2))
		
		
		Select Case strFlag
			Case "\?"
				bDisplayHelp = True
				
				
				
			Case Else
				bInvalidArgument = True
		End Select
	
	End If
End Sub
Sub DisplayHelp 																			'DisplayHelp V.1
	
 	WScript.Echo VbCrLf
 	WScript.Echo " \?	- Display help"
 	WScript.Echo VbCrLf
End Sub 
'==========================================================
'==========================================================	
Function CheckOS(strComputer)																'Check OS V 1.0
	Dim objWMIService												
	Dim colItems
	Dim objItem
	Dim ThisOS
	Dim ThisArchitecture
	
	' set wshShell= createObject("wscript.shell") 'for Pop up windows ' ****************  <<<<<<<< Delete later !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
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
Function PKiller(TheComputer,TheProcessToKill)												'Process Killer V 1
	Dim objWMIService
	Dim objProcess
	Dim colProcess
	
	Set objWMIService 	= GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & TheComputer & "\root\cimv2")
	Set colProcess 		= objWMIService.ExecQuery ("Select * from Win32_Process Where Name = " & TheProcessToKill )
	
	For Each objProcess in colProcess
		objProcess.Terminate()
	Next
	
	set objWMIService	= nothing
	Set colProcess 		= nothing
	PKiller = "Done"
End Function
'==========================================================
'==========================================================
Function WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue) 		'String wrght to reg v 1
	DIM objReg
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	TheReturn = objReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)	
	if TheReturn <> 0 then
		wshshell.popup "code return for " & strKeyPath & strValueName & strValue & " is " &  TheReturn'Debug
	end if	
end function
Function DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)	'Dword wright to reg V 1
	DIM objReg
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	TheReturn = objReg.SetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)	
	if TheReturn <> 0 then
		wshshell.popup "code return for " & strKeyPath & strValueName & dwValue & " is " &  TheReturn 'Debug
	end if	
end function
'==========================================================	
'==========================================================
wshshell.popup"Function Test is Done ",3 
'==========================================================
'>>>>>>>>> Network Functions Suiet V1.0' <<<<<<<<<
'==========================================================	
Function ReverseDNSLookup(sIPAddress)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
	DIM oShell,oFSO,sTemp,sTempFile,fFile,sResults,aNameTemp,aName
	
	
  If Not IsIP(sIPAddress) Then
    ReverseDNSLookup = "Failed-error.error"
    Exit Function
  End If

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1

  Set oShell = CreateObject("WScript.Shell")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
  sTempFile = sTemp & "\" & oFSO.GetTempName

  oShell.Run "%comspec% /c nslookup " & sIPAddress & ">" & sTempFile, 0, True

  Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsDefault)
  sResults = fFile.ReadAll
  fFile.Close
  oFSO.DeleteFile (sTempFile)

  If InStr(sResults, "Name:") Then
    aNameTemp = Split(sResults, "Name:")
    aName = Split(Trim(aNameTemp(1)), Chr(13))
  
    ReverseDNSLookup = aName(0)
  Else
    ReverseDNSLookup = "Failed-error.error"
  End If
  
  Set oShell = Nothing
  Set oFSO = Nothing
End Function

Function DNSLookup(sAlias)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
	
	dim oShell,oFSO,fFile,aIP,aIPTemp
	Dim sTemp,sTempFile,sResults
	
	
	
  If len(sAlias) = 0 Then
    DNSLookup = "Failed."
    Exit Function
  End If

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1

  Set oShell = CreateObject("WScript.Shell")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
  sTempFile = sTemp & "\" & oFSO.GetTempName

  oShell.Run "%comspec% /c nslookup " & sAlias & ">" & sTempFile, 0, True

  Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsDefault)
  sResults = fFile.ReadAll
  fFile.Close
  oFSO.DeleteFile (sTempFile)

  aIP = Split(sResults, "Address:")

  If UBound(aIP) < 2 Then
    DNSLookup = "Failed."
  Else
    aIPTemp = Split(aIP(2), Chr(13))
    DNSLookup = trim(aIPTemp(0))
  End If
  
  Set oShell = Nothing
  Set oFSO = Nothing
End Function

Function IsIP(sIPAddress)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
 Dim aOctets,sOctet
 
 
  aOctets = Split(sIPAddress,".")
  
  If IsArray(aOctets) Then
    If UBound(aOctets) = 3 Then
      For Each sOctet In aOctets
        On Error Resume Next
        sOctet = Trim(sOctet)
        sOctet = sOctet + 0
        On Error Goto 0
        
        If IsNumeric(sOctet) Then
          If sOctet < 0 Or sOctet > 256 Then
            IsIP = False
            Exit Function
          End If
        Else
          IsIP = False
          Exit Function
        End If
      Next
      IsIP = True
    Else
      IsIP = False
    End If
  Else
    IsIP = False
  End If
End Function

Function Get_IP_Gateway(strComputer) 
	'This function reutrns an arry with each actives NICs gateway ip4 address 
	Dim odjNetwork_interface_CArd,NIC_Set_Information,objWMIService
	reDim The_GPSC(0)
	Dim int_Couter_A 
	
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	set  NIC_Set_Information = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
	
	int_Couter_A = 0
   	Get_IP_Gateway = ""
    
			
            For Each odjNetwork_interface_CArd in NIC_Set_Information 
				
			    If odjNetwork_interface_CArd.IPEnabled = True Then
						
					reDim preserve The_GPSC(int_Couter_A)
					
					If verbos = true then  '<<<< Debug>>>>
						wshshell.popup"odjNetwork_interface_CArd.IPEnabled = " & odjNetwork_interface_CArd.IPEnabled ,3 
					end if 
					
					The_GPSC(int_Couter_A) = GetMultiString_FromArray(odjNetwork_interface_CArd.DefaultIPGateway, ",")
					
					If verbos = true then  '<<<< Debug>>>>	
						wshShell.popup"odjNetwork_interface_CArd.DefaultIPGateway = " & the_GPSC(int_Couter_A) & " int_Couter_A = " & int_Couter_A ,3 
					end if 
					
					int_Couter_A = int_Couter_A + 1
					
				end if
			
			next 
			
	Get_IP_Gateway 				= The_GPSC	
	Set objWMIService 			= nothing
	set  NIC_Set_Information	= nothing
	
	
	
	
end Function	

Function GetMultiString_FromArray( ArrayString, Seprator)
   Dim StrMultiArray
	

   If IsNull ( ArrayString ) Then

        StrMultiArray = ArrayString

    else

        StrMultiArray = Join( ArrayString, Seprator )

   end if

   GetMultiString_FromArray = StrMultiArray

End Function

Function Get_My_location_By_IP4_Address(strPath_of_file,strFile_name,The_computer)' Vertion 1 alpha 
	
	DIM strMyGateway
	DIm strGateway_Hostname_By_IP4
	Dim aryBigData
	Dim aryBigDataSlpit
	Dim strMylocation,strMylocationcode
	DIM int_Counter_A , intBCounter
	If strPath_of_file 	= "" then 
		strPath_of_file = Get_Then_current_Drive()
	end if 
		
	if strFile_name = "" then 
		strFile_name 		= "IPGPS_Data.csv" 'Data format (IP gateway,Location Name,Location number,check switch)
	End if
		
	if The_computer = "" then 
		The_computer 		= "."
	End if 	


	
	aryBigData 					= strRead_Data_FromCSV(strPath_of_file,strFile_name) 'getting data to be searched
	
		If Err.Number <> 0 Then
			if Err.Number = 53 then 
				'If verbos = true then  '<<<< Debug>>>>	
					wshshell.popup Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description,3 
				'END IF 	
			End if	
		Err.Clear
		WScript.Quit 1
		End If
	
	strMyGateway 				= Get_IP_Gateway(The_computer)
   

	if 	strMyGateway(0) = "" then 
		If verbos = true then  '<<<< Debug>>>>	
			wshshell.popup"No Acctive NIC  ",3 
		End if
	else
		
		for int_Counter_A = 0 to ubound(strMyGateway)
				
			If verbos = true then  '<<<< Debug>>>>	
				wshShell.popup	"this is the Gateway " & strMyGateway(int_Counter_A),3
			End if
			
			strGateway_Hostname_By_IP4 = ReverseDNSLookup(strMyGateway(int_Counter_A)) 
			
			If verbos = true then  '<<<< Debug>>>>
				wshShell.popup  strGateway_Hostname_By_IP4,3
			End if
			
		next 
	
	

	
		For intBCounter = 0 to ubound(strMyGateway)
			bolCheck_if_gateway_is_in_database = false
			
			'WshShell.popup ubound(aryBigData)'<<<< Debug>>>>
			
			
			for int_Counter_A = 0 to ubound(aryBigData)
				aryBigDataSlpit 		= Split(aryBigData(int_Counter_A),",")
				
				If verbos = true then  '<<<< Debug>>>>
					'WshShell.popup "checking " & aryBigDataSlpit(0) & " To " & strMyGateway(intBCounter),1 'Deug 
				End if
				
				
				IF aryBigDataSlpit(0)	= strMyGateway(intBCounter) THEN
					
					strMylocation 		= aryBigDataSlpit(1)
					strMylocationcode	= aryBigDataSlpit(2)
					Get_My_location_By_IP4_Address = Array(strMylocation,strMylocationcode)
					
					bolCheck_if_gateway_is_in_database 			= true
					If verbos = true then  '<<<< Debug>>>>
						WshShell.POPUP "Location is " & strMylocation,3
						
					End if
					EXIT For 'stop searching
				END If
				
			Next
		IF bolCheck_if_gateway_is_in_database = false then 
				'WshShell.POPUP "We did not find this address , do you with to add it (y/n)"
		
		
		
		
		
		
		
		end if
		
		Next	

				'WshShell.POPUP strMylocation 

	end if
	

End Function 
Function Get_Then_current_Drive()
	
	Dim strPath,strDrive,objFSO,objShell
	
	Set objFSO 		= CreateObject("Scripting.FileSystemObject")
	Set objShell 	= CreateObject("Wscript.Shell")
	strPath 		= objShell.CurrentDirectory
	strDrive 		= objFSO.GetDriveName(strPath)
	Get_Then_current_Drive =  strPath

End function 
Function strRead_Data_FromCSV(strFilePath,strDataFile)
	
	
	If verbos = true then 
		wshShell.popup"Data File PATH " & strFilePath & "\" & strDataFile ,3
	end if 
	
	Dim objFSO
	Dim objTextFile
	Dim strTheDataPath
	Dim strNextLine
	Dim aryWorkingDataSet()
	Dim TheHoastName
	Dim intACounter: intACounter = 0
	Const ForReading = 1
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	if  strDataFile = "" then 'strFilePath = "" or
			strRead_Data_FromCSV = ""
		 else
				
		Set objTextFile = objFSO.OpenTextFile(strFilePath & "\" & strDataFile , ForReading)
	
		Do Until objTextFile.AtEndOfStream
			ReDim preserve aryWorkingDataSet(intACounter) 
			aryWorkingDataSet(intACounter)				=  Replace(objTextFile.Readline,Chr(34),"",1,2)
			intACounter 								= intACounter + 1
		Loop
		strRead_Data_FromCSV = aryWorkingDataSet
	end  if 
	
End Function
Function Get_site_name_by_gateway_hostname()
	dim The_IP_adress
	Dim Hostname
	Dim Get_Splited_NAme
	DIM Site
 
	The_IP_adress 		= Get_IP_Gateway(".")
	Hostname			= ReverseDNSLookup(The_IP_adress(0))
 
	Get_Splited_NAme = Split(Hostname,".")
	Site = Split(Get_Splited_NAme(0),"-")
	Get_site_name_by_gateway_hostname = Site(1)
end Function 
'>>>>>>>>>                            <<<<<<<<<

