Option Explicit
'==========================================================
' LANG 					: VBScript
' NAME 					: VBstarter V 2.0
' AUTHOR				: Clifford Dinel Eastman ii	
' Mail					: CDEASTM4FUSD@gmail.com
' VERSION 				:
' DATE 					:	
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
	Dim strComputer 'The Computer being acted on 
		
	Dim TheProsess	'The current  process being exacuted 
	Dim Error_Info	'Error reporting 
	Dim MyOS
	DIM verbos		'Display messages
	Dim TheReturn
	
	
'==========================================================
' Program's Variable Declarations and Constants
'==========================================================
	Dim Collection
	Dim Groupping 
	Rem DIM Site
	DIM system_number
	Dim SerialNumber
	Dim Input_keyed
	dim The_Data
	Dim A_looper
	






'==========================================================	
' Program's Default settings
'==========================================================	
'If verbos = true then 	
	'	
	'End if












'==========================================================
' Main Body
'==========================================================
	set wshShell= Wscript.CreateObject("WScript.Shell")
	If verbos = true then 	
		wshshell.popup"Starting",3 
	End if
	
	'On Error resume next    	
	'GetOptions 'Get Options from user

	REM ********************************
	Collection = Get_Collection_file_Name()
	REM ********************************
	If verbos = true then 	
		wshshell.popup Collection,5
	end if 
	
	REM ********************************
	The_Data = Bulk_Serialnumber_Input()
	REM ********************************
	If verbos = true then 	
		for A_looper = 0 to Ubound(The_Data)
			wshshell.popup The_Data(A_looper)
		next
	END IF 
	
	TheReturn = Wright_BULK_Data(The_Data,Collection)
	
	
	
	
	
	
	
	If verbos = true then 	
	 wshshell.popup"Done  ",3
	End if
''==========================================================
'Program's Funktions and Subroutines 
'==========================================================

Function Get_Collection_file_Name()' Vertion 1.0 Aplha
	DIM txt_Collection
	Dim Data_Check
	Dim My_Anser_is
	 
	
	Data_Check = False
	Do While Data_check = False	
		txt_Collection 					= Trim(Inputbox("Enter Collection # ","Input Required"))
		If 	txt_Collection 				= "" Then
			txt_Collection 				= Trim(Inputbox(" Data enter error pleace re-enter Collection file name or Q to quit  ","Input Required"))
			If UCase (txt_Collection)	= "Q" Then 
				
				WScript.Quit
				
			End If
		End If 
		
		
		My_Anser_is = UCase(InputBox("Is This Correct File name Y or N " & txt_Collection & ".csv" ,"Input Required"))
		If My_Anser_is = "Y" Then
			Data_Check = True
			Get_Collection_file_Name 	= Trim(txt_Collection) & ".csv"
		End If
	Loop 	
	
	
	
	
	
	
End Function 	


Function Scan_SerialNumber()
	Dim txt_SerialNumber
	Dim Data_check
	txt_SerialNumber =""
	
	
	
	txt_SerialNumber = Inputbox("Scan Computer Serial Number Barcode","Scan Barcode")
	
	Scan_SerialNumber = txt_SerialNumber
	If verbos = true then
		wshshell.popup Scan_SerialNumber,1
	end if 	
	
End Function 

Function Bulk_Serialnumber_Input()
	
	
	Dim A_SerialNumber
	reDim Data_List(0)
	Dim looper
	Dim Check_for_Dup_Loop
	Dim DupCheck
	Dim A_Counter

	Data_List(0) 	= "" 
	looper 			= true
	A_Counter 		= 0
	A_SerialNumber	= ""
	DupCheck 		= false
	
	
	Do While Looper = true
		
		If verbos = true then
			wshshell.popup A_Counter,5
		end if
		
		A_SerialNumber = Scan_SerialNumber()
		
		if A_SerialNumber = "" then
			looper = False
		
		ELSE 
			
			For Check_for_Dup_Loop = 0 to Ubound(Data_List)
				if 	A_SerialNumber = Data_List(Check_for_Dup_Loop) then
					DupCheck = true
				end if	
			next

			
			if DupCheck = True then 
				wshshell.popup "Dupilcate data entery re-enter ",5
				DupCheck = False
				
			else 
					if A_Counter >  Ubound(Data_List) then 
						redim preserve Data_List(A_Counter)
						Data_List(A_Counter) = A_SerialNumber
						A_counter = A_Counter + 1
					else
						Data_List(A_Counter) = A_SerialNumber
						A_counter = A_Counter + 1
					end if		

			end if 
		end if	
		
	loop

	Bulk_Serialnumber_Input = Data_List
	
End Function

Function Wright_BULK_Data(Data2wright,File2write2)
	Dim objFSO 'file system object
	DIM objTextStream
	Dim Looper
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	If (objFSO.FileExists(File2write2)) Then
		Set objTextStream = objFSO.OpenTextFile(File2write2,8, True)
	ELSE
		Set objTextStream = objFSO.OpenTextFile(File2write2,8, True)
		objTextStream.WriteLine("SerialNumber" & "," )
	END IF 
	
	
	For looper = 0 to UBound(Data2wright)
		objTextStream.WriteLine(Data2wright(looper) & "," )
		If verbos = true then
			wshshell.popup Data2wright(looper),5
		END IF	
		
	next
    

	

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