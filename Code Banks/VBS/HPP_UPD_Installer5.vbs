Option Explicit
'==========================================================
' LANG 					: VBScript
' NAME 					: HPP_UPD_Installer
' AUTHOR				: Clifford Dinel Eastman ii	
' Mail					: CDEASTM4FUSD@gmail.com
' VERSION 				: 5
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
	Dim TheDriver_path,ThePrinter_port,ThePrinter_Name,Remove_all_Printers
'==========================================================	
' Program's Default settings
'==========================================================	
	
	verbos				= False
	TheDriver_path 		= ""
	ThePrinter_port		= ""
	ThePrinter_Name 	= ""
	Remove_all_Printers	= false
	strComputer 		= "."

'==========================================================
' Main Body
'==========================================================
	set wshShell= Wscript.CreateObject("WScript.Shell")
	
	If verbos = true then 	
		wshshell.popup"Starting Script ",3 
	End if
	
	'On Error resume next    



	Dim intAcouter
	Dim Printers_Install_List
	Dim MyiniFilePath
	Dim Printer_Data 'Printer setup data
	
'==========================================================	
	GetOptions 'Get Options from user
'==========================================================	
	
	
	If verbos = true then 	
		wshShell.Popup MyiniFilePath,3
	End if
	Printers_Install_List = Get_printer_data_From_ini(MyiniFilePath)
	
	If verbos = true then
		wshShell.popup UBound(Printers_Install_List) ' debug 
	End if
	
	MyOS = CheckOS(strComputer)
	
	
	For intAcouter = 1 to UBound(Printers_Install_List)
		Printer_Data = Split(Printers_Install_List(intAcouter),",")
		If Verbos = true then 
			wshShell.POPUP "Installing " & Printer_Data(0),3  ' debug 
		end if
		TheReturn = Install_Universal_HP_Network_Printer("",Printer_Data(1), Printer_Data(0) )
	NEXT






	


	
'==========================================================			
	If verbos = true then 	
	 wshshell.popup"Script is Done ",3
	End if
''==========================================================
'Program's Funktions and Subroutines 
'==========================================================
	
Function Install_Universal_HP_Network_Printer(Driver_path,Printer_Port,Printer_Name)
				
		If Driver_path = "" Then
			if MyOS(1) = "64-bit" then
				Driver_path = wshShell.CurrentDirectory & "\HP\64BIT\upd-pcl6-x64-6.3.0.21178\Install.exe"
				'Driver_path = """\\rhs--rnd\Master Profiles\Drivers\Printer\Universal\HP Universal Print Driver\upd-pcl6-x64-6.3.0.21178\Install.exe"""
				'wshShell.Run Driver_path & " /h /q /sm" & Trim(Printer_Port) & " /n" & Trim(Printer_Name),1,true
			else
				Driver_path = wshShell.CurrentDirectory & "\HP\32Bit\upd-pcl6-x32-6.3.0.21178\Install.exe"
				'Driver_path = """\\rhs--rnd\Master Profiles\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x32-5.5.0.12834\Install.exe"""
				'wshShell.Run Driver_path & " /h /q /sm" & Trim(Printer_Port) & " /n" & Trim(Printer_Name),1,true
			end if
		end if 
	
	'wshShell.Run Driver_path & " /h /q /smDHS-LMC-PRN01 /n""DHS-LMC-Printer #01 """, 1, true
	 wshShell.Run Driver_path & " /q /h /npf /sm" & Trim(Printer_Port) & " /n" & Trim(Printer_Name),1,true

End function

Sub Removing_all_Printers(ThisComputer)
	dim	objShell
	set objShell = Wscript.CreateObject("WScript.Shell")
	Set objWMIService = GetObject("winmgmts:\\" & ThisComputer & "\root\cimv2")
	Set colInstalledPrinters =  objWMIService.ExecQuery("Select * from Win32_Printer")

	For Each objPrinter in colInstalledPrinters
		objPrinter.Delete_
	Next
	Set objShell = nothing 
End Sub

'==========================================================
'System Funktions ,Subroutines and Dependencies
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
	If (newArguments < 3) Then
		bInvalidArgument = True
	Else
		strFlag = Ucase(Left(strOption,3))
		
		
		Select Case strFlag
			Case "/?"
				bDisplayHelp = True
			
			Case "/VB"	
				verbos= true	
				
						
			Case "/IN"	
				MyiniFilePath = Trim(right(strOption,newArguments - 3))
			
			Case "/RP"	
				Remove_all_Printers = true
			
						
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
		'wshShell.POPUP "code return for " & strKeyPath & strValueName & strValue & " is " &  TheReturn'Debug
	end if	
end function
Function DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)	'Dword wright to reg V 1
	DIM objReg
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	TheReturn = objReg.SetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)	
	if TheReturn <> 0 then
		'wshShell.POPUP "code return for " & strKeyPath & strValueName & dwValue & " is " &  TheReturn 'Debug
	end if	
end function
'==========================================================	
'==========================================================
Function Get_printer_data_From_ini(FilePath) 'Get_printer_data_From_ini 1.0 beta 
	Dim Printer_count 	'Number of printer in group
	Dim Printer_List() 	'Array with all printers data 
	Dim Section, aKey
	Section 	= "Printers"
	AKey 		= "Number"
	
	'wshShell.POPUP FilePath
	Printer_count = ReadIni(FilePath,Section,aKey)
	
	if Printer_count < 1 then
		wshshell.popup"No printer data ",3 
		 Wscript.Quit 1
	Else
		redim Printer_List(Printer_count)
	end if	
		
		'wshshell.popup Printer_count
	For intAcouter = 1 to Printer_count
		aKey = "printer" & intAcouter
		Printer_List(intAcouter) = ReadIni(FilePath,Section,aKey)
		'wshShell.popup Printer_List(intAcouter)
	next	
	 
	
	Get_printer_data_From_ini = Printer_List
	
End Function 
'[Section 003]
Function ReadIni(myFilePath, mySection, myKey ) 'Readini ver 1.0
	
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
    If objFSO.FileExists(strFilePath) Then
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
Function UpDataINI(myFilePath,INILineToUpdate,NewValue)'Updatini ver 1.0
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
'[End Section 003]
'==========================================================
