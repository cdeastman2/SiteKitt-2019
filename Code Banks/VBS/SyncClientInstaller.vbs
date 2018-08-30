'========================11================================
' LANG 					: VBScript
' NAME 					: SCSetup.vbs
' AUTHOR				: Clifford Dinel Eastman ii	
' Mail					: Clifford.Eastmanii@fresnounified.org
' VERSION 				: Alpah
' DATE 					: 12/2/2015
' Description			: install and setup Smartware Sync Client 2011 
'						: 
' COMMENTS 				: 
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
' Settings
	Dim objArguments, NumberOfArguments,strPath
	Dim theGo,strComputer,MyOS,strProcessKill
	Dim strTeacherComputer
	
	set wshShell= createObject("wscript.shell")	
	Set objArguments = WScript.Arguments
	wshShell.popup"Starting Script ",3 
	
	
	If (objArguments.Count > 0) Then
		For NumberOfArguments = 0 To objArguments.Count - 1
			strTeacherComputer = CStr(objArguments(NumberOfArguments))
		Next
	else
		'wshShell.popup "no Parameters ",10
		strTeacherComputer = "         "
		wshShell.popup strTeacherCompute,3
	End If
	
	
'==========================================================
		strComputer ="."
		const HKEY_LOCAL_MACHINE = &H80000002
		MyOS = CheckOS(strComputer)
		strProcessKill = "'dax64.exe'"

		
		
		
'==========================================================
' Main Body
'==========================================================
	On error goto next '******* Unrem for RCcode *******
		
		return = PKiller(strComputer,strProcessKill)
		
		
	if  Check4file() = 1 then
		'wshShell.popup "not installing client goint to setup",5 'debug
	else
		'wshShell.popup "installing Client", 5
		strPath = "\\sccm-data\Software\Regions\Mclane\Software_Source\SyncClient\SmartSync2011Student\SMARTSyncStudent.msi"
		wshShell.Run strPath & " /qb",1,true
		
	end if
	
	
	if MyOS(1) = "64-bit" then
			'wshShell.popup"64-bit",2 'For Debug
			strKeyPath = "SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student\"
			call setSyncSeting(strKeyPath)
	else
			'wshShell.popup"32-bit",2 'For Debug
			strKeyPath = "SOFTWARE\SMART Technologies\SMART Sync Student\"
			call setSyncSeting(strKeyPath)
	end if
		
	
	strValueName = "ConnectIP"
	strValue = strTeacherComputer 
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	
	strValueName = "Visible"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	
	wshShell.run (chr(34) & "C:\Program Files (x86)\SMART Technologies\Sync Student\SyncClient.exe"& chr(34)),1,True

	wshshell.popup"Script is Done ",3
'==========================================================
'Fuctions and Subroutines 
'==========================================================

Function PKiller(strComputer,strProcessToKill)
	Dim objWMIService, objProcess, colProcess
	Set objWMIService 	= GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcess 		= objWMIService.ExecQuery ("Select * from Win32_Process Where Name = " & strProcessToKill )
	
	For Each objProcess in colProcess
		objProcess.Terminate()
	Next
	
	set objWMIService	= nothing
	Set colProcess 		= nothing
	PKiller = "Done"
End Function

Sub setSyncSeting(strKeyPath)
	
	
	strValueName = "ConnectIP"
	strValue = strTeacherComputer
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	
	strValueName = "RedrawHooks"
	strValue = &H000003e8
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "LanguageID"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "MirrorDriver"
	strValue = &H000003e8
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "MulticastTTL"
	strValue = &H00000020
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "UnicastNoDelay"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "StudentIDMode"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "PasswordHash"
	strValue = ""
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "NICListLength"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "UseNamingServer"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "NTGroupListLength"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "Student Path"
	strValue = "C:\\Program Files\\SMART Technologies\\Sync Student\\"
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "ActiveDirStudentIdField"
	strValue = "cn"
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "CtrlAltDelSettings"
	strValue = &H00000002
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "EnableFileTransfer"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	'strValueName = "Visible"
	'strValue = &H0000001
	'Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "StudentID"
	strValue = "AnonymousID"
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "Build Version"
	strValue = "10.0.574.0"
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "ConnectTeacherID"
	strValue = ""
	Call WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)


	strValueName = "EnableHelp"
	sstrValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "DisplayExit"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "StoreFilesToMyDocs"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "AutoStart"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "CustomSharedFolder"
	strValue = ""
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "SecurityUsed"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "ConnectionUsed"
	strValue = &H00000003
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "EnableNICDefaultOrder"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "EnableQuestions"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "NamingServerLoc"
	sstrValue = ""
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	
	strValueName = "EnableChat"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "NamingServerPassedTest"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "HWGPUCapture"
	strValue = &H000003e8
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "GPUBoolEnableTimings"
	strValue = &H0000000
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "GPUBoolMergeRegions"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "GPUTileCount"
	strValue = &H0000001
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "GPUScaleFactor"
	strValue = &H00000003
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)

	strValueName = "GPUThresholdPixelCorrection"
	strValue = &H000003e8
	Call DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	

End Sub

Function Check4file()
	Dim go
	set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists("C:\Program Files (x86)\SMART Technologies\Sync Student\SyncClient.exe") Then
	
		'wshShell.popup "The file does exist",10 'Debug
		go = 1
	Else
		'wshShell.popup "The file does not exist.",10 'Debug
		go = 0
	End If
	
	set objFSO = nothing
	Check4file = go
end function

Function WrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	Return = objReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)	
	if return <> 0 then
		wshshell.popup "code return for " & strKeyPath & strValueName & strValue & " is " &  Return'Debug
	end if	
end function

Function DwordWrightToReg(strComputer,HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	Return = objReg.SetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)	
	if return <> 0 then
		wshshell.popup "code return for " & strKeyPath & strValueName & dwValue & " is " &  Return 'Debug
	end if	
end function

	
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