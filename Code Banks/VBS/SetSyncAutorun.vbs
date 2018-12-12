	set wshShell= createObject("wscript.shell")	
	wshshell.popup "starting ",1
	strComputer ="."
	const HKEY_LOCAL_MACHINE = &H80000002
	Call ActivateAuotrunsync()
	wshshell.popup "Done ",5

sub ActivateAuotrunsync()
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
	strValueName = Chr(34) & "SMART Sync Helper Service" & Chr(34) 'Not user throws exseption 
	strValue = "C:\\Program Files\\SMART Technologies\\Sync Student\\SyncClient.exe\"
	Return = objReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
		wshshell.popup "code return = " & Return,10 'Debug
End Sub