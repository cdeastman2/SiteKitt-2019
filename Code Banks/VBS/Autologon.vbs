'=========================================================
' LANG 					: VBscript 
' NAME 					: Autologon.vbs	
' AUTHOR				: Clifford Dinel Eastman ii	
' Email					: Clifford.eastmanii@Fresnounified.org
' VERSION 				: .5
' DATE 					: 09.28.2015
' Description			: setup computer for autologin user local account
' COMMENTS 				: Have fun, use at your own risk
''=========================================================
'==========================================================
' Settings
	set wshShell= createObject("wscript.shell") 'for Pop up windows 
	WshShell.Popup "Starting Script Autologon.vbs ",2 ' Script launch verifactioin 
	
	call AutoLogOn
	
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
	