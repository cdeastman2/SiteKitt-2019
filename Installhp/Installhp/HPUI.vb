Module HPUI
    Dim strPrinterDriver, strPrinterName, strDriverLocation, strPrinterINFlocation, strIPPort, Printertype
    Dim strComputer



   
    REM Printer Infomation
	strPrinterName 			="Addicott Office Ricoh"
	strIPPort				="172.26.58.54"
	Printertype				= 0


MyOS = CheckOS(strComputer)
    'wshshell.popup myos(1) 'debug

	if MyOS(1) = "64-bit" then
		If Printertype = 1 then
			strDriverLocation		="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x64-5.5.0.12834"
			strPrinterINFlocation 	="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x64-5.5.0.12834\hpbuio45l.inf"
			strPrinterDriver 		="HP Universal Printing PCL 5"
		else
			strDriverLocation		="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\Ricoh\64bit"
			strPrinterINFlocation 	="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\Ricoh\64bit\oemsetup.inf"
			strPrinterDriver 		="PCL6 Driver for Universal Print"
		end if

	else
		if PrinterType = 1 then
			strDriverLocation		="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x32-5.5.0.12834"
			strPrinterINFlocation 	="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\HP Universal Print Driver\pcl5-x32-5.5.0.12834\hpbuio45l.inf"
			strPrinterDriver 		="HP Universal Printing PCL 5"
		else

			strDriverLocation		="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\Ricoh\32bit"
			strPrinterINFlocation 	="\\rhs-rnd\Master Profile\Drivers\Printer\Universal\Ricoh\32bit\oemsetup.inf"
			strPrinterDriver 		="PCL6 Driver for Universal Print"
		end if

	end if 

return = Prnport(strIPPort)   

RETURN = Prndrvr(strPrinterDriver,strPrinterINFlocation,strDriverLocation)

return = Prncnfg(strPrinterName,strPrinterDriver,strIPPort)

    REM WSHSHELL.POPUP "DONe" 




    Function Prncnfg(PrinterName, PrinterDriver, strIPPort)
        objShell = CreateObject("WScript.Shell")
        Dim Setting
        SETTING = "-a -p " & chr(34) & PrinterName & Chr(34) & " -m " & chr(34) & PrinterDriver & chr(34) & " -r " & chr(34) & "IP_" & strIPPort & chr(34)
        REM WSHSHELL.POPUP SETTING 
	return = objShell.Run("C:\Windows\System32\Cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnmngr.vbs " & SETTING,2,true)

    End Function

    Function Prndrvr(PrinterDriver, PrinterINFlocation, DriverLocation)
        objShell = CreateObject("WScript.Shell")
        Dim Setting

        SETTING = "-a -m " & chr(34) & PrinterDriver & chr(34) & " -i " & chr(34) & PrinterINFlocation & chr(34)
        REM WSHSHELL.POPUP SETTING 

	return = objShell.Run("C:\Windows\System32\Cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\Prndrvr.vbs " & SETTING,2,true)

    End Function

    Function Prnport(PortIP)
        Dim Setting
        Setting = "-a -r IP_" & PortIP & " -h " & PortIP

        REM WSHSHELL.POPUP SETTING 
        objShell = CreateObject("WScript.Shell")
	return = objShell.Run("C:\Windows\System32\Cscript.exe C:\Windows\System32\Printing_Admin_Scripts\en-US\Prnport.vbs " & SETTING,2,true)

    End Function

    Function CheckOS(strComputer)
        objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
        colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", , 48)
        For Each objItem In colItems

            ThisOS = objItem.Name
            ThisArchitecture = objItem.OSArchitecture
            CheckOS = Array(ThisOS, ThisArchitecture)
        Next

    End Function

End Module
