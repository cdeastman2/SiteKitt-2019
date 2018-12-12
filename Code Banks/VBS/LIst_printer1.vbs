


strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

For Each objPrinter in colPrinters
    If Not objPrinter.Attributes And 64 Then 
        strText = strText & objPrinter.Name & vbCrLf
		 WScript.Echo strText
    End If
Next
