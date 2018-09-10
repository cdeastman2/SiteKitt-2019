Imports System
Imports System.Management
Imports System.Windows.Forms
Imports System.Data.OleDb
Module SCCM_Data
    Private objSBO As ManagementObjectSearcher 'Computer BIOS
    REM  Private colBIOS As ManagementObjectSearcher 'Computer BIOS
    Public Function BIO_DATA()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_BIOS")

        Dim objMgmt
        Dim strSerialNumber As String = ""
        For Each objMgmt In objSBO.Get
            strSerialNumber = Trim(objMgmt("SerialNumber").ToString())
        Next

        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Dim strManufacturer = ""
        Dim strModel = ""
        Dim strDNSHostName = ""
        Dim strDomian = ""
        Dim strFirst_Data_In = Now()

        Dim strUUID As String = ""

        For Each objMgmt In objSBO.Get
            strManufacturer = objMgmt("Manufacturer").ToString()
            strModel = objMgmt("Model").ToString()
            strDNSHostName = objMgmt("DNSHostName").ToString()
            strDomian = objMgmt("Domain").ToString()
        Next
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystemProduct")

        For Each objMgmt In objSBO.Get
            strUUID = objMgmt("UUID").ToString()
        Next
        REM 



        Dim All_My_MACs(1)
        Dim Acounter As Integer = 0
        Dim NIC_MACS As String = ""
        Dim MAC_Addresss As String = ""
        Try
            Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_NetworkAdapter")
            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("NetConnectionID") IsNot Nothing Then

                    If queryObj("Manufacturer").ToString <> "Microsoft" Then
                        If Acounter = 0 Then
                            All_My_MACs(Acounter) = queryObj("MACAddress")
                        Else
                            All_My_MACs(Acounter) = queryObj("MACAddress")
                        End If
                        Acounter = Acounter + 1
                        If Acounter > 1 Then
                            ReDim Preserve All_My_MACs(Acounter)
                        End If
                    End If
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show("An error occurred while querying for WMI data: " & err.Message)
        End Try

        For Each NIC_MACS In All_My_MACs
            If MAC_Addresss = "" Then
                MAC_Addresss = NIC_MACS
            Else
                MAC_Addresss = MAC_Addresss & "," & NIC_MACS
            End If
        Next

        BIO_DATA = {strSerialNumber, strUUID, strManufacturer, strModel, strDNSHostName, strDomian, strFirst_Data_In, MAC_Addresss}

    End Function

End Module
