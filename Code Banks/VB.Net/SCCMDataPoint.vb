Imports System
Imports System.Management
REM Imports System.Windows.Forms
Public Class SCCMDataPoint
    Public ModualName As String = "CurrentSystems"
    Private objSBO As ManagementObjectSearcher 'Computer BIOS

    Public Function SerialNumber()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_BIOS")
        Dim objMgmt
        Dim strSerialNumber As String = Nothing


        For Each objMgmt In objSBO.Get
            strSerialNumber = Trim(objMgmt("SerialNumber").ToString())
        Next
        SerialNumber = strSerialNumber
        objSBO = Nothing

    End Function

    Public Function UUID()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystemProduct")
        Dim objMgmt
        Dim txtUUID = Nothing

        Try
            For Each objMgmt In objSBO.Get
                txtUUID = Trim(objMgmt("UUID").ToString())
            Next


        Catch ex As Exception
            txtUUID = Nothing
        End Try


        UUID = txtUUID
        objSBO = Nothing


    End Function

    Public Function Manufacturer()
        Manufacturer = Nothing
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Try
            For Each objMgmt In objSBO.Get
                Manufacturer = objMgmt("Manufacturer").ToString()
            Next
        Catch
        End Try


        objSBO = Nothing

    End Function

    Public Function Model()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Model = Nothing
        Try
            For Each objMgmt In objSBO.Get
                Model = objMgmt("Model").ToString()
            Next
        Catch ex As Exception

        End Try

    End Function

    Public Function DNSHostName()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        DNSHostName = Nothing
        Try
            For Each objMgmt In objSBO.Get
                DNSHostName = objMgmt("DNSHostName").ToString()
            Next
        Catch ex As Exception

        End Try

    End Function

    Public Function Domian()
        objSBO = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        Domian = Nothing
        Try
            For Each objMgmt In objSBO.Get
                Domian = objMgmt("Domain").ToString()
            Next
        Catch ex As Exception

        End Try

    End Function

    Public Function MACAddress()


        Dim All_My_MACs(1)
        Dim Acounter As Integer = 0
        Dim NIC_MACS As String = ""
        Dim MAC_Addresss As String = ""
        Try
            Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_NetworkAdapter")
            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("NetConnectionID") IsNot Nothing Then

                    If queryObj("Manufacturer").ToString IsNot "Microsoft" Then

                        All_My_MACs(Acounter) = queryObj("MACAddress")
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
                REM MsgBox(NIC_MACS)'debug 
                If NIC_MACS <> " " Then
                    MAC_Addresss = MAC_Addresss & "," & NIC_MACS
                End If

            End If
        Next
        MACAddress = All_My_MACs


    End Function


End Class
