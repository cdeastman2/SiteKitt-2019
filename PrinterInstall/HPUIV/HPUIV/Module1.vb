Imports System.Management

Module Module1


    Public Sub Get_printer_list()
        ' Use the ObjectQuery to get the list of configured printers
        Dim oquery As System.Management.ObjectQuery = New System.Management.ObjectQuery("SELECT * FROM Win32_Printer")

        Dim mosearcher As System.Management.ManagementObjectSearcher = New System.Management.ManagementObjectSearcher(oquery)

        Dim moc As System.Management.ManagementObjectCollection = mosearcher.Get()

        For Each mo As ManagementObject In moc
            Dim pdc As System.Management.PropertyDataCollection = mo.Properties
            For Each pd As System.Management.PropertyData In pdc

                If CBool(mo("Network")) Then
                    Form1.ListBox1.Items.Add(mo(pd.Name))
                End If
            Next pd
        Next mo

    End Sub



End Module
