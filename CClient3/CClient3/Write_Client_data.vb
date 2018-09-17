Module Write_Client_data
    Private access As New DBControle

    Public Sub AddSystem(MyPrimare_Collection)
        Dim MyCurrentSystem = New SCCMDataPoint


        Dim MySerialNumber As String
        Dim MyUUID As String
        Dim MComputerName As String
        Dim MyDomain As String = ""
        Dim Mymanufature As String
        Dim MyModel As String
        Dim MyMACs As String = ""


        REM  Dim MyPrimare_Collection
        Dim Lastupdated = Now()

        REM Debug *************************************************
        MySerialNumber = MyCurrentSystem.SerialNumber
        MyUUID = MyCurrentSystem.UUID
        MComputerName = MyCurrentSystem.DNSHostName
        Mymanufature = MyCurrentSystem.Manufacturer
        MyModel = MyCurrentSystem.Model


        For Each R In MyCurrentSystem.MACAddress
            If MyMACs = "" Then
                MyMACs = R
            Else
                If R <> "" Then
                    MyMACs = MyMACs & "," & R
                End If

                End If
        Next


        access.AddParam("@SerialNumber", MySerialNumber)
        REM  MsgBox(MySerialNumber)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyUUID", MyUUID)
        REM  MsgBox(MyUUID)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MYComputerName", MComputerName)
        REM MsgBox(MComputerName)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MainSCCMCollection", MyPrimare_Collection)
        REM  MsgBox(MyPrimare_Collection)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyDomain", MyDomain)
        MsgBox(MyDomain) 'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyNamufaturer", Mymanufature)
        REM  MsgBox(Mymanufature)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyModel", MyModel)
        REM  MsgBox(MyModel)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyMACS", MyMACs)
        REM MsgBox(MyMACs)'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<

        access.AddParam("@MyDate", Now())
        REM  MsgBox(Now())'Debug <<<<<<<<<<<<<<<<<<<<<<<<<<


        access.ExecQuery("INSERT INTO CPUINVO (SerialNumber,UUID,ComputerName,Main_SCCM_Group,Manufacturer,Model,MACS,domain)" & _
                         "Values (@SerialNumber,@MyUUID,@MYComputerName,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")

        If Not String.IsNullOrEmpty(access.Exception) Then
            MsgBox("Shite ! = " & access.Exception)

        End If








    End Sub



End Module
