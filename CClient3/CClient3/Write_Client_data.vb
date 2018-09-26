Module Write_Client_data
    Private access As New DBControle

    
    Public Sub AddSystem(MyPrimare_Collection)
        REM MsgBox(MyPrimare_Collection) 'Debug 
        Dim MyCurrentSystem = New SCCMDataPoint
        Dim MySerialNumber As String
        Dim MyUUID As String

        Dim MComputerName As String
        Dim MyDomain As String


        Dim Mymanufature As String
        Dim MyModel As String
        Dim MyMACs As String = ""
        REM  Dim MyPrimare_Collection


        Dim Lastupdated = Now()

        REM Debug *************************************************
        MySerialNumber = MyCurrentSystem.SerialNumber
        MyUUID = MyCurrentSystem.UUID

        MComputerName = MyCurrentSystem.DNSHostName
        MyDomain = MyCurrentSystem.Domian

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
        REM  MsgBox(MySerialNumber)

        access.AddParam("@MyUUID", MyUUID)
        REM  MsgBox(MyUUID)

        access.AddParam("@MYComputerName", MComputerName)
        REM MsgBox(MComputerName)

        access.AddParam("@MYdomain", MyDomain)
        REM MsgBox(MyDomain)



        access.AddParam("@MainSCCMCollection", MyPrimare_Collection)
        REM  MsgBox(MyPrimare_Collection)


        access.AddParam("@MyNamufaturer", Mymanufature)
        REM  MsgBox(Mymanufature)

        access.AddParam("@MyModel", MyModel)
        REM  MsgBox(MyModel)

        access.AddParam("@MyMACS", MyMACs)
        REM MsgBox(MyMACs)


        access.AddParam("@Site_Computername", MyMACs)
        REM access.AddParam("@MyDate", Now())
        REM  MsgBox(Now())


        access.ExecQuery("INSERT INTO CPUINVO (SerialNumber,UUID,ComputerName,theDomain,Main_SCCM_Group,Manufacturer,Model,MACS)" & _
                         "Values (@SerialNumber,@MyUUID,@MYComputerName,@Mydomain,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")

        If Not String.IsNullOrEmpty(access.Exception) Then
            MsgBox("Shite ! = " & access.Exception)

        End If








    End Sub


    Private Sub UpdateRecord()
        
        REM   Access.ExecQuery("UPDATE members " & _
        REM                 "SET username=@user,[Password]=@pass,email=@email,website=@website,Active=@active " & _
        REM                "WHERE id=@uid")

    End Sub



    Public Sub UpDateSystem()


        access.ExecQuery("UPDATE members " & _
                        "SET username=@user,[Password]=@pass,email=@email,website=@website,Active=@active " & _
                        "WHERE id=@uid")





    End Sub

End Module
