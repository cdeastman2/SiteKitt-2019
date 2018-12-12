Module Write_Client_data
    Private access As New DBControle

    
    Public Sub AddSystem(MyPrimare_Collection)


        MsgBox(MyPrimare_Collection) 'Debug 


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
        REM seem enephishant <!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! 
        For Each This_MAC_Address In MyCurrentSystem.MACAddress
            If MyMACs = "" Then
                MyMACs = This_MAC_Address
            Else
                If This_MAC_Address <> "" Then
                    MyMACs = MyMACs & "," & This_MAC_Address
                End If

            End If
        Next
        REM ******************************************************************

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

        REM  access.AddParam("@Site_Computername", MyMACs)
        REM access.AddParam("@MyDate", Now())
        REM  MsgBox(Now())


        access.ExecQuery("INSERT INTO CPUINVO (SerialNumber,UUID,ComputerName,theDomain,Main_SCCM_Group,Manufacturer,Model,MACS)" & _
                     "Values (@SerialNumber,@MyUUID,@MYComputerName,@Mydomain,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")

        If Not String.IsNullOrEmpty(access.Exception) Then
            MsgBox(access.Exception.ToString)

            If Left(access.Exception.ToString, 11) = "-2147467259" Then

                REM  If access.Exception = "-2147467259" Then
                REM  MsgBox("Data be updated  = " & access.Exception)

                REM  access.ExecQuery("UPDATE CPUINVO (SerialNumber,UUID,ComputerName,theDomain,Main_SCCM_Group,Manufacturer,Model,MACS)" & _
                REM                 "Values (@SerialNumber,@MyUUID,@MYComputerName,@Mydomain,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")

                If Not String.IsNullOrEmpty(access.Exception) Then
                    MsgBox(access.Exception.ToString)
                End If


            Else

                MsgBox("Err! = " & access.Exception)

            End If

        End If

    End Sub


    Private Sub UpdateRecord(MyCurrentSystem)
        'Updata Parameters - Order maters!

        access.AddParam("@MYComputerName", MyCurrentSystem.MComputerName)
        REM MsgBox(MComputerName)

        access.AddParam("@MYdomain", MyCurrentSystem.MyDomain)
        REM MsgBox(MyDomain)

        access.AddParam("@MainSCCMCollection", MyCurrentSystem.MyPrimare_Collection)
        REM  MsgBox(MyPrimare_Collection)

        REM access.AddParam("@Site_Computername", MyCurrentSystem.

        access.ExecQuery("UPDATE CPUINVO (SerialNumber,UUID,ComputerName,theDomain,Main_SCCM_Group,Manufacturer,Model,MACS)" & _
                                 "Values (@SerialNumber,@MyUUID,@MYComputerName,@Mydomain,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")

    End Sub



    Public Sub UpDateSystem(MyPrimare_Collection)
        MsgBox(MyPrimare_Collection) 'Debug 
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


        REM access.AddParam("@Site_Computername", MyMACs)
        REM access.AddParam("@MyDate", Now())
        REM  MsgBox(Now())



        access.ExecQuery("UPDATE members " & _
                        "SET username=@user,[Password]=@pass,email=@email,website=@website,Active=@active " & _
                        "WHERE id=@uid")


        If Not String.IsNullOrEmpty(access.Exception) Then

            If access.Exception.ToString = "-2147467259" Then

                REM  If access.Exception = "-2147467259" Then
                MsgBox("Data be updated  = " & access.Exception)

                REM  access.ExecQuery("UPDATE CPUINVO (SerialNumber,UUID,ComputerName,theDomain,Main_SCCM_Group,Manufacturer,Model,MACS)" & _
                REM                    "Values (@SerialNumber,@MyUUID,@MYComputerName,@Mydomain,@MainSCCMCollection,@MyNamufaturer,@MyModel,@MyMACS); ")
                REM End If


            Else

                MsgBox("Shite ! = " & access.Exception)

            End If

        End If









    End Sub

End Module
