Dim connection
Dim computer
Dim userName
Dim userPassword
Dim password 'Password object

Wscript.StdOut.Write "Computer you want to connect to (Enter . for local): "
computer = WScript.StdIn.ReadLine

If computer = "." Then
    userName = ""
    userPassword = ""
Else
    Wscript.StdOut.Write "Please enter the user name: "
    userName = WScript.StdIn.ReadLine

    Set password = CreateObject("ScriptPW.Password") 
    WScript.StdOut.Write "Please enter your password:" 
    userPassword = password.GetPassword() 
End If

Set connection = Connect(computer,userName,userPassword)

If Err.Number<>0 Then
    Wscript.Echo "Call to connect failed"
End If

Call SNIPPETMETHODNAME (connection)

Sub SNIPPETMETHODNAME(connection)
   ' Insert snippet code here.
End Sub

Function Connect(server, userName, userPassword)
    On Error Resume Next

    Dim net
    Dim localConnection
    Dim swbemLocator
    Dim swbemServices
    Dim providerLoc
    Dim location

    Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")

    swbemLocator.Security_.AuthenticationLevel = 6 'Packet Privacy

    ' If  the server is local, don not supply credentials.
    Set net = CreateObject("WScript.NetWork") 
    If UCase(net.ComputerName) = UCase(server) Then
        localConnection = true
        userName = ""
        userPassword = ""
        server = "."
    End If

    ' Connect to the server.
    Set swbemServices= swbemLocator.ConnectServer _
            (server, "root\sms",userName,userPassword)
    If Err.Number<>0 Then
        Wscript.Echo "Couldn't connect: " + Err.Description
        Connect = null
        Exit Function
    End If


    ' Determine where the provider is and connect.
    Set providerLoc = swbemServices.InstancesOf("SMS_ProviderLocation")

        For Each location In providerLoc
            If location.ProviderForLocalSite = True Then
                Set swbemServices = swbemLocator.ConnectServer _
                 (location.Machine, "root\sms\site_" + _
                    location.SiteCode,userName,userPassword)
                If Err.Number<>0 Then
                    Wscript.Echo "Couldn't connect:" + Err.Description
                    Connect = Null
                    Exit Function
                End If
                Set Connect = swbemServices
                Exit Function
            End If
        Next
    Set Connect = null ' Failed to connect.
End Function