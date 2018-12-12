
Imports System.Diagnostics
Module Set_Sync_Reg_defaults
    Dim Error_List
    REM Dim Parameters_List() As Array
    Dim strFlag As String
    Dim strTeacherComputer As String = ""

    Public Function Set_Sync(holer)
        Dim TheDataList(6)

        For Each ParaMe As String In My.Application.CommandLineArgs

            strFlag = LCase(Left(ParaMe, 3))

            REM  Form1.ListBox1.Items.Add(strFlag) 'debug 

            Select Case strFlag

                Case "\ip"

                    TheDataList(0) = Right(ParaMe, Len(ParaMe) - 3)
                    REM  Form1.ListBox1.Items.Add(TheDataList(0))


                Case "\cv"
                    REM TheDataList(1) = Right(ParaMe, Len(ParaMe) - 3)
                    If Right(ParaMe, Len(ParaMe) - 3) <> "0" Then
                        TheDataList(1) = 1
                    Else
                        TheDataList(1) = 0

                    End If

                    REM Form1.ListBox1.Items.Add(TheDataList(1))


            End Select

        Next

        KillSyn()

        Error_List = setSyncSeting64(TheDataList)

        StartSyn()

    End Function

    Private Function setSyncSeting64(TheData)
        REM MsgBox("setting sync")
        Dim strValueName
        Dim strValue
        Dim IP As String




        Try

            ''==========================================================	

            If TheData(0) = Nothing Then
                IP = "Mayday.staff.fusd.local"
            Else
                IP = TheData(0)
            End If

            strValueName = "ConnectIP" 'TheData(0)
            strValue = IP
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "Student Path" 'TheData(2)
            strValue = "C:\\Program Files\\SMART Technologies\\Sync Student\\"
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "Visible" 'TheData(1)
            strValue = TheData(1)
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "ActiveDirStudentIdField" 'TheData(3)
            strValue = "cn"
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "StudentID" 'TheData(4
            strValue = "AnonymousID"
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "NamingServerLoc" 'TheData(5)
            strValue = ""
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "ConnectTeacherID" 'TheData(6)
            strValue = ""
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "RedrawHooks"
            strValue = &H3E8
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "LanguageID"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "MirrorDriver"
            strValue = &H3E8
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "MulticastTTL"
            strValue = &H20
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "UnicastNoDelay"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "StudentIDMode"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "PasswordHash"
            strValue = ""
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "NICListLength"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "UseNamingServer"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "NTGroupListLength"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "CtrlAltDelSettings"
            strValue = &H2
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "EnableFileTransfer"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "Build Version"
            strValue = "10.0.574.0"
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================


            ''==========================================================
            strValueName = "EnableHelp"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "DisplayExit"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "StoreFilesToMyDocs"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)

            strValueName = "AutoStart"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "CustomSharedFolder"
            strValue = ""
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================
            ''==========================================================

            strValueName = "SecurityUsed"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "ConnectionUsed"
            strValue = &H3
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "EnableNICDefaultOrder"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "EnableQuestions"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "EnableChat"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "NamingServerPassedTest"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "HWGPUCapture"
            strValue = &H3E8
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "GPUBoolEnableTimings"
            strValue = &H0
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "GPUBoolMergeRegions"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "GPUTileCount"
            strValue = &H1
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "GPUScaleFactor"
            strValue = &H3
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================

            ''==========================================================
            strValueName = "GPUThresholdPixelCorrection"
            strValue = &H3E8
            My.Computer.Registry.SetValue("HKEY_Local_MACHINE\SOFTWARE\Wow6432Node\SMART Technologies\SMART Sync Student", strValueName, strValue)
            ''==========================================================
        Catch ex As Exception

        End Try
    End Function

    Private Sub StartSyn()
        Process.Start("dax64.exe")
    End Sub

    Private Sub KillSyn()
        REM MsgBox("killing Sync") 'debug

        For Each p As Process In Process.GetProcessesByName("dax64")
            REM MsgBox(p.Id) ' debug
            p.Kill()
            p.Close() REM 
        Next

       






    End Sub

    Private Sub CloseSync()
        Do Until Process.GetProcessesByName("dax64").Count < 1
            For Each p As Process In Process.GetProcessesByName("dax64")
                p.CloseMainWindow()
            Next
        Loop
    End Sub

End Module
