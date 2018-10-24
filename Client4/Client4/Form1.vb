Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        REM Call Reg.Main()



        Dim regVersion = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Aprien\\TestApp\\1.0", True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            MsgBox(" ' Key doesn't exist; create it")
            regVersion = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\\Aprien\\TestApp\\1.0")
        Else
            MsgBox(" ' Key  exist; create it")

        End If

        Dim intVersion As Integer = 0
        If regVersion IsNot Nothing Then
            intVersion = regVersion.GetValue("Version", 0)
            intVersion = intVersion + 1
            regVersion.SetValue("Version", intVersion)
            regVersion.Close()
        End If

    End Sub
End Class
