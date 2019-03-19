Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim _RegistryKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList")
2
       For Each _KeyName As String In _RegistryKey.GetSubKeyNames()
3
           Using SubKey As RegistryKey = _RegistryKey.OpenSubKey(_KeyName)
4
               Dim _Users As String = DirectCast(SubKey.GetValue("ProfileImagePath"), String)
5
               Dim _Username As String = System.IO.Path.GetFileNameWithoutExtension(_Users)
6
               ListBox1.Items.Add(_Username)
7
           End Using
8
       Next

    End Sub
End Class
