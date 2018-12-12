Imports System.IO
Imports System.Data.OleDb
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        REM UpdateDgv()
        watch()
    End Sub


    Private Sub UpdateDgv()
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = GetDatatableFromAccess()
    End Sub

    Private Sub UpdateDgvThread()
        RemoveHandler watcher.Changed, AddressOf OnChanged
        DataGridView1.Invoke(New Action(AddressOf UpdateDgv))
        AddHandler watcher.Changed, AddressOf OnChanged
    End Sub

    Public watcher As FileSystemWatcher

    Private Sub watch()
        watcher = New FileSystemWatcher()
        watcher.Path = "E:\Alpha_Profiles\Data\"
        watcher.NotifyFilter = NotifyFilters.LastWrite
        watcher.Filter = "*.*"
        AddHandler watcher.Changed, AddressOf OnChanged
        watcher.EnableRaisingEvents = True
    End Sub


    Private Sub OnChanged(ByVal source As Object, ByVal e As FileSystemEventArgs)
        Debug.WriteLine("file was changed")
        ' update datagridview
        Dim t1 As New System.Threading.Thread(AddressOf UpdateDgvThread)
        t1.Start()
    End Sub

    Public Function GetDatatableFromAccess() As DataTable

        Dim connString As String =
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\john\Desktop\Desktop 01-04-2017\Database11.accdb"

        Dim results As New DataTable()

        Using conn As New OleDbConnection(connString)
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM TestTable", conn)
            conn.Open()
            Dim adapter As New OleDbDataAdapter(cmd)
            adapter.Fill(results)
        End Using

        Return results
    End Function

End Class