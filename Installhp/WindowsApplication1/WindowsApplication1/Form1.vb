Imports System.Diagnostics
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim theRetrun
        Dim data(6)
        REM data(0) = "My-cde.staff.fusd.loacl"
        theRetrun = Set_Sync(data)
        Application.Exit()
        End










    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub gogo()

     
    End Sub

    Private Sub load_prosses()
     

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Call load_prosses()
        Call gogo()
        Call load_prosses()
    End Sub

End Class
