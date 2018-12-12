Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Hide()

        Dim TheReturn
        TheReturn = Set_Sync()
        Application.Exit()
        End



    End Sub
End Class
