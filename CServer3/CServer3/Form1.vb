Public Class Form1
    REM  Private sql As New SQLControl

    REM Private sql As New SQLControl

    Private Access As New DBControle
    REM Dim Current_System_Info As New SCCMDataPoint

    Dim Site_Selection As DataTable
    Dim Collection_Selection As DataTable



    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        REM setup site list 
        Access.ExecQuery("Select * FROM Primary_site_Profile order by Site")
        If Not String.IsNullOrEmpty(Access.Exception) Then

            MsgBox(Access.Exception)
        Else
            Site_Selection = Access.Data_Table
        End If
        REM **********************************************

    End Sub
End Class
