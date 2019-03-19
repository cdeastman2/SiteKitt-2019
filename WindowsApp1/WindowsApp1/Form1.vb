Imports System.Drawing.Printing
Imports System.Management

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each s As String In PrinterSettings.InstalledPrinters
            MessageBox.Show(s)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


    End Sub
End Class
