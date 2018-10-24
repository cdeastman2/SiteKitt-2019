
Imports System.Net

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Print_CartTags("Test-", "Cart", 20)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim My_Name_is As String = ""
        Dim Incident As String = "1234567890"
        Dim txtSite As String = "FUSD"
        Dim txtLocation As String = "Room 01"
        Dim txtInvoTrackingNumber As String = "9876543210"
        Dim txtNotes As String = "Some Notes "

        My_Name_is = Get_My_name()
        TextBox1.Text = My_Name_is


        Call Print_ServiceTag(My_Name_is, Incident, txtSite, txtLocation, txtInvoTrackingNumber, txtNotes)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        TextBox1.Text = IPAddress.Any.ToString()

        TextBox2.Text = IPAddress.Broadcast.ToString()

        TextBox3.Text = IPAddress.Loopback.ToString()

        TextBox4.Text = IPAddress.None.ToString()

        Dim hostName As [String] = IPAddress.IPv6Loopback.ToString








        
        TextBox5.Text = hostName










    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub
End Class
