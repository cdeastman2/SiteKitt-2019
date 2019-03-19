Imports System.Drawing.Text
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pfc As PrivateFontCollection = New PrivateFontCollection
        pfc.AddFontFile("Fonts\ennobled.ttf")
        Label1.Font = New Font(pfc.Families(0), 65, FontStyle.Regular)
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub Changemyfont(Newfou)




    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Rman As Resources.ResourceManager
        Dim pfc As PrivateFontCollection = New PrivateFontCollection
        Rman = New Resources.ResourceManager("TechTag.Resources", System.Reflection.Assembly.GetExecutingAssembly)



        pfc.AddFontFile(Rman.GetObject("IDAUTOMATIONHC39M CODE 39 BARCODE.TTf"))


        Label1.Font = New Font(pfc.Families(0), 65, FontStyle.Regular)
        Label1.ForeColor = Color.Red
        REM Label1.Text = 

    End Sub
End Class
