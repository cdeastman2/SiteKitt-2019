Public Class Form1
    Dim strFlag As String
    Dim strTeacherComputer As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each ParaMe As String In My.Application.CommandLineArgs

            strFlag = LCase(LSet(ParaMe, 3))




            ListBox1.Items.Add(strFlag)

            Select Case strFlag

                Case "/ip"

                    strTeacherComputer = RSet(ParaMe, Len(ParaMe))

                    ListBox1.Items.Add(strTeacherComputer)

            End Select




        Next
    End Sub

    

End Class

