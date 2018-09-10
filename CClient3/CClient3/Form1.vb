Public Class CClient_Freefall
    Private Access As New DBControle

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CClient_Freefall_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Dim Site_Selection As DataSet


        Access.ExecQuery("Select * from Primary_Site_Profile")
        If Not String.IsNullOrEmpty(Access.Exception) Then

            MsgBox(Access.Exception)

        Else
            For Each Record As DataRow In Access.Data_Table.Rows




        End If














        REM curren set data set by Cserver 
        Dim CurrentSet = ""
        Dim Stop_for_Numbering As Boolean = False
        Dim DateSet As Date = Now
        REM 
        REM getting current set data for database 
        Access.ExecQuery("Select * from Currentset")
        If Not String.IsNullOrEmpty(Access.Exception) Then

            MsgBox(Access.Exception)

        Else

            For Each Record As DataRow In Access.Data_Table.Rows
                CurrentSet = Record("CSID")
                Stop_for_Numbering = Record("SFC")
                DateSet = Record("setdata")
            Next

        End If
        REM **********************************************

        ComboBox1.SelectedText = CurrentSet



        Dim MyList As Array




    End Sub


End Class
