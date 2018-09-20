Public Class CClient_Freefall
    Private Access As New DBControle
    Dim Current_System_Info As New SCCMDataPoint

    Dim Site_Selection As DataTable
    Dim Collection_Selection As DataTable
    Dim My_Systems_data As DataTable

    REM This is the primary user inputted data every thing else in gathered from computer
    Dim My_site As String = Nothing
    Dim My_collection_Selection As String = Nothing
    Dim MY_sITE_SYSTEM_NAME As String = Nothing
    Dim mY_lmc_BARCODE As String = Nothing
    REM **********************************************






    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CClient_Freefall_MenuStart(sender As Object, e As EventArgs) Handles Me.MenuStart

    End Sub

    Private Sub CClient_Freefall_MinimumSizeChanged(sender As Object, e As EventArgs) Handles Me.MinimumSizeChanged

    End Sub

    Private Sub CClient_Freefall_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter

    End Sub

    Private Sub CClient_Freefall_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Timer1.Stop()
        Timer1.Enabled = False
    End Sub

    Private Sub CClient_Freefall_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles Me.QueryAccessibilityHelp

    End Sub

    Private Sub CClient_Freefall_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        REM Setup ***************************************************
        ProgressBar1.Value = ProgressBar1.Maximum
        ProgressBar1.BackColor = Color.BlueViolet


        REM curren set data set by Cserver 
        Dim CurrentSet = ""
        Dim Stop_for_Numbering As Boolean = False
        Dim DateSet As Date = Now
        REM 
        REM getting current set data for database 
        REM **********************************************

        REM setting data









        REM  MsgBox(Current_System_Info.SerialNumber) 'debug
        Access.ExecQuery("Select * from cpuinvo where SerialNumber = '" & Current_System_Info.SerialNumber & "'")
        If Not String.IsNullOrEmpty(Access.Exception) Then

            MsgBox(Access.Exception) 'debug
        Else
            My_Systems_data = Access.Data_Table
        End If
        REM **********************************************















        REM setup site list 
        Access.ExecQuery("Select * FROM Primary_site_Profile")
        If Not String.IsNullOrEmpty(Access.Exception) Then

            MsgBox(Access.Exception)
        Else
            Site_Selection = Access.Data_Table
        End If
        REM **********************************************

        REM seup current sellection
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

        REM setup Collection list 
        Access.ExecQuery("Select * FROM Collections_Profile order by site")
        If Not String.IsNullOrEmpty(Access.Exception) Then
            MsgBox(Access.Exception)
        Else
            Collection_Selection = Access.Data_Table
        End If
        Access = Nothing
        REM **********************************************



        REM DISPLAYING DATA AND INFO

        ComboBox1.SelectedText = CurrentSet
        ComboBox1.Items.Add(CurrentSet)

        For Each Collection As DataRow In Collection_Selection.Rows
            ComboBox1.Items.Add(Collection("Collection_ID"))
        Next

        RadioButton1.PerformClick()
        ListBox1.Items.Add(Current_System_Info.SerialNumber)
        ListBox1.Items.Add(Current_System_Info.DNSHostName)
        ListBox1.Items.Add(Current_System_Info.Domian)
        ListBox1.Items.Add(Current_System_Info.Manufacturer)
        ListBox1.Items.Add(Current_System_Info.Model)
        REM  ListBox1.Items.Add(Current_System_Info.MACAddress.Length - 1) DEBUG <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 0 To Current_System_Info.MACAddress.Length - 1
            Try
                ListBox1.Items.Add(Current_System_Info.MACAddress(i))
            Catch ex As Exception

            End Try

        Next
        REM **********************************************

        My_collection_Selection = CurrentSet
        REM  ListBox1.Items.Add(My_collection_Selection)'debug <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        Timer1.Enabled = True







    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        REM  Access.ExecQuery("Select * from Primary_Site_Profile")

        REM  If Not String.IsNullOrEmpty(Access.Exception) Then

        REM MsgBox(Access.Exception)

        REM Else

        REM    Site_Selection = Access.Data_Table
        For Each accRecord As DataRow In Site_Selection.Rows
            CBSite_Selection.Items.Add(accRecord("Site"))
        Next

        REM  End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        REM Access.ExecQuery("Select * from Primary_Site_Profile")

        CBSite_Selection.Items.Clear()
        REM Site_Selection = Access.Data_Table

            For Each accRecord As DataRow In Site_Selection.Rows
                If accRecord("Sitetype") = 1 Then
                    CBSite_Selection.Items.Add(accRecord("Site"))
                End If

            Next


    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        CBSite_Selection.Items.Clear()
        REM Site_Selection = Access.Data_Table

        For Each accRecord As DataRow In Site_Selection.Rows
            If accRecord("Sitetype") = 2 Then
                CBSite_Selection.Items.Add(accRecord("Site"))
            End If
        Next
    End Sub

    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        CBSite_Selection.Items.Clear()
        REM Site_Selection = Access.Data_Table
        For Each accRecord As DataRow In Site_Selection.Rows
            If accRecord("Sitetype") = 3 Then
                CBSite_Selection.Items.Add(accRecord("Site"))
            End If
        Next

    End Sub

    Private Sub RadioButton4_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click
        CBSite_Selection.Items.Clear()
        REM  Site_Selection = Access.Data_Table
        For Each accRecord As DataRow In Site_Selection.Rows
            If accRecord("Sitetype") = 4 Then
                CBSite_Selection.Items.Add(accRecord("Site"))
            End If
        Next
    End Sub

    Private Sub CBSite_Selection_GotFocus(sender As Object, e As EventArgs) Handles CBSite_Selection.GotFocus
        Timer1.Stop()
        Timer1.Enabled = False
    End Sub

    Private Sub CBSite_Selection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBSite_Selection.SelectedIndexChanged

    End Sub

    Private Sub CBSite_Selection_SelectedValueChanged(sender As Object, e As EventArgs) Handles CBSite_Selection.SelectedValueChanged
        My_site = CBSite_Selection.Text
        ListBox1.Items.Add(My_site)
        ComboBox1.Items.Clear()

        For Each Site_Collection_group As DataRow In Collection_Selection.Rows

            If Site_Collection_group("site") = My_site Then
                ComboBox1.Items.Add(Site_Collection_group("Collection_id"))
            End If


        Next




    End Sub

    Private Sub ComboBox1_DropDownStyleChanged(sender As Object, e As EventArgs) Handles ComboBox1.DropDownStyleChanged

    End Sub

    Private Sub RadioButton5_Click(sender As Object, e As EventArgs) Handles RadioButton5.Click
        CBSite_Selection.Items.Clear()
        REM  Site_Selection = Access.Data_Table
        For Each accRecord As DataRow In Site_Selection.Rows
            REM  If accRecord("Sitetype") = 4 Then
            CBSite_Selection.Items.Add(accRecord("Site"))
            REM  End If
        Next
    End Sub

    Private Sub ComboBox1_HelpRequested(sender As Object, hlpevent As HelpEventArgs) Handles ComboBox1.HelpRequested

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged









    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged


        My_collection_Selection = ComboBox1.SelectedItem
        ListBox1.Items.Add(My_collection_Selection) 'debug <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AddSystem(My_collection_Selection)
        Timer1.Dispose()
        Application.Exit()



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value > 0 Then
            ProgressBar1.Value = ProgressBar1.Value - 1
        Else
            Timer1.Stop()



            REM  btnStart_Click(sender, New EventArgs())
            Button1_Click(sender, New EventArgs())
            
        End If
    End Sub
End Class
