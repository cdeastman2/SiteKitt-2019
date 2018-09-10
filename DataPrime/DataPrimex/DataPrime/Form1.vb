Public Class Form1
    Private Access As New DBControle

  

    Private Sub Button2_Click(sender As Object, e As EventArgs)

        Access.ExecQuery("SELECT site FROM Primary_Site_Profile order by site")
        If Not String.IsNullOrEmpty(Access.Exception) Then MsgBox(Access.Exception) : Exit Sub
        REM    DataGridView1.DataSource = Access.Data_Table

        For Each r As DataRow In Access.Data_Table.Rows
            ComboBox1.Items.Add(r("Site"))
        Next

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Access.ExecQuery("SELECT site FROM Primary_Site_Profile order by sitetype , site asc")
        If Not String.IsNullOrEmpty(Access.Exception) Then MsgBox(Access.Exception) : Exit Sub
        REM    DataGridView1.DataSource = Access.Data_Table

        For Each r As DataRow In Access.Data_Table.Rows
            ComboBox1.Items.Add(r("Site"))
        Next



    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        REM  Dim index As Integer
        REM index = e.RowIndex
        REM Dim selectedRow As DataGridViewRow
        REM selectedRow = DataGridView1.Rows(index)
        REM TextBox1.Text = selectedRow.Cells(0).Value.ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim MySollection


        MySollection = ComboBox1.Text
        REM MsgBox(MySollection)



    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        End If
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click
        
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        Dim This_Systems_dataPoints
        Dim The_system_we_are_looking_for_serialnumer = ""
        Dim my_Query_is
        Dim MyCurrentSystem = New SCCMDataPoint


        This_Systems_dataPoints = BIO_DATA()



        TextBox1.Text = This_Systems_dataPoints(0).ToString  'Debug

        my_Query_is = "SELECT * FROM CPUinvo where SerialNumber = '" & This_Systems_dataPoints(0) & "'"
        REM   my_Query_is = "SELECT * FROM CPUinvo where SerialNumber = 'G1NLCX015120017'" ' & This_Systems_dataPoints(0) & "'"



        REM MsgBox(my_Query_is) 'Debug
        REM get records

        Access.ExecQuery(my_Query_is)
        If Not String.IsNullOrEmpty(Access.Exception) Then
            MsgBox(Access.Exception) : Exit Sub
        End If

        If Not Access.Data_Table Is Nothing Then
            DataGridView2.Refresh()
            DataGridView2.DataSource = Access.Data_Table
        End If

       
        ListBox1.Items.Add(MyCurrentSystem.SerialNumber)
        ListBox1.Items.Add(MyCurrentSystem.UUID)
        ListBox1.Items.Add(MyCurrentSystem.Manufacturer)
        ListBox1.Items.Add(MyCurrentSystem.Model)
        ListBox1.Items.Add(MyCurrentSystem.DNSHostName)
        ListBox1.Items.Add(MyCurrentSystem.Domian)
        ListBox1.Items.Add(MyCurrentSystem.MACAddress(0))
        ListBox1.Items.Add(MyCurrentSystem.MACAddress(1))


        Access.ExecQuery("SELECT * FROM Currentset ")
        Dim may = Access.Data_Table


        For Each row As DataRow In may.Rows
            ListBox1.Items.Add(row.Item("CSID"))

        Next








    End Sub

    Private Sub TabPage3_GotFocus(sender As Object, e As EventArgs) Handles TabPage3.GotFocus
        
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            CheckBox3.Checked = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            CheckBox2.Checked = False
            CheckBox1.Checked = False
        End If
    End Sub
End Class
