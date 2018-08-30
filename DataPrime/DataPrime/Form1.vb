Public Class Form1
    Private Access As New DBControle
    Dim Site_Filter As Integer = 0
    Dim Current_Site As String = Nothing
    Dim Current_Collection As String = Nothing

    Private Sub Button2_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Form1_MenuStart(sender As Object, e As EventArgs) Handles Me.MenuStart

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        REM Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by sitetype , site asc")
        REM If Not String.IsNullOrEmpty(Access.Exception) Then MsgBox(Access.Exception) : Exit Sub
        REM    DataGridView1.DataSource = Access.Data_Table

        REM  For Each r As DataRow In Access.Data_Table.Rows
        REM ComboBox1.Items.Add(r("Site"))
        REM  Next



    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        REM  Dim index As Integer
        REM index = e.RowIndex
        REM Dim selectedRow As DataGridViewRow
        REM selectedRow = DataGridView1.Rows(index)
        REM TextBox1.Text = selectedRow.Cells(0).Value.ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Current_Site = ComboBox1.Text
        REM MsgBox(MySollection)
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click
        
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        

    End Sub

    Private Sub TabPage3_GotFocus(sender As Object, e As EventArgs) Handles TabPage3.GotFocus
        
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub RadioButton1_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

    End Sub

    Private Sub TabPage1_DragEnter(sender As Object, e As DragEventArgs) Handles TabPage1.DragEnter

    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        REM  ComboBox1.SelectedText = Nothing 'debug
        REM  ComboBox1.SelectedText = Current_Site
        Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site") ' rem not needed

        If Not String.IsNullOrEmpty(Access.Exception) Then MsgBox(Access.Exception) : Exit Sub 'checking for error thrown in class 



        For Each r As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(r("SiteType")) 'debug 
            If Site_Filter = 0 Then
                ComboBox1.Items.Add(r("Site"))
            Else
                If r("SiteType") = Site_Filter Then
                    ComboBox1.Items.Add(r("Site"))
                End If

            End If

        Next


        REM DataGridView1.DataSource = Access.Data_Table



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        Site_Filter = 3
        REM    ListBox2.Items.Clear() 'Debug
        ComboBox1.Items.Clear()
        REM  Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site")
        For Each Dataitem As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(Dataitem("SiteType")) 'debug 
            If Dataitem("SiteType") = Site_Filter Then
                ComboBox1.Items.Add(Dataitem("Site"))
            End If
        Next
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox0.Enter
        Site_Filter = 1





    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click
        Site_Filter = 4

        REM    ListBox2.Items.Clear() 'Debug
        ComboBox1.Items.Clear()
        REM  Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site")
        For Each Dataitem As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(Dataitem("SiteType")) 'debug 
            If Dataitem("SiteType") = Site_Filter Then
                ComboBox1.Items.Add(Dataitem("Site"))
            End If
        Next
    End Sub

    Private Sub RadioButton1_Click1(sender As Object, e As EventArgs) Handles RadioButton1.Click
        Site_Filter = 1
        REM    ListBox2.Items.Clear() 'Debug
        ComboBox1.Items.Clear()
        REM ComboBox1.SelectedText = Current_Site

        REM  Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site")
        For Each Dataitem As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(Dataitem("SiteType")) 'debug 
            If Dataitem("SiteType") = Site_Filter Then
                ComboBox1.Items.Add(Dataitem("Site"))
            End If
        Next

    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        Site_Filter = 2
        REM    ListBox2.Items.Clear() 'Debug
        ComboBox1.Items.Clear()
        REM  Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site")
        For Each Dataitem As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(Dataitem("SiteType")) 'debug 
            If Dataitem("SiteType") = Site_Filter Then
                ComboBox1.Items.Add(Dataitem("Site"))
            End If
        Next
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub RadioButton9_Click(sender As Object, e As EventArgs) Handles RadioButton9.Click

        REM    ListBox2.Items.Clear() 'Debug
        ComboBox1.Items.Clear()
        REM  Access.ExecQuery("SELECT * FROM Primary_Site_Profile order by site")
        For Each Dataitem As DataRow In Access.Data_Table.Rows
            REM ListBox2.Items.Add(Dataitem("SiteType")) 'debug 
            REM If Dataitem("SiteType") = Site_Filter Then
            ComboBox1.Items.Add(Dataitem("Site"))
            REM End If
        Next
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        If Current_Site <> Nothing Then

            Access.ExecQuery("SELECT * FROM Collections_Profile where site ='" & Current_Site & "' order by Collection_ID") ' rem not needed
            If Not String.IsNullOrEmpty(Access.Exception) Then MsgBox(Access.Exception) : Exit Sub 'checking for error thrown in class 
            DataGridView1.DataSource = Access.Data_Table




        End If

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub
End Class
