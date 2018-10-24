Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        CBSite_Selection.Items.Clear()
        REM Site_Selection = Access.Data_Table

		
		
		
		
		
        For Each accRecord As DataRow In Site_Selection.Rows
            If accRecord("Sitetype") = 2 Then
                CBSite_Selection.Items.Add(accRecord("Site"))
            End If
        Next
    End Sub