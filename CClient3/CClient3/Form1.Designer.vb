<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CClient_Freefall
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CClient_Freefall))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox0 = New System.Windows.Forms.GroupBox()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CBSite_Selection = New System.Windows.Forms.ComboBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox0.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 25
        Me.ListBox1.Location = New System.Drawing.Point(17, 266)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(611, 329)
        Me.ListBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 238)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(201, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Client System Data "
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(17, 191)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(300, 33)
        Me.ComboBox1.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 163)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(184, 25)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Current Collection"
        '
        'GroupBox0
        '
        Me.GroupBox0.Controls.Add(Me.RadioButton5)
        Me.GroupBox0.Controls.Add(Me.RadioButton4)
        Me.GroupBox0.Controls.Add(Me.RadioButton3)
        Me.GroupBox0.Controls.Add(Me.RadioButton1)
        Me.GroupBox0.Controls.Add(Me.RadioButton2)
        Me.GroupBox0.Location = New System.Drawing.Point(17, 12)
        Me.GroupBox0.Name = "GroupBox0"
        Me.GroupBox0.Size = New System.Drawing.Size(778, 83)
        Me.GroupBox0.TabIndex = 14
        Me.GroupBox0.TabStop = False
        Me.GroupBox0.Text = "Site Filter"
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(647, 30)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(118, 29)
        Me.RadioButton5.TabIndex = 4
        Me.RadioButton5.Text = "Full List"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(524, 30)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(107, 29)
        Me.RadioButton4.TabIndex = 3
        Me.RadioButton4.Text = "Others"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(346, 30)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(159, 29)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.Text = "High School"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(12, 30)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(144, 29)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.Text = "Elemetries"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(162, 30)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(178, 29)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Midle Schools"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 25)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Sellect Site"
        '
        'CBSite_Selection
        '
        Me.CBSite_Selection.FormattingEnabled = True
        Me.CBSite_Selection.Location = New System.Drawing.Point(17, 124)
        Me.CBSite_Selection.Name = "CBSite_Selection"
        Me.CBSite_Selection.Size = New System.Drawing.Size(484, 33)
        Me.CBSite_Selection.TabIndex = 12
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(17, 651)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(307, 31)
        Me.TextBox1.TabIndex = 15
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(350, 651)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(278, 31)
        Me.TextBox2.TabIndex = 16
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 623)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(210, 25)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Site Computer Name"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(358, 623)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(143, 25)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "LMC Barcode"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(17, 712)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 70)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "Done"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 839)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(783, 23)
        Me.ProgressBar1.TabIndex = 20
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(350, 703)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(142, 70)
        Me.Button2.TabIndex = 21
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CClient_Freefall
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 904)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.GroupBox0)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CBSite_Selection)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CClient_Freefall"
        Me.Text = "CClient Form"
        Me.GroupBox0.ResumeLayout(False)
        Me.GroupBox0.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox0 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CBSite_Selection As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button2 As System.Windows.Forms.Button

End Class
