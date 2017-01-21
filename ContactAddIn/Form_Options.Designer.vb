<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Options
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.Button_OK = New System.Windows.Forms.Button
        Me.Button_Apply = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.NumericUpDown_Delay = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown_GroupSize = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.CheckBox_SendInGroups = New System.Windows.Forms.CheckBox
        Me.TextBox_MessageFolder = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button_BrowseMessages = New System.Windows.Forms.Button
        Me.Button_BrowseContacts = New System.Windows.Forms.Button
        Me.TextBox_ContactFolder = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown_Delay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_GroupSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button_Cancel
        '
        Me.Button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Cancel.Location = New System.Drawing.Point(224, 335)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.Button_Cancel.TabIndex = 2
        Me.Button_Cancel.Text = "Cancel"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Button_OK
        '
        Me.Button_OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button_OK.Location = New System.Drawing.Point(143, 335)
        Me.Button_OK.Name = "Button_OK"
        Me.Button_OK.Size = New System.Drawing.Size(75, 23)
        Me.Button_OK.TabIndex = 1
        Me.Button_OK.Text = "OK"
        Me.Button_OK.UseVisualStyleBackColor = True
        '
        'Button_Apply
        '
        Me.Button_Apply.Enabled = False
        Me.Button_Apply.Location = New System.Drawing.Point(305, 335)
        Me.Button_Apply.Name = "Button_Apply"
        Me.Button_Apply.Size = New System.Drawing.Size(75, 23)
        Me.Button_Apply.TabIndex = 3
        Me.Button_Apply.Text = "Apply"
        Me.Button_Apply.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(368, 317)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.TextBox_MessageFolder)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Button_BrowseMessages)
        Me.TabPage1.Controls.Add(Me.Button_BrowseContacts)
        Me.TabPage1.Controls.Add(Me.TextBox_ContactFolder)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(360, 291)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "General"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.NumericUpDown_Delay)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown_GroupSize)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CheckBox_SendInGroups)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 161)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(188, 100)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Message Grouping"
        '
        'NumericUpDown_Delay
        '
        Me.NumericUpDown_Delay.Location = New System.Drawing.Point(90, 68)
        Me.NumericUpDown_Delay.Name = "NumericUpDown_Delay"
        Me.NumericUpDown_Delay.Size = New System.Drawing.Size(84, 20)
        Me.NumericUpDown_Delay.TabIndex = 4
        Me.NumericUpDown_Delay.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'NumericUpDown_GroupSize
        '
        Me.NumericUpDown_GroupSize.Location = New System.Drawing.Point(90, 42)
        Me.NumericUpDown_GroupSize.Name = "NumericUpDown_GroupSize"
        Me.NumericUpDown_GroupSize.Size = New System.Drawing.Size(84, 20)
        Me.NumericUpDown_GroupSize.TabIndex = 3
        Me.NumericUpDown_GroupSize.Value = New Decimal(New Integer() {50, 0, 0, 0})
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Delay (minutes)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Group Size"
        '
        'CheckBox_SendInGroups
        '
        Me.CheckBox_SendInGroups.AutoSize = True
        Me.CheckBox_SendInGroups.Checked = True
        Me.CheckBox_SendInGroups.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_SendInGroups.Location = New System.Drawing.Point(10, 19)
        Me.CheckBox_SendInGroups.Name = "CheckBox_SendInGroups"
        Me.CheckBox_SendInGroups.Size = New System.Drawing.Size(150, 17)
        Me.CheckBox_SendInGroups.TabIndex = 0
        Me.CheckBox_SendInGroups.Text = "Send Messages in Groups"
        Me.CheckBox_SendInGroups.UseVisualStyleBackColor = True
        '
        'TextBox_MessageFolder
        '
        Me.TextBox_MessageFolder.Location = New System.Drawing.Point(11, 77)
        Me.TextBox_MessageFolder.Name = "TextBox_MessageFolder"
        Me.TextBox_MessageFolder.ReadOnly = True
        Me.TextBox_MessageFolder.Size = New System.Drawing.Size(262, 20)
        Me.TextBox_MessageFolder.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Default Message folder:"
        '
        'Button_BrowseMessages
        '
        Me.Button_BrowseMessages.Location = New System.Drawing.Point(279, 75)
        Me.Button_BrowseMessages.Name = "Button_BrowseMessages"
        Me.Button_BrowseMessages.Size = New System.Drawing.Size(75, 23)
        Me.Button_BrowseMessages.TabIndex = 5
        Me.Button_BrowseMessages.Text = "Browse..."
        Me.Button_BrowseMessages.UseVisualStyleBackColor = True
        '
        'Button_BrowseContacts
        '
        Me.Button_BrowseContacts.Location = New System.Drawing.Point(279, 27)
        Me.Button_BrowseContacts.Name = "Button_BrowseContacts"
        Me.Button_BrowseContacts.Size = New System.Drawing.Size(75, 23)
        Me.Button_BrowseContacts.TabIndex = 2
        Me.Button_BrowseContacts.Text = "Browse..."
        Me.Button_BrowseContacts.UseVisualStyleBackColor = True
        '
        'TextBox_ContactFolder
        '
        Me.TextBox_ContactFolder.Location = New System.Drawing.Point(11, 29)
        Me.TextBox_ContactFolder.Name = "TextBox_ContactFolder"
        Me.TextBox_ContactFolder.ReadOnly = True
        Me.TextBox_ContactFolder.Size = New System.Drawing.Size(262, 20)
        Me.TextBox_ContactFolder.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Default Contact folder:"
        '
        'Form_Options
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Cancel
        Me.ClientSize = New System.Drawing.Size(392, 366)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button_Apply)
        Me.Controls.Add(Me.Button_OK)
        Me.Controls.Add(Me.Button_Cancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_Options"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Options"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumericUpDown_Delay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_GroupSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_OK As System.Windows.Forms.Button
    Friend WithEvents Button_Apply As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button_BrowseContacts As System.Windows.Forms.Button
    Friend WithEvents TextBox_ContactFolder As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_MessageFolder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button_BrowseMessages As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown_Delay As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_GroupSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_SendInGroups As System.Windows.Forms.CheckBox
End Class
