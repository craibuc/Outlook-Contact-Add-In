<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form_Export
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Text_Destination As System.Windows.Forms.TextBox
	Public WithEvents Command_Browse_Destination As System.Windows.Forms.Button
	Public WithEvents Command_Browse_Source As System.Windows.Forms.Button
	Public WithEvents Text_Source As System.Windows.Forms.TextBox
	Public WithEvents CommandButton_Cancel As System.Windows.Forms.Button
	Public WithEvents CommandButton_OK As System.Windows.Forms.Button
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text_Destination = New System.Windows.Forms.TextBox
        Me.Command_Browse_Destination = New System.Windows.Forms.Button
        Me.Command_Browse_Source = New System.Windows.Forms.Button
        Me.Text_Source = New System.Windows.Forms.TextBox
        Me.CommandButton_Cancel = New System.Windows.Forms.Button
        Me.CommandButton_OK = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Text_Destination
        '
        Me.Text_Destination.AcceptsReturn = True
        Me.Text_Destination.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Destination.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Destination.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Destination.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Destination.Location = New System.Drawing.Point(8, 96)
        Me.Text_Destination.MaxLength = 0
        Me.Text_Destination.Name = "Text_Destination"
        Me.Text_Destination.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Destination.Size = New System.Drawing.Size(285, 20)
        Me.Text_Destination.TabIndex = 2
        Me.Text_Destination.Text = "C:\Documents and Settings\buchacr\Desktop"
        '
        'Command_Browse_Destination
        '
        Me.Command_Browse_Destination.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Browse_Destination.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Browse_Destination.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Browse_Destination.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Browse_Destination.Location = New System.Drawing.Point(300, 96)
        Me.Command_Browse_Destination.Name = "Command_Browse_Destination"
        Me.Command_Browse_Destination.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Browse_Destination.Size = New System.Drawing.Size(72, 25)
        Me.Command_Browse_Destination.TabIndex = 3
        Me.Command_Browse_Destination.Text = "Browse..."
        Me.Command_Browse_Destination.UseVisualStyleBackColor = False
        '
        'Command_Browse_Source
        '
        Me.Command_Browse_Source.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Browse_Source.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Browse_Source.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Browse_Source.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Browse_Source.Location = New System.Drawing.Point(300, 28)
        Me.Command_Browse_Source.Name = "Command_Browse_Source"
        Me.Command_Browse_Source.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Browse_Source.Size = New System.Drawing.Size(72, 25)
        Me.Command_Browse_Source.TabIndex = 1
        Me.Command_Browse_Source.Text = "Browse..."
        Me.Command_Browse_Source.UseVisualStyleBackColor = False
        '
        'Text_Source
        '
        Me.Text_Source.AcceptsReturn = True
        Me.Text_Source.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Source.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Source.Enabled = False
        Me.Text_Source.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Source.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Source.Location = New System.Drawing.Point(8, 28)
        Me.Text_Source.MaxLength = 0
        Me.Text_Source.Name = "Text_Source"
        Me.Text_Source.ReadOnly = True
        Me.Text_Source.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Source.Size = New System.Drawing.Size(285, 20)
        Me.Text_Source.TabIndex = 0
        '
        'CommandButton_Cancel
        '
        Me.CommandButton_Cancel.BackColor = System.Drawing.SystemColors.Control
        Me.CommandButton_Cancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.CommandButton_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CommandButton_Cancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommandButton_Cancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CommandButton_Cancel.Location = New System.Drawing.Point(300, 180)
        Me.CommandButton_Cancel.Name = "CommandButton_Cancel"
        Me.CommandButton_Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CommandButton_Cancel.Size = New System.Drawing.Size(72, 25)
        Me.CommandButton_Cancel.TabIndex = 5
        Me.CommandButton_Cancel.Text = "Cancel"
        Me.CommandButton_Cancel.UseVisualStyleBackColor = False
        '
        'CommandButton_OK
        '
        Me.CommandButton_OK.BackColor = System.Drawing.SystemColors.Control
        Me.CommandButton_OK.Cursor = System.Windows.Forms.Cursors.Default
        Me.CommandButton_OK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommandButton_OK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CommandButton_OK.Location = New System.Drawing.Point(224, 180)
        Me.CommandButton_OK.Name = "CommandButton_OK"
        Me.CommandButton_OK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CommandButton_OK.Size = New System.Drawing.Size(72, 25)
        Me.CommandButton_OK.TabIndex = 4
        Me.CommandButton_OK.Text = "OK"
        Me.CommandButton_OK.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(285, 17)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Save vCards to this folder:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(285, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Export Contacts in this folder:"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(13, 55)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(115, 18)
        Me.CheckBox1.TabIndex = 8
        Me.CheckBox1.Text = "Include subfolders"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Form_Export
        '
        Me.AcceptButton = Me.CommandButton_OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.CommandButton_Cancel
        Me.ClientSize = New System.Drawing.Size(378, 214)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Text_Destination)
        Me.Controls.Add(Me.Command_Browse_Destination)
        Me.Controls.Add(Me.Command_Browse_Source)
        Me.Controls.Add(Me.Text_Source)
        Me.Controls.Add(Me.CommandButton_Cancel)
        Me.Controls.Add(Me.CommandButton_OK)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_Export"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "vCard Export"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
#End Region 
End Class