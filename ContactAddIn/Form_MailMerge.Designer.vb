<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form_MailMerge
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter

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
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TextBox_Filter = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_Filter = New System.Windows.Forms.TabPage
        Me.Command_Browse_Source = New System.Windows.Forms.Button
        Me.Text_Source = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabPage_Recipients = New System.Windows.Forms.TabPage
        Me.Button_SelectNone = New System.Windows.Forms.Button
        Me.Button_SelectAll = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.TabPage_Message = New System.Windows.Forms.TabPage
        Me.Button_BrowseTemplates = New System.Windows.Forms.Button
        Me.TextBox_MessageTemplate = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ListView_Drafts = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.TabPage_Summary = New System.Windows.Forms.TabPage
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.Button_Finish = New System.Windows.Forms.Button
        Me.Button_Back = New System.Windows.Forms.Button
        Me.Button_Next = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage_Filter.SuspendLayout()
        Me.TabPage_Recipients.SuspendLayout()
        Me.TabPage_Message.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_Filter
        '
        Me.TextBox_Filter.AcceptsReturn = True
        Me.TextBox_Filter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox_Filter.Location = New System.Drawing.Point(3, 82)
        Me.TextBox_Filter.Multiline = True
        Me.TextBox_Filter.Name = "TextBox_Filter"
        Me.TextBox_Filter.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox_Filter.Size = New System.Drawing.Size(551, 269)
        Me.TextBox_Filter.TabIndex = 14
        Me.TextBox_Filter.Text = "[Send Mail]=True"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 65)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 14)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Enter filter:"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage_Filter)
        Me.TabControl1.Controls.Add(Me.TabPage_Recipients)
        Me.TabControl1.Controls.Add(Me.TabPage_Message)
        Me.TabControl1.Controls.Add(Me.TabPage_Summary)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(568, 384)
        Me.TabControl1.TabIndex = 16
        '
        'TabPage_Filter
        '
        Me.TabPage_Filter.Controls.Add(Me.Command_Browse_Source)
        Me.TabPage_Filter.Controls.Add(Me.Text_Source)
        Me.TabPage_Filter.Controls.Add(Me.Label2)
        Me.TabPage_Filter.Controls.Add(Me.Label5)
        Me.TabPage_Filter.Controls.Add(Me.TextBox_Filter)
        Me.TabPage_Filter.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_Filter.Name = "TabPage_Filter"
        Me.TabPage_Filter.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Filter.Size = New System.Drawing.Size(560, 357)
        Me.TabPage_Filter.TabIndex = 2
        Me.TabPage_Filter.Text = "Filter"
        Me.TabPage_Filter.UseVisualStyleBackColor = True
        '
        'Command_Browse_Source
        '
        Me.Command_Browse_Source.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Browse_Source.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Browse_Source.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Browse_Source.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Browse_Source.Location = New System.Drawing.Point(249, 30)
        Me.Command_Browse_Source.Name = "Command_Browse_Source"
        Me.Command_Browse_Source.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Browse_Source.Size = New System.Drawing.Size(72, 25)
        Me.Command_Browse_Source.TabIndex = 18
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
        Me.Text_Source.Location = New System.Drawing.Point(9, 32)
        Me.Text_Source.MaxLength = 0
        Me.Text_Source.Name = "Text_Source"
        Me.Text_Source.ReadOnly = True
        Me.Text_Source.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Source.Size = New System.Drawing.Size(234, 20)
        Me.Text_Source.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(5, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(137, 17)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Choose a Contact folder:"
        '
        'TabPage_Recipients
        '
        Me.TabPage_Recipients.Controls.Add(Me.Button_SelectNone)
        Me.TabPage_Recipients.Controls.Add(Me.Button_SelectAll)
        Me.TabPage_Recipients.Controls.Add(Me.ListView1)
        Me.TabPage_Recipients.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_Recipients.Name = "TabPage_Recipients"
        Me.TabPage_Recipients.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Recipients.Size = New System.Drawing.Size(560, 329)
        Me.TabPage_Recipients.TabIndex = 1
        Me.TabPage_Recipients.Text = "Recipients"
        Me.TabPage_Recipients.UseVisualStyleBackColor = True
        '
        'Button_SelectNone
        '
        Me.Button_SelectNone.Location = New System.Drawing.Point(87, 6)
        Me.Button_SelectNone.Name = "Button_SelectNone"
        Me.Button_SelectNone.Size = New System.Drawing.Size(75, 23)
        Me.Button_SelectNone.TabIndex = 1
        Me.Button_SelectNone.Text = "Select None"
        Me.Button_SelectNone.UseVisualStyleBackColor = True
        '
        'Button_SelectAll
        '
        Me.Button_SelectAll.Location = New System.Drawing.Point(6, 6)
        Me.Button_SelectAll.Name = "Button_SelectAll"
        Me.Button_SelectAll.Size = New System.Drawing.Size(75, 23)
        Me.Button_SelectAll.TabIndex = 0
        Me.Button_SelectAll.Text = "Select All"
        Me.Button_SelectAll.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.AllowColumnReorder = True
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.CheckBoxes = True
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader6, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader7})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(6, 35)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(551, 288)
        Me.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView1.TabIndex = 2
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "File As"
        Me.ColumnHeader6.Width = 77
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "First Name"
        Me.ColumnHeader2.Width = 91
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Last Name"
        Me.ColumnHeader3.Width = 104
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Company"
        Me.ColumnHeader4.Width = 94
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Email"
        Me.ColumnHeader5.Width = 72
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Last Contacted"
        Me.ColumnHeader7.Width = 93
        '
        'TabPage_Message
        '
        Me.TabPage_Message.Controls.Add(Me.Button_BrowseTemplates)
        Me.TabPage_Message.Controls.Add(Me.TextBox_MessageTemplate)
        Me.TabPage_Message.Controls.Add(Me.Label1)
        Me.TabPage_Message.Controls.Add(Me.Label4)
        Me.TabPage_Message.Controls.Add(Me.ListView_Drafts)
        Me.TabPage_Message.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_Message.Name = "TabPage_Message"
        Me.TabPage_Message.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Message.Size = New System.Drawing.Size(560, 329)
        Me.TabPage_Message.TabIndex = 0
        Me.TabPage_Message.Text = "Message"
        Me.TabPage_Message.UseVisualStyleBackColor = True
        '
        'Button_BrowseTemplates
        '
        Me.Button_BrowseTemplates.Location = New System.Drawing.Point(249, 31)
        Me.Button_BrowseTemplates.Name = "Button_BrowseTemplates"
        Me.Button_BrowseTemplates.Size = New System.Drawing.Size(75, 23)
        Me.Button_BrowseTemplates.TabIndex = 2
        Me.Button_BrowseTemplates.Text = "Browse..."
        Me.Button_BrowseTemplates.UseVisualStyleBackColor = True
        '
        'TextBox_MessageTemplate
        '
        Me.TextBox_MessageTemplate.Location = New System.Drawing.Point(9, 32)
        Me.TextBox_MessageTemplate.Name = "TextBox_MessageTemplate"
        Me.TextBox_MessageTemplate.Size = New System.Drawing.Size(234, 20)
        Me.TextBox_MessageTemplate.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Choose a Message Template:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(11, 65)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(232, 17)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Select a message from the Drafts folder:"
        '
        'ListView_Drafts
        '
        Me.ListView_Drafts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView_Drafts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.ListView_Drafts.Location = New System.Drawing.Point(9, 85)
        Me.ListView_Drafts.Name = "ListView_Drafts"
        Me.ListView_Drafts.Size = New System.Drawing.Size(354, 238)
        Me.ListView_Drafts.TabIndex = 4
        Me.ListView_Drafts.UseCompatibleStateImageBehavior = False
        Me.ListView_Drafts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Subject"
        Me.ColumnHeader1.Width = 350
        '
        'TabPage_Summary
        '
        Me.TabPage_Summary.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_Summary.Name = "TabPage_Summary"
        Me.TabPage_Summary.Size = New System.Drawing.Size(560, 329)
        Me.TabPage_Summary.TabIndex = 4
        Me.TabPage_Summary.Text = "Summary"
        Me.TabPage_Summary.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Cancel.Location = New System.Drawing.Point(505, 409)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.Button_Cancel.TabIndex = 17
        Me.Button_Cancel.Text = "Cancel"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Button_Finish
        '
        Me.Button_Finish.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Finish.Location = New System.Drawing.Point(424, 409)
        Me.Button_Finish.Name = "Button_Finish"
        Me.Button_Finish.Size = New System.Drawing.Size(75, 23)
        Me.Button_Finish.TabIndex = 18
        Me.Button_Finish.Text = "&Finish"
        Me.Button_Finish.UseVisualStyleBackColor = True
        '
        'Button_Back
        '
        Me.Button_Back.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Back.Enabled = False
        Me.Button_Back.Location = New System.Drawing.Point(262, 409)
        Me.Button_Back.Name = "Button_Back"
        Me.Button_Back.Size = New System.Drawing.Size(75, 23)
        Me.Button_Back.TabIndex = 19
        Me.Button_Back.Text = "< &Back"
        Me.Button_Back.UseVisualStyleBackColor = True
        '
        'Button_Next
        '
        Me.Button_Next.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Next.Location = New System.Drawing.Point(343, 409)
        Me.Button_Next.Name = "Button_Next"
        Me.Button_Next.Size = New System.Drawing.Size(75, 23)
        Me.Button_Next.TabIndex = 20
        Me.Button_Next.Text = "&Next >"
        Me.Button_Next.UseVisualStyleBackColor = True
        '
        'Form_MailMerge
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(592, 444)
        Me.Controls.Add(Me.Button_Next)
        Me.Controls.Add(Me.Button_Back)
        Me.Controls.Add(Me.Button_Finish)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.TabControl1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_MailMerge"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "eMail Merge Wizard"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_Filter.ResumeLayout(False)
        Me.TabPage_Filter.PerformLayout()
        Me.TabPage_Recipients.ResumeLayout(False)
        Me.TabPage_Message.ResumeLayout(False)
        Me.TabPage_Message.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TextBox_Filter As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_Message As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Recipients As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Filter As System.Windows.Forms.TabPage
    Public WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ListView_Drafts As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TabPage_Summary As System.Windows.Forms.TabPage
    Public WithEvents Command_Browse_Source As System.Windows.Forms.Button
    Public WithEvents Text_Source As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Button_SelectAll As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Button_SelectNone As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button_BrowseTemplates As System.Windows.Forms.Button
    Friend WithEvents TextBox_MessageTemplate As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_Finish As System.Windows.Forms.Button
    Friend WithEvents Button_Back As System.Windows.Forms.Button
    Friend WithEvents Button_Next As System.Windows.Forms.Button
#End Region
End Class