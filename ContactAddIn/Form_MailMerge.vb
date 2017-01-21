Imports Outlook = Microsoft.Office.Interop.Outlook

Friend Class Form_MailMerge
	Inherits System.Windows.Forms.Form

    Private lvwColumnSorter As ListViewColumnSorter

    Private _Application As Outlook.Application
    Private _ContactFolder As Outlook.MAPIFolder
    Private _MessageTemplate As Outlook.MailItem
    Private MAPI As Outlook.NameSpace

    Private Form_Progress As Form_Progress
    Private CancellationPending As Boolean = False

    Public Filter As Integer = 0
    Public Recipients As Integer = 1
    Public Message As Integer = 2
    Public Summary As Integer = 3

#Region "Protected Properties"

    Protected ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' MAPI folder that will be the source of Contacts used in the Mail Merge and vCard Exporter
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property ContactFolder() As Outlook.MAPIFolder
        Get
            If _ContactFolder Is Nothing Then
                If My.Settings.DefaultContactFolderId = String.Empty Then
                    _ContactFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
                Else
                    _ContactFolder = Me.Application.GetNamespace("MAPI").GetFolderFromID(My.Settings.DefaultContactFolderId)
                End If
            End If
            Return _ContactFolder
        End Get
        Set(ByVal value As Outlook.MAPIFolder)
            If value.DefaultItemType <> Outlook.OlItemType.olContactItem Then _
                Throw New Exception("Must be a ContactItem folder.")

            _ContactFolder = value
        End Set
    End Property

    ''' <summary>
    ''' The Message Template that will be sent to each Contact
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property MessageTemplate() As Outlook.MailItem
        Get
            Return _MessageTemplate

        End Get
        Set(ByVal value As Outlook.MailItem)
            _MessageTemplate = value
        End Set
    End Property

#End Region

#Region "Navigation"

    Public Sub Initialize(ByRef application As Outlook.Application)

        Try

            _Application = application
            MAPI = application.GetNamespace("MAPI")

            Me.Text_Source.Text = CStr(Me.ContactFolder.Name)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        End Try

    End Sub
	
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        Select Case Me.TabControl1.SelectedIndex
            Case Filter
            Case Recipients
                BindRecipients()
            Case Message
            Case Summary
        End Select

    End Sub

    Private Sub Button_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Back.Click

        Me.TabControl1.SelectedIndex -= 1

    End Sub

    Private Sub Button_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Next.Click

        Me.TabControl1.SelectedIndex += 1

    End Sub

    Private Sub Button_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Button_Cancel.Click

        Me.Close()

    End Sub
	
    Private Sub Button_Finish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Finish.Click

        If Me.MessageTemplate Is Nothing Then
            MsgBox("Please select a Message Template", MsgBoxStyle.Exclamation, "Contact Add-In")
            Exit Sub
        End If

        Form_Progress = New Form_Progress
        With Form_Progress
            .Text = "Sending..."
            AddHandler .Cancelled, AddressOf ProcessingCancelled
            .Show()
        End With

        'BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        'BackgroundWorker1.RunWorkerAsync()
        Process()

    End Sub

#End Region


#Region "Filter"

    Private Sub Command_Browse_Source_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Browse_Source.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Dim applicationDataDirectory As String = Environment.SpecialFolder.ApplicationData.ToString

            'Folder = Me.Application.GetNamespace("MAPI").PickFolder
            Folder = MAPI.PickFolder

            If Folder IsNot Nothing Then

                Me.ContactFolder = Folder
                Me.Text_Source.Text = CStr(Me.ContactFolder.Name)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Folder = Nothing

        End Try

    End Sub

#End Region

#Region "Recipients"

    Private Sub BindRecipients()

        Dim Items As Outlook.Items

        Try

            Me.ListView1.Items.Clear()

            If Me.TextBox_Filter.Text = String.Empty Then
                Items = ContactFolder.Items
            Else
                Items = ContactFolder.Items.Restrict(Me.TextBox_Filter.Text)
            End If

            'For Each Item As Object In ContactFolder.Items.Restrict(Me.TextBox_Filter.Text)
            For Each Item As Object In Items

                If TypeOf Item Is Outlook.ContactItem Then

                    'cast item as contact
                    Dim contactItem As Outlook.ContactItem = CType(Item, Outlook.ContactItem)

                    Dim lvi As ListViewItem = Me.ListView1.Items.Add(contactItem.FileAs)
                    With lvi
                        .Tag = contactItem.EntryID
                        .Checked = True

                        .SubItems.AddRange(New String() {contactItem.FirstName, contactItem.LastName, contactItem.CompanyName, contactItem.Email1Address})
                        If contactItem.UserProperties("Last Contacted") IsNot Nothing Then
                            .SubItems.Add(contactItem.UserProperties("Last Contacted").Value.ToString)
                        End If
                    End With

                    contactItem = Nothing

                End If

            Next Item

            For Each item As ColumnHeader In Me.ListView1.Columns
                item.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
            Next

        Catch ex As Exception
            Throw

        Finally
            Items = Nothing

        End Try

    End Sub

    Private Sub Button_SelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_SelectAll.Click
        For Each Item As ListViewItem In Me.ListView1.Items
            Item.Checked = True
        Next
    End Sub

    Private Sub Button_SelectNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_SelectNone.Click
        For Each Item As ListViewItem In Me.ListView1.Items
            Item.Checked = False
        Next
    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick

        ' Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.ListView1.Sort()

    End Sub

#End Region

#Region "Message"

    Private Sub Button_BrowseTemplates_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_BrowseTemplates.Click

        With Me.OpenFileDialog1
            .Filter = "Outlook Templates (*.oft)|*.oft|" & "All files|*.*"
            .DefaultExt = "*.oft"
            .InitialDirectory = Environment.SpecialFolder.ApplicationData & "\Microsoft\Templates"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim Drafts As Outlook.MAPIFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
                Me.MessageTemplate = Me.Application.CreateItemFromTemplate(.FileName, Drafts)
                Me.MessageTemplate.Save()

                GetDrafts()

            End If
        End With

    End Sub

    Private Sub GetDrafts()

        Dim Drafts As Outlook.MAPIFolder

        Try

            Drafts = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderDrafts)
            Dim LI As System.Windows.Forms.ListViewItem

            For Each Item As Object In Drafts.Items
                If TypeOf Item Is Outlook.MailItem Then

                    'cast item as mail message
                    Dim mailItem As Outlook.MailItem = CType(Item, Outlook.MailItem)

                    With Me.ListView_Drafts
                        LI = .Items.Add(mailItem.EntryID, mailItem.Subject, "")
                    End With

                    mailItem = Nothing

                End If
            Next Item

            If Me.ListView_Drafts.Items.Count > 0 Then
                Me.ListView_Drafts.Items(0).Selected = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Drafts = Nothing

        End Try

    End Sub

    Private Sub ListView_Drafts_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs)

        Me.MessageTemplate = Me.Application.GetNamespace("MAPI").GetItemFromID(e.Item.Text)

    End Sub

    Private Sub ListView_Drafts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView_Drafts.SelectedIndexChanged

        If Me.ListView_Drafts.SelectedItems.Count > 0 Then

            Dim entryId As String = Me.ListView_Drafts.SelectedItems(0).Name

            Me.MessageTemplate = Me.Application.GetNamespace("MAPI").GetItemFromID(entryId)

        End If

    End Sub

#End Region

#Region "Summary"

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Form_Progress = New Form_Progress
        Form_Progress.ShowDialog()

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'alert progress dialog
        Form_Progress.ReportProgress(e.ProgressPercentage)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Form_Progress.Close()

        Dim count As Integer = CInt(e.Result)
        MsgBox(String.Format("Processing Completed.  {0} messages were sent.", count), MsgBoxStyle.Information, "eMail Merge Wizard")

    End Sub

    Private Sub Process()

        Try

            Dim mailer As New ContactMailer(Me.Application)
            With mailer
                .MessageTemplate = Me.MessageTemplate

                .SendInGroups = My.Settings.SendMessagesInGroups
                .GroupSize = My.Settings.MessageGroupSize
                .MessageDelay = My.Settings.MessageDelay

                AddHandler .Processed, AddressOf ProcessingCompleted

                .Recurse = False
                '.Process(Me.ContactFolder, Me.TextBox_Filter.Text)

                Dim i As Integer = 0

                For Each Item As ListViewItem In Me.ListView1.CheckedItems

                    'increment
                    i += 1

                    'If BackgroundWorker1.CancellationPending Then Exit For
                    If CancellationPending Then Exit For

                    'notify UI
                    Dim p As Integer = CInt(i / Me.ListView1.CheckedItems.Count * 100)
                    Form_Progress.ReportProgress(p)

                    'send message
                    Dim contactItem As Outlook.ContactItem = MAPI.GetItemFromID(Item.Tag)
                    .SendMessage(contactItem)
                    contactItem = Nothing

                    System.Windows.Forms.Application.DoEvents()

                Next

                'capture the number of messages sent
                'e.Result = i
                Form_Progress.Close()

                MsgBox(String.Format("Processing Completed.  {0} messages were sent.", i), MsgBoxStyle.Information, "eMail Merge Wizard")
                Me.Close()

            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "eMail Merge Wizard")

        Finally
            'remove draft
            Me.MessageTemplate.Delete()

        End Try

    End Sub

    Private Sub ProcessingCancelled(ByVal sender As System.Object, ByVal e As EventArgs)

        CancellationPending = True

    End Sub

    Private Sub ProcessingCompleted(ByVal sender As System.Object, ByVal e As MailMergeEventArgs)

        Form_Progress.Close()

        MsgBox(String.Format("Processing Completed.  {0} messages were sent.", 1), MsgBoxStyle.Information, "eMail Merge Wizard")
        Me.Close()

    End Sub

#End Region

End Class