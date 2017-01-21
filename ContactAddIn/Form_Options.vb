Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class Form_Options

    Private _Application As Outlook.Application
    Private _ContactFolder As Outlook.MAPIFolder
    Private _MessageFolder As Outlook.MAPIFolder

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

            Me.Button_Apply.Enabled = True

        End Set
    End Property

    ''' <summary>
    ''' MAPI folder that will be the source of Messages used by the Contact Extractor.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property MessageFolder() As Outlook.MAPIFolder
        Get
            If _MessageFolder Is Nothing Then
                If My.Settings.DefaultMessageFolderId = String.Empty Then
                    _MessageFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)
                Else
                    _MessageFolder = Me.Application.GetNamespace("MAPI").GetFolderFromID(My.Settings.DefaultMessageFolderId)
                End If
            End If

            Return _MessageFolder
        End Get
        Set(ByVal value As Outlook.MAPIFolder)
            If value.DefaultItemType <> Outlook.OlItemType.olMailItem Then _
                Throw New Exception("Must be a MailItem folder.")

            _MessageFolder = value

            Me.Button_Apply.Enabled = True

        End Set
    End Property

#End Region

    Public Sub Initialize(ByRef application As Outlook.Application)

        Try

            _Application = application

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        End Try

    End Sub

    Private Sub Form_Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.TextBox_ContactFolder.Text = CStr(Me.ContactFolder.Name)
        Me.TextBox_MessageFolder.Text = CStr(Me.MessageFolder.Name)

        Me.CheckBox_SendInGroups.Checked = My.Settings.SendMessagesInGroups
        Me.NumericUpDown_GroupSize.Value = My.Settings.MessageGroupSize
        Me.NumericUpDown_Delay.Value = My.Settings.MessageDelay

        Me.Button_Apply.Enabled = False

    End Sub

    Private Sub Button_Apply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Apply.Click
        Apply_Changes()
        Me.TabControl1.Focus()
        Me.Button_Apply.Enabled = False
    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Button_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_OK.Click
        Apply_Changes()
        Me.Close()
    End Sub

    Private Sub Button_BrowseContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_BrowseContacts.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Folder = Me.Application.GetNamespace("MAPI").PickFolder

            If Folder IsNot Nothing Then

                Me.ContactFolder = Folder
                Me.TextBox_ContactFolder.Text = CStr(Folder.Name)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Folder = Nothing

        End Try

    End Sub

    Private Sub Button_BrowseMessages_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_BrowseMessages.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Folder = Me.Application.GetNamespace("MAPI").PickFolder

            If Folder IsNot Nothing Then

                Me.MessageFolder = Folder
                Me.TextBox_MessageFolder.Text = CStr(Folder.Name)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Folder = Nothing

        End Try

    End Sub

    Private Sub Apply_Changes()

        Try
            With My.Settings
                .DefaultContactFolderId = Me.ContactFolder.EntryID
                .DefaultMessageFolderId = Me.MessageFolder.EntryID
                .SendMessagesInGroups = Me.CheckBox_SendInGroups.Checked
                .MessageGroupSize = CInt(Me.NumericUpDown_GroupSize.Value)
                .MessageDelay = CInt(Me.NumericUpDown_Delay.Value)
                .Save()
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        End Try

    End Sub

    Private Sub CheckBox_SendInGroups_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_SendInGroups.CheckedChanged

        Me.NumericUpDown_GroupSize.Enabled = Me.CheckBox_SendInGroups.Checked
        Me.NumericUpDown_Delay.Enabled = Me.CheckBox_SendInGroups.Checked

        Me.Button_Apply.Enabled = True

    End Sub

    Private Sub NumericUpDown_GroupSize_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown_GroupSize.ValueChanged
        Me.Button_Apply.Enabled = True
    End Sub

    Private Sub NumericUpDown_Delay_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown_Delay.ValueChanged
        Me.Button_Apply.Enabled = True
    End Sub

End Class