Option Strict Off
Option Explicit On

Imports Outlook = Microsoft.Office.Interop.Outlook

Friend Class Form_Extract
    Inherits System.Windows.Forms.Form

    Private _Application As Outlook.Application
    Private _ContactFolder As Outlook.MAPIFolder
    Private _MailFolder As Outlook.MAPIFolder

#Region "Protected Properties"

    Protected ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' Destination MAPI folder that contains the Contacts
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property ContactFolder() As Outlook.MAPIFolder
        Get
            Return _ContactFolder
        End Get
        Set(ByVal value As Outlook.MAPIFolder)
            If value.DefaultItemType <> Outlook.OlItemType.olContactItem Then _
                Throw New Exception("Choose a ContactItem folder")

            _ContactFolder = value
        End Set
    End Property

    ''' <summary>
    ''' Source MAPI folder that will be scanned for Contact information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Property MailFolder() As Outlook.MAPIFolder
        Get
            Return _MailFolder
        End Get
        Set(ByVal value As Outlook.MAPIFolder)
            If value.DefaultItemType <> Outlook.OlItemType.olMailItem Then _
                Throw New Exception("Choose a MailItem folder")

            _MailFolder = value
        End Set
    End Property

#End Region

    Public Sub Initialize(ByRef application As Outlook.Application)

        Try

            _Application = application

            'set source message folder
            Me.MailFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)
            Me.Text_Source.Text = CStr(Me.MailFolder.Name)

            'set default contact folder
            Me.ContactFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
            Me.Text_Destination.Text = CStr(Me.ContactFolder.Name)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        End Try

    End Sub

    Private Sub Command_Browse_Destination_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Browse_Destination.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Folder = Me.Application.GetNamespace("MAPI").PickFolder
            If Folder IsNot Nothing Then
                If Folder.DefaultItemType <> Outlook.OlItemType.olContactItem Then _
                    MsgBox("Choose a folder that supports Contacts.")

                Me.ContactFolder = Folder
                Me.Text_Destination.Text = CStr(Me.ContactFolder.Name)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Folder = Nothing

        End Try

    End Sub

    Private Sub Command_Browse_Source_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Browse_Source.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Folder = Me.Application.GetNamespace("MAPI").PickFolder
            If Folder IsNot Nothing Then

                'todo: fix this logic
                If Folder.DefaultItemType <> Outlook.OlItemType.olMailItem Then _
                    MsgBox("Choose a folder that supports Messages.")

                _MailFolder = Folder
                Me.Text_Source.Text = CStr(Me.MailFolder.Name)

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Folder = Nothing

        End Try

    End Sub

    Private Sub CommandButton_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommandButton_Cancel.Click

        Me.Close()

    End Sub

    Private Sub CommandButton_OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommandButton_OK.Click

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Try

            Dim CE As New ContactExtractor(Me.Application)
            With CE
                .Destination = Me.ContactFolder
                .Recurse = False
                AddHandler .Processed, AddressOf ProcessingCompleted
                .Process(Me.MailFolder)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default

        End Try

    End Sub

    Private Sub ProcessingCompleted(ByVal sender As System.Object, ByVal e As ExtractEventArgs)

        MsgBox(String.Format("Processing Completed.  {0} contacts were extracted.", e.ContactsExtracted), MsgBoxStyle.Information, "Contact Extractor")
        Me.Close()

    End Sub

End Class