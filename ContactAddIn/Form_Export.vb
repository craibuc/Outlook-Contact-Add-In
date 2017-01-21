Option Strict Off
Option Explicit On

Imports Outlook = Microsoft.Office.Interop.Outlook

Friend Class Form_Export
	Inherits System.Windows.Forms.Form

    Private _Application As Outlook.Application
    Private _ContactFolder As Outlook.MAPIFolder

#Region "Protected Properties"

    Protected ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' Source MAPI folder that contains the Contacts to be exported to vCards.
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

#End Region

    Public Sub Initialize(ByRef application As Outlook.Application)

        Try

            _Application = application

            'set default contact folder
            Me.ContactFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
            Me.Text_Source.Text = CStr(Me.ContactFolder.Name)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        End Try

    End Sub

    Private Sub Command_Browse_Destination_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Browse_Destination.Click

        With Me.FolderBrowserDialog1
            .RootFolder = Environment.SpecialFolder.Desktop
            .Description = "Select a Folder"
            .ShowNewFolderButton = True
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.Text_Destination.Text = .SelectedPath
            End If
        End With

    End Sub

    Private Sub Command_Browse_Source_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Browse_Source.Click

        Dim Folder As Outlook.MAPIFolder

        Try

            Folder = Me.Application.GetNamespace("MAPI").PickFolder
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

    Private Sub CommandButton_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommandButton_Cancel.Click

        Me.Close()

    End Sub

    Private Sub CommandButton_OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommandButton_OK.Click

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Try

            Dim CE As New ContactExporter(Me.Application)
            With CE
                .Destination = Me.Text_Destination.Text
                .Recurse = False
                AddHandler .Processed, AddressOf ProcessingCompleted
                .Process(Me.ContactFolder)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Contact Add-In")

        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default

        End Try

    End Sub

    Private Sub ProcessingCompleted(ByVal sender As System.Object, ByVal e As ExportEventArgs)

        MsgBox(String.Format("Processing Completed.  {0} vCards were created.", e.ContactsExported), MsgBoxStyle.Information, "vCard Exporter")
        Me.Close()

    End Sub

End Class