Imports Outlook = Microsoft.Office.Interop.Outlook

<System.Runtime.InteropServices.ProgId("ContactAddIn.ContactExporter")> _
Public Class ContactExporter

    Private _Application As Outlook.Application
    Private _Recurse As Boolean = False
    Private _Destination As String = String.Empty

    Public Event Processed(ByVal sender As Object, ByVal e As ExportEventArgs)

#Region "Public Properties"

    Public ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' folder where vCards are saved.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Destination() As String
        Get
            Return _Destination
        End Get
        Set(ByVal value As String)
            _Destination = value
        End Set
    End Property

    ''' <summary>
    ''' process subfolders
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Recurse() As Boolean
        Get
            Return _Recurse
        End Get
        Set(ByVal value As Boolean)
            _Recurse = value
        End Set
    End Property

#End Region

    Public Sub New(ByVal instance As Outlook.Application)
        _Application = instance
    End Sub

#Region "Public Methods"

    Public Sub Process(ByRef source As Outlook.MAPIFolder)

        'quality control
        If source Is Nothing Then Exit Sub
        If Me.Destination = String.Empty Then Exit Sub

        'todo: this may not be declared correctly
        Dim counter As Integer = 0

        Try

            For Each Item As Object In source.Items
                If TypeOf Item Is Outlook.ContactItem Then

                    Dim contactItem As Outlook.ContactItem = CType(Item, Outlook.ContactItem)

                    'calculate path
                    Dim path As String = Me.Destination & "\" & Right(source.FolderPath, Len(source.FolderPath) - InStr(3, source.FolderPath, "\"))

                    'create folder
                    Dim di As New System.IO.DirectoryInfo(path)
                    With di
                        If Not .Exists Then .Create()
                    End With

                    'save as vCard
                    contactItem.SaveAs(path & "\" & contactItem.Subject & ".vcf", Outlook.OlSaveAsType.olVCard)

                    counter += 1

                    contactItem = Nothing

                End If
            Next Item

            Dim Subfolder As Outlook.MAPIFolder
            If Recurse Then
                For Each Subfolder In source.Folders
                    Process(Subfolder)
                Next Subfolder
            End If

            'todo: this may raise too many events
            RaiseEvent Processed(Me, New ExportEventArgs(counter))

        Catch ex As Exception
            Throw ex

        Finally

        End Try

    End Sub

#End Region

End Class

Public Class ExportEventArgs
    Inherits EventArgs

    Public ContactsExported As Integer

    Public Sub New()
    End Sub
    Public Sub New(ByVal contactsExported As Integer)
        Me.ContactsExported = contactsExported
    End Sub

End Class