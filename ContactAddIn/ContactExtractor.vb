Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

<System.Runtime.InteropServices.ProgId("ContactAddIn.ContactExtractor")> _
Public Class ContactExtractor

    Private _Application As Outlook.Application
    Private _Destination As Outlook.MAPIFolder
    Private _Recurse As Boolean = False

    Public Delegate Sub ProcessHandler(ByVal sender As Object, ByVal e As EventArgs)
    Public Event Processed(ByVal sender As Object, ByVal e As ExtractEventArgs)

#Region "Public Properties"

    Public ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' folder where Contacts are saved.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Destination() As Outlook.MAPIFolder
        Get
            Return _Destination
        End Get
        Set(ByVal value As Outlook.MAPIFolder)
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

    Public Sub Process(ByVal mailFolder As Outlook.MAPIFolder)

        'quality control
        If mailFolder Is Nothing Or Me.Destination Is Nothing Then Exit Sub

        Dim counter As Integer = 0

        Try

            For Each Item As Object In mailFolder.Items
                If TypeOf Item Is Outlook.MailItem Then

                    Dim mailItem As Outlook.MailItem = CType(Item, Outlook.MailItem)

                    If Me.Destination.Items.Restrict(String.Format("[Email1Address]='{0}' OR [Email2Address]='{0}' OR [Email3Address]='{0}'", mailItem.SenderEmailAddress)).Count = 0 Then

                        'Dim contactItem As Outlook.MailItem = mailItem.Copy
                        Dim contactItem As Outlook.ContactItem = Me.Application.CreateItem(Outlook.OlItemType.olContactItem)
                        With contactItem
                            .FullName = mailItem.SenderName

                            .Body = mailItem.Body

                            'For Each file As Outlook.Attachment In mailItem.Attachments
                            '    .Attachments.Add(file, Type.Missing, Type.Missing, Type.Missing)
                            'Next

                            .Email1Address = mailItem.SenderEmailAddress
                            .Email1AddressType = mailItem.SenderEmailType
                            .Email1DisplayName = mailItem.SenderName

                            If .UserProperties.Item("Send Mail") Is Nothing Then
                                .UserProperties.Add("Send Mail", Outlook.OlUserPropertyType.olYesNo, False)
                            End If
                            .UserProperties.Item("Send Mail").Value = True

                            'If .UserProperties.Item("Folder Property") Is Nothing Then
                            '    .UserProperties.Add("Folder Property", Outlook.OlUserPropertyType.olText, True)
                            '    .UserProperties.Item("Folder Property").Value = "my folder property"
                            'End If

                            .Save()
                            .Move(Me.Destination)

                        End With
                        contactItem = Nothing

                        counter += 1

                    End If

                    mailItem = Nothing

                End If
            Next Item

            Dim Subfolder As Outlook.MAPIFolder
            If Recurse Then
                For Each Subfolder In mailFolder.Folders
                    Process(Subfolder)
                Next Subfolder
            End If

            RaiseEvent Processed(Me, New ExtractEventArgs(counter))

        Catch ex As Exception
            Throw ex

        Finally           
            GC.Collect()

        End Try

    End Sub

#End Region

End Class

Public Class ExtractEventArgs
    Inherits EventArgs

    Public ContactsExtracted As Integer

    Public Sub New()
    End Sub
    Public Sub New(ByVal contactsExtracted As Integer)
        Me.ContactsExtracted = contactsExtracted
    End Sub

End Class
