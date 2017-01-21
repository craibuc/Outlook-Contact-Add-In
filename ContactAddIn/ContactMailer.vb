Imports Outlook = Microsoft.Office.Interop.Outlook

<System.Runtime.InteropServices.ProgId("ContactAddIn.ContactMailer")> _
Public Class ContactMailer

    Private _Application As Outlook.Application
    Private _SendInGroups As Boolean = False
    Private _GroupSize As Integer = 50
    Private _MessageDelay As Integer = 5
    Private _Recurse As Boolean = False
    Private _MessageTemplate As Outlook.MailItem

    Public Event Processed(ByVal sender As Object, ByVal e As MailMergeEventArgs)

#Region "Public Properties"

    Public ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    ''' <summary>
    ''' how many email messages to send in bulk
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SendInGroups() As Boolean
        Get
            Return _SendInGroups
        End Get
        Set(ByVal value As Boolean)
            _SendInGroups = value
        End Set
    End Property

    ''' <summary>
    ''' # messages to send as a group
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GroupSize() As Integer
        Get
            Return _GroupSize
        End Get
        Set(ByVal value As Integer)
            _GroupSize = value
        End Set
    End Property

    ''' <summary>
    ''' delay between groups
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MessageDelay() As Integer
        Get
            Return _MessageDelay
        End Get
        Set(ByVal value As Integer)
            _MessageDelay = value
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

    ''' <summary>
    ''' template MailItem
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MessageTemplate() As Outlook.MailItem
        Get
            Return _MessageTemplate
        End Get
        Set(ByVal value As Outlook.MailItem)
            _MessageTemplate = value
        End Set
    End Property

#End Region

    Public Sub New(ByVal instance As Outlook.Application)
        _Application = instance
    End Sub

#Region "Public Methods"

    Public Function GetDraftBySubject(ByRef subject As String) As Outlook.MailItem

        Dim draftFolder As Outlook.MAPIFolder
        Dim draft As Outlook.MailItem

        Try

            'get root folder
            draftFolder = Me.Application.GetNamespace("MAPI").GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderDrafts)

            'get draft
            'todo: convert to Restrict()
            draft = CType(draftFolder.Items.Find(String.Concat("[Subject] = '", subject, "'")), Outlook.MailItem)

            Return draft

        Catch ex As Exception
            Throw ex

        Finally
            draftFolder = Nothing
            draft = Nothing

        End Try

    End Function

    Public Sub OpenTemplate(ByRef path As String)

        Try

            Me.MessageTemplate = CType(Me.Application.CreateItemFromTemplate(path), Outlook.MailItem)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")

        End Try

    End Sub

    Public Sub Process(ByRef Folder As Outlook.MAPIFolder, ByVal Filter As String)

        'raise an error
        If Folder Is Nothing Then Exit Sub
        If Me.MessageTemplate Is Nothing Then Exit Sub

        Static Counter As Integer
        Static DeferUntil As Date

        Try

            'Filter = String.Concat(Filter, " AND NOT ([Email1Address] Is Nothing)")
            Filter = String.Concat(Filter, " AND [Email1Address]<>", Chr(34), Chr(34))
            'Filter = String.Concat("NOT([Email1Address] IS NULL)")
            'Filter = "[Email1Address]<>" & Chr(34) & Chr(34)

            For Each Item As Object In Folder.Items.Restrict(Filter)
                If TypeOf Item Is Outlook.ContactItem Then

                    'cast item as contact
                    Dim contactItem As Outlook.ContactItem = CType(Item, Outlook.ContactItem)

                    If Not (contactItem.Email1Address Is Nothing And contactItem.Email2Address Is Nothing And contactItem.Email3Address Is Nothing) Then

                        'create a new message from the template
                        Dim mailItem As Outlook.MailItem = CType(Me.MessageTemplate.Copy, Outlook.MailItem)
                        With mailItem
                            'add email address
                            .Recipients.Add(contactItem.Email1Address)
                            .Recipients.ResolveAll()

                            'substitute variables
                            Dim ct As New ContactTokenCollection(contactItem)
                            .Body = ct.Replace(.Body)

                        End With

                        If SendInGroups Then

                            'increment # of messages sent
                            Counter += 1

                            'if # of messages sent is at limit
                            If (Counter - 1) Mod GroupSize = 0 Then
                                DeferUntil = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, (Counter - 1) / GroupSize * MessageDelay, Now)
                            End If

                            mailItem.DeferredDeliveryTime = DeferUntil

                        End If

                        'send
                        mailItem.Send()
                        'mailItem = Nothing

                    End If

                    contactItem = Nothing

                End If

            Next Item

            Dim Subfolder As Microsoft.Office.Interop.Outlook.MAPIFolder
            If Recurse Then
                For Each Subfolder In Folder.Folders
                    Process(Subfolder, Filter)
                Next Subfolder
            End If

            RaiseEvent Processed(Me, New MailMergeEventArgs(Counter))

        Catch ex As Exception
            Throw ex

        Finally

        End Try

    End Sub

    Public Sub SendMessage(ByRef contact As Outlook.ContactItem)

        If (contact.Email1Address Is Nothing And contact.Email2Address Is Nothing And contact.Email3Address Is Nothing) Then Exit Sub

        Static Counter As Integer

        Try

            'create a new message from the template
            Dim mailItem As Outlook.MailItem = CType(Me.MessageTemplate.Copy, Outlook.MailItem)
            With mailItem

                'add email address
                .Recipients.Add(contact.Email1Address)
                .Recipients.ResolveAll()

                'substitute variables
                Dim ct As New ContactTokenCollection(contact)
                .Body = ct.Replace(.Body)

            End With

            Dim Sent As Date = Now

            If Me.SendInGroups Then

                'increment # of messages sent
                Counter += 1


                'if # of messages sent is at limit
                If (Counter - 1) Mod Me.GroupSize = 0 Then
                    Sent = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, (Counter - 1) / Me.GroupSize * Me.MessageDelay, Now)
                End If

                mailItem.DeferredDeliveryTime = Sent

            End If

            With contact
                If .UserProperties("Last Contacted") Is Nothing Then
                    .UserProperties.Add("Last Contacted", Outlook.OlUserPropertyType.olDateTime)
                End If
                .UserProperties("Last Contacted").Value = Sent
                .Save()
            End With

            'send
            mailItem.Send()
            mailItem = Nothing

        Catch ex As Exception
            Throw

        End Try

    End Sub

#End Region

End Class

Public Class MailMergeEventArgs
    Inherits EventArgs

    Public MessagesSent As Integer

    Public Sub New()
    End Sub
    Public Sub New(ByVal messagesSent As Integer)
        Me.MessagesSent = messagesSent
    End Sub

End Class