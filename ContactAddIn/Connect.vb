Imports Extensibility
Imports System.Runtime.InteropServices
Imports Office = Microsoft.Office.Core
Imports Outlook = Microsoft.Office.Interop.Outlook

<GuidAttribute("3D7BED27-C25D-4FD9-BAF6-1B45E1AFC9AF"), ProgIdAttribute("ContactAddIn.Connect")> _
Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _Application As Outlook.Application
    Private _AddIn As Office.COMAddIn

    Private WithEvents Button_ContactAddIn As Office.CommandBarPopup
    Private WithEvents Button_MailMerge As Office.CommandBarButton
    Private WithEvents Button_vCard As Office.CommandBarButton
    Private WithEvents Button_ContactExtractor As Office.CommandBarButton
    Private WithEvents Button_About As Office.CommandBarButton
    Private WithEvents Button_Options As Office.CommandBarButton

    Private form_MailMerge As Form_MailMerge
    Private form_Export As Form_Export
    Private form_Extract As Form_Extract
    Private form_About As Form_About
    Private form_Options As Form_Options

#Region "Public Properties"

    Public ReadOnly Property Application() As Outlook.Application
        Get
            Return _Application
        End Get
    End Property

    Public ReadOnly Property AddIn() As Office.COMAddIn
        Get
            Return _AddIn
        End Get
    End Property

#End Region

#Region "Private Methods"

    Private Sub Button_MailMerge_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles Button_MailMerge.Click

        form_MailMerge = New Form_MailMerge
        With form_MailMerge
            .Initialize(Me.Application)
            .Show()
        End With

    End Sub

    Private Sub Button_vCard_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles Button_vCard.Click

        form_Export = New Form_Export
        With form_Export
            .Initialize(Me.Application)
            .Show()
        End With

    End Sub

    Private Sub Button_ContactExtractor_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles Button_ContactExtractor.Click

        form_Extract = New Form_Extract
        With form_Extract
            .Initialize(Me.Application)
            .Show()
        End With

    End Sub

    Private Sub Button_About_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles Button_About.Click

        form_About = New Form_About
        With form_About
            .Show()
        End With

    End Sub

    Private Sub Button_Options_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles Button_Options.Click

        form_Options = New Form_Options()
        With form_Options
            .Initialize(Me.Application)
            .Show()
        End With

    End Sub

#End Region

#Region "IDTExtensibility2"

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        System.Windows.Forms.Application.EnableVisualStyles()

        _Application = CType(application, Outlook.Application)
        _AddIn = CType(addInInst, Office.COMAddIn)

        ' If you aren't in startup, manually call OnStartupComplete.
        If (connectMode <> Extensibility.ext_ConnectMode.ext_cm_Startup) Then _
           Call OnStartupComplete(custom)

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete

        Dim MenuBar As Office.CommandBar = Me.Application.ActiveExplorer.CommandBars.ActiveMenuBar
        Dim ToolsMenu As Office.CommandBarPopup = CType(MenuBar.Controls(5), Office.CommandBarPopup)

        'add menu items
        Button_ContactAddIn = CType(ToolsMenu.Controls.Add(Office.MsoControlType.msoControlPopup, , , , True), Office.CommandBarPopup)
        With Button_ContactAddIn
            .BeginGroup = True
            .Caption = "Contact AddIn"
            .Visible = True
        End With

        Button_MailMerge = CType(Button_ContactAddIn.Controls.Add(Office.MsoControlType.msoControlButton, , , , True), Office.CommandBarButton)
        With Button_MailMerge
            'AddHandler Button_MailMerge.Click, AddressOf Button_MailMerge_Click
            .Caption = "eMailMerge Wizard"
            .Style = Office.MsoButtonStyle.msoButtonCaption
            .Visible = True
        End With

        Button_vCard = CType(Button_ContactAddIn.Controls.Add(Office.MsoControlType.msoControlButton, , , , True), Office.CommandBarButton)
        With Button_vCard
            .Caption = "Export to vCard"
            .Style = Office.MsoButtonStyle.msoButtonCaption
            .Visible = True
        End With

        Button_ContactExtractor = CType(Button_ContactAddIn.Controls.Add(Office.MsoControlType.msoControlButton, , , , True), Office.CommandBarButton)
        With Button_ContactExtractor
            .Caption = "Extract Contacts"
            .Style = Office.MsoButtonStyle.msoButtonCaption
            .Visible = True
        End With

        Button_Options = CType(Button_ContactAddIn.Controls.Add(Office.MsoControlType.msoControlButton, , , , True), Office.CommandBarButton)
        With Button_Options
            .BeginGroup = True
            .Caption = "Options..."
            .Style = Office.MsoButtonStyle.msoButtonCaption
            .Visible = True
        End With

        Button_About = CType(Button_ContactAddIn.Controls.Add(Office.MsoControlType.msoControlButton, , , , True), Office.CommandBarButton)
        With Button_About
            .BeginGroup = True
            .Caption = "About Contact Add-In"
            .Style = Office.MsoButtonStyle.msoButtonCaption
            .Visible = True
        End With

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
        On Error Resume Next

        If RemoveMode <> Extensibility.ext_DisconnectMode.ext_dm_HostShutdown Then Call OnBeginShutdown(custom)

        _AddIn = Nothing
        _Application = Nothing

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
        On Error Resume Next

        'remove menu items
        Button_MailMerge.Delete()
        Button_MailMerge = Nothing

        Button_vCard.Delete()
        Button_vCard = Nothing

        Button_ContactExtractor.Delete()
        Button_ContactExtractor = Nothing

        Button_About.Delete()
        Button_About = Nothing

        Button_Options.Delete()
        Button_Options = Nothing

        Button_ContactAddIn.Delete()
        Button_ContactAddIn = Nothing

    End Sub

#End Region

End Class
