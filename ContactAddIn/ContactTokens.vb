Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection

Public Class ContactTokenCollection
    Inherits System.Collections.Specialized.NameValueCollection

    Private _Item As Outlook.ContactItem
    Private _contactItemType As Type = GetType(Outlook.ContactItem)

    Public Sub New(ByVal item As Outlook.ContactItem)

        _Item = item

        Me.Add("%FirstName%", "FirstName")
        Me.Add("%LastName%", "LastName")
        Me.Add("%Title%", "Title")  'Mr., Mrs., ...
        Me.Add("%WebPage%", "WebPage")  'Mr., Mrs., .

    End Sub

    Public Function Replace(ByVal value As String) As String

        Try

            For Each Key As String In Me.Keys

                Dim propName As String = Me.Item(Key)
                'Dim pi As PropertyInfo = CType(_contactItemType.GetProperty(propName), PropertyInfo)
                'Dim propValue As String = pi.GetValue(_Item, Nothing).ToString
                Dim propValue As String = CallByName(_Item, propName, CallType.Get)
                value = value.Replace(Key, propValue)

            Next

            Return value

        Catch ex As Exception
            Throw ex

        End Try

    End Function

End Class
