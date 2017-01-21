Imports System.Windows.Forms

Public Class Form_Progress

    Public Event Cancelled(ByVal sender As Object, ByVal e As EventArgs)

    Public Sub ReportProgress(ByVal percentProgress As Integer)
        Me.ProgressBar1.Value = percentProgress
    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click

        RaiseEvent Cancelled(Me, New EventArgs)

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()

    End Sub

End Class
