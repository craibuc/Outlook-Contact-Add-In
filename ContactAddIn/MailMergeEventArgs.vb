Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("MailMergeEventArgs_NET.MailMergeEventArgs")> Public Class MailMergeEventArgs
	Implements _IEventArgs
	
	
	'UPGRADE_NOTE: Folder was upgraded to Folder_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Folder_Renamed As Microsoft.Office.Interop.Outlook.MAPIFolder
	Public MessageTemplate As Microsoft.Office.Interop.Outlook.MailItem
	Public SendInGroups As Boolean
	Public GroupSize As Short
	Public MessageDelay As Short
	Public Recurse As Boolean
	
	'--------------------------------------------------------------------------------
	'IEventArgs Methods
	'--------------------------------------------------------------------------------
	Private ReadOnly Property IEventArgs_ToString() As String Implements _IEventArgs.ToString_Renamed
		Get
			IEventArgs_ToString = "MailMergeEventArgs"
		End Get
	End Property
End Class