Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ExportEventArgs_NET.ExportEventArgs")> Public Class ExportEventArgs
	Implements _IEventArgs
	
	
	Public Destination As String
	'UPGRADE_NOTE: Folder was upgraded to Folder_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Folder_Renamed As Microsoft.Office.Interop.Outlook.MAPIFolder
	Public Recurse As Boolean
	
	'--------------------------------------------------------------------------------
	'IEventArgs Methods
	'--------------------------------------------------------------------------------
	Private ReadOnly Property IEventArgs_ToString() As String Implements _IEventArgs.ToString_Renamed
		Get
			IEventArgs_ToString = "ExportEventArgs"
		End Get
	End Property
End Class