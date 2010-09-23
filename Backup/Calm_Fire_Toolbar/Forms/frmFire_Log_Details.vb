' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Log_Details

	Private Sub TreeView1_DblClick() Handles TreeView1.DblClick
		Dim filename As String = ""
		If TreeView1.SelectedItem.Text <> "" And TreeView1.SelectedItem.Text Like "Filename:*" Then
			filename = Replace(TreeView1.SelectedItem.Text, "Filename: ", "")
		Else
			Exit Sub
		End If
		Dim retval As Double = 27
		
		Dim pCursor As ESRI.ArcGIS.Framework.IMouseCursor = New ESRI.ArcGIS.Framework.MouseCursor()
		
		' cursor will change here.
		pCursor.SetCursor(2) ' hourglass
		
		If Fileexists(filename) = True Then
			modCommon_Functions.ExecAss(Me.hWnd, filename)
		End If
		
	End Sub

End Class
