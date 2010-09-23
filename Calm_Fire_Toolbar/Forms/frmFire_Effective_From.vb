' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Effective_From

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This form allows the user to specify whether to append the new shape to
	'             the existing active shapes, or specify the new shape as the new active shape
	'Called By: Create Fire Shape Classes
	'Calls: None
	'Accepts: None
	'Returns: AppendtoCurrent
	'******************************************************************************************************************
	Public AppendToCurrent As Boolean
	Public OKButtonClicked As Boolean 'added 29/9/06 to check which button is clicked
	
	Private Sub btnCancel_Click() Handles btnCancel.Click
		OKButtonClicked = False
		Me.Hide()
	End Sub

	Private Sub btnOK_Click() Handles btnOK.Click
		OKButtonClicked = True
		'changed 29-08-2007 by FMS request - GrantB
		If ckbAppend.Value = 0 Then
			If MsgBox6("Existing features will be made inactive and no longer be displayed." & ControlChars.NewLine & "To retain existing feature click Cancel and Enable 'Add new shapes to existing features'", MsgBoxStyle.OkCancel, "DEC Fire Toolbar") = MsgBoxResult.Cancel Then
				Exit Sub
			End If
		End If
		Me.Hide()
	End Sub

	Private Sub ckbAppend_Click() Handles ckbAppend.Click
		If ckbAppend.Value = 1 Then
			Me.txtDate.Locked = True
			Me.txtDate.BackColor = SystemColors.ScrollBar 
			AppendToCurrent = True
		Else
			Me.txtDate.Locked = False
			Me.txtDate.BackColor = SystemColors.Window 
			AppendToCurrent = False
		End If
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		If ckbAppend.Value = 1 Then
			AppendToCurrent = True
		End If
	End Sub

End Class
