' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Logon

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: Logon Form. A Logon is required to be able to create or open an incident
	'Called By: clsFire_Logon
	'Calls: None
	'Accepts: None
	'Returns: None
	'******************************************************************************************************************
	
	Private Sub btnCancel_Click() Handles btnCancel.Click
		
		Unload6(Me)
		
	End Sub

	Private Sub btnLogon_Click() Handles btnLogon.Click
		
		If Me.txtUser.Text = "" Then
			MsgBox6("Please Enter a User Name", MsgBoxStyle.Exclamation, "User Name Required")
			Me.txtUser.SetFocus()
			Exit Sub
		End If
		
		modGlobals.g_Username = Me.txtUser.Text
		
		If Me.cbDept.Text = "DEC" Then
			modGlobals.g_UserDept = g_Fire_Dept.DEC
		ElseIf Me.cbDept.Text = "FESA" Then
			modGlobals.g_UserDept = g_Fire_Dept.FESA
		End If
		
		Unload6(Me)
		
	End Sub

End Class
