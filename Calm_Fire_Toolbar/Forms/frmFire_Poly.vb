' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Poly

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This form specifies parameters to attribute Fire Poly's
	'Called By: clsedit_task_fire_poly
	'Calls: None
	'Accepts: None
	'Returns: None
	'******************************************************************************************************************
	
	Public Did_Cancel As Boolean
	
	Private Sub Cancel_btn_Click() Handles Cancel_btn.Click
		Did_Cancel = True
		Me.Hide()
	End Sub

	Private Sub OK_btn_Click() Handles OK_btn.Click
		Did_Cancel = False
		Me.Hide()
	End Sub

End Class
