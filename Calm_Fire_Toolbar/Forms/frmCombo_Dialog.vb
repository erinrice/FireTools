' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmCombo_Dialog

	Public didok As Boolean
	Public listvalue As String = ""
	Public frmCanceled As Boolean
	
	Private Sub Cancel_bn_Click() Handles Cancel_bn.Click
		frmCanceled = True
		Unload6(Me)
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		didok = False
	End Sub

	Private Sub OK_bn_Click() Handles OK_bn.Click
		frmCanceled = False
		listvalue = Me.Values_cb.List(Values_cb.ListIndex)
		didok = True
		Unload6(Me)
	End Sub

End Class
