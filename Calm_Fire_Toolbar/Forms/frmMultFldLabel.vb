' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmMultFldLabel

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: FIMT
	'Description
	'Called By
	'Calls: None
	'Accepts
	'Returns: None
	'******************************************************************************************************************
	
	' Events
	Public Event UpdateFieldNames(ByRef sCurrentLayerName As String)
	Public Event LabelFeatures(ByRef sCurrentLayerName As String)
	
	' Constants
	Private Const c_sModuleFileName As String = "Incident Tools\frmMultFldLabel.frm"
	
	Private Sub cmbLayers_Click() Handles cmbLayers.Click
		On Error GoTo ErrorHandler
		
13:
		RaiseEvent UpdateFieldNames(ByVal6(Me.cmbLayers.Text))
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmbLayers_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdApply_Click() Handles cmdApply.Click
		On Error GoTo ErrorHandler
		
23:
		RaiseEvent LabelFeatures(ByVal6(Me.cmbLayers.Text))
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdCancel_Click() Handles cmdCancel.Click
		On Error GoTo ErrorHandler
		
33:
		Me.Hide()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdOK_Click() Handles cmdOK.Click
		On Error GoTo ErrorHandler
		
43:
		Me.Hide()
44:
		RaiseEvent LabelFeatures(ByVal6(Me.cmbLayers.Text))
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		Exit Sub
ErrorHandler:
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

End Class
