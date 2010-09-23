' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Point

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This form specifies parameters to attribute Fire Points
	'Called By: clsedit_task_fire_point
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

	Public Sub populatePointCB()
		'Populate the combobox of point feature types
		Me.Feature_type_cb.Clear()
		Me.Feature_type_cb.AddItem(("2WD"))
		Me.Feature_type_cb.AddItem(("4WD"))
		Me.Feature_type_cb.AddItem(("Airstrip"))
		Me.Feature_type_cb.AddItem(("Assembly Area"))
		Me.Feature_type_cb.AddItem(("Base Camp"))
		Me.Feature_type_cb.AddItem(("Building"))
		Me.Feature_type_cb.AddItem(("Control Centre"))
		Me.Feature_type_cb.AddItem(("DEC Tower"))
		Me.Feature_type_cb.AddItem(("Divisional Point"))
		Me.Feature_type_cb.AddItem(("Dozer"))
		Me.Feature_type_cb.AddItem(("Dwelling"))
		Me.Feature_type_cb.AddItem(("Escape Route"))
		Me.Feature_type_cb.AddItem(("FESA Tower"))
		Me.Feature_type_cb.AddItem(("Fire Origin"))
		Me.Feature_type_cb.AddItem(("First Aid"))
		Me.Feature_type_cb.AddItem(("Fixed Wing"))
		Me.Feature_type_cb.AddItem(("Float"))
		Me.Feature_type_cb.AddItem(("Gang Truck"))
		Me.Feature_type_cb.AddItem(("Heavy Duty"))
		Me.Feature_type_cb.AddItem(("Helibase"))
		Me.Feature_type_cb.AddItem(("Hot Spot"))
		Me.Feature_type_cb.AddItem(("Information"))
		Me.Feature_type_cb.AddItem(("Light Unit"))
		Me.Feature_type_cb.AddItem(("Loader"))
		Me.Feature_type_cb.AddItem(("OTHER Tower"))
		Me.Feature_type_cb.AddItem(("Operations Point"))
		Me.Feature_type_cb.AddItem(("Ramp"))
		Me.Feature_type_cb.AddItem(("Refuge Site"))
		Me.Feature_type_cb.AddItem(("Rotary Wing"))
		Me.Feature_type_cb.AddItem(("Spot Fire"))
		Me.Feature_type_cb.AddItem(("Staging Area"))
		Me.Feature_type_cb.AddItem(("Ute"))
		Me.Feature_type_cb.AddItem(("Vehicle Control Point"))
		Me.Feature_type_cb.AddItem(("Water Point"))

		Me.Feature_type_cb.ListIndex = 0
	End Sub

End Class
