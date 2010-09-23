' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Module modLayerSubtype

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: FIMT
	'Description
	'Called By
	'Calls
	'Accepts
	'Returns
	'******************************************************************************************************************
	
	' Variables used by the Error handler function - DO NOT REMOVE
	' UPGRADE_INFO (#0561): The 'c_ModuleFileName' symbol was defined without an explicit "As" clause.
	Private Const c_ModuleFileName As String = "modLayerSubtype.bas"
	
	' UPGRADE_INFO (#0551): The 'FeatureLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function LayerHasSubtypes(ByRef FeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef DefaultCode As Integer) As Boolean
		On Error GoTo ErrorHandler
		
		'  Returns a boolean indicating if the specified layer has
		'  subtypes.  If true, the default subtype value will be
		'  returned in the DefaultCode argument as well.
		
		Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pSubtypes As ESRI.ArcGIS.Geodatabase.ISubtypes
		Dim bHaveSubtypes As Boolean
		
16:
		DefaultCode = 0
17:
		bHaveSubtypes = False
		
19:
		pFeatureClass = FeatureLayer.FeatureClass
20:
		If TypeOf pFeatureClass Is ESRI.ArcGIS.Geodatabase.ISubtypes Then
21:
			pSubtypes = pFeatureClass
22:
			bHaveSubtypes = pSubtypes.HasSubtype
23:
			If bHaveSubtypes Then
24:
				DefaultCode = pSubtypes.DefaultSubtypeCode
25:
			End If
26:
		End If
		
28:
		Return bHaveSubtypes
		
ErrorHandler:
		HandleError(True, "LayerHasSubtypes " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Function


End Module
