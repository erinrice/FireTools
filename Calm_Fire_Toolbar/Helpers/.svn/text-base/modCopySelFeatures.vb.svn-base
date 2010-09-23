' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Module modCopySelFeatures

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: FIMT
	'Description
	'Called By
	'Calls: None
	'Accepts
	'Returns: None
	'******************************************************************************************************************
	
	' Variables used by the Error handler function - DO NOT REMOVE
	
	Private Const c_sModuleFileName As String = "Shared\CopySelFeatures.bas"
	
	' Variables used by the Error handler function - DO NOT REMOVE
	
	'   Function to copy selected features from the map
	'   into the target layer.
	
	'  Because this gets called as part of a larger operation when
	'  copying in perimeter layers, do not start and end an edit
	'  operation here - do it in whatever routine calls this one.
	
	' UPGRADE_INFO (#0551): The 'MXApp' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'TargetLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function CopySelFeatures(ByRef MXApp As ESRI.ArcGIS.ArcMapUI.IMxApplication, ByRef TargetLayer As ESRI.ArcGIS.Carto.IFeatureLayer) As ESRI.ArcGIS.esriSystem.ISet
		On Error GoTo ErrorHandler
		
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		' UPGRADE_INFO (#0701): The 'peditor' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pSelection As ESRI.ArcGIS.Carto.ISelection
		Dim pEnumFeature As ESRI.ArcGIS.Geodatabase.IEnumFeature
		Dim pSourceFeature As ESRI.ArcGIS.Geodatabase.IFeature
		Dim pSourceFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pTargetFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pTargetFeature As ESRI.ArcGIS.Geodatabase.IFeature
		Dim TargetShpField As Integer
		' UPGRADE_INFO (#0701): The 'pSubtypes' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pSubtypes As ESRI.ArcGIS.Geodatabase.ISubtypes
		Dim TargetHasSubtypes As Boolean
		Dim TargetSubtype As Integer
		Dim pGeometry As ESRI.ArcGIS.Geometry.IGeometry
		Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pTargetSpatialRef As ESRI.ArcGIS.Geometry.ISpatialReference
		Dim pTargetClone As ESRI.ArcGIS.esriSystem.IClone
		Dim pSourceGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pSourceSpatialRef As ESRI.ArcGIS.Geometry.ISpatialReference
		Dim pSourceClone As ESRI.ArcGIS.esriSystem.IClone
		Dim pRowsubtypes As ESRI.ArcGIS.Geodatabase.IRowSubtypes
		Dim pApp As ESRI.ArcGIS.Framework.IApplication
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatSet As ESRI.ArcGIS.esriSystem.ISet
		'  Variables used to show progress
		
		Dim pStepProg As ESRI.ArcGIS.esriSystem.IStepProgressor
43:
		pApp = MXApp
44:
		pStepProg = pApp.StatusBar.ProgressBar
		
		' Get the target feature class and determine its shape type,
		' whether it has subtypes, and the default subtype code.
		
49:
		pTargetFC = TargetLayer.FeatureClass
50:
		TargetShpField = pTargetFC.FindField(pTargetFC.ShapeFieldName)
51:
		TargetHasSubtypes = LayerHasSubtypes(TargetLayer, TargetSubtype)
		
53:
		pGeoDataset = pTargetFC
54:
		pTargetSpatialRef = pGeoDataset.SpatialReference
55:
		pTargetClone = pTargetSpatialRef
		
		'  Set up the progress bar for selected features.
		
59:
		pMap = pMxDoc.ActiveView ' GetFocusMap(MXApp)
60:
		With pStepProg
61:
			.MaxRange = pMap.SelectionCount
62:
			.MinRange = 0
63:
			.Position = 0
64:
			.Message = "Copying..."
65:
			.Show()
66:
		End With
67:
		DoEvents6()
68:
		pStepProg.Message = "Copying..."
		
		'  Start processing features...
		
72:
		pFeatSet = New ESRI.ArcGIS.esriSystem.Set()
73:
		pSelection = pMap.FeatureSelection
74:
		pEnumFeature = pSelection
75:
		'Set pEditor = GetEditor(MXApp)
76:
		''    pEditor.StartOperation
		
78:
		pEnumFeature.Reset()
79:
		pSourceFeature = pEnumFeature.Next()
80:
		Do While pSourceFeature IsNot Nothing
81:
			pSourceFC = pSourceFeature.Class
82:
			If pSourceFC.ShapeType = pTargetFC.ShapeType Then
83:
				pSourceFC = pSourceFeature.Class
84:
				pSourceGeoDataset = pSourceFC
85:
				pSourceSpatialRef = pSourceGeoDataset.SpatialReference
86:
				pSourceClone = pSourceSpatialRef
87:
				pGeometry = pSourceFeature.ShapeCopy
88:
				If Not pSourceClone.IsEqual(pTargetClone) Then
89:
					pGeometry.Project(pTargetSpatialRef)
90:
				End If
91:
				pTargetFeature = pTargetFC.CreateFeature()
92:
				pTargetFeature.Shape = pGeometry
93:
				pRowsubtypes = pTargetFeature
94:
				pRowsubtypes.InitDefaultValues()
				'
				' If there are matching fields in the source and target feature classes
				' copy any valid values from the source to the target.
				'
99:
				'Call CopyFeatureAttributes(pSourceFeature, pTargetFeature)
				'
				'  Most fields should have default values by now.
				'  Go through any remaining ones that are still null
				'  and set them to "" or 0 or whatever.  Most of these
				'  will be overwritten by values from the incoming feature,
				'  but we want to be sure not to leave any nulls.
				'
				'Dim lFldIndex As Long
				'Dim FieldType As esriFieldType
				
110:
				'For lFldIndex = 0 To pTargetFeature.Fields.FieldCount - 1
111:
				'If IsNull(pTargetFeature.Value(lFldIndex)) Then
112:
				'If pTargetFeature.Fields.Field(lFldIndex).Editable Then
				'
				'  We have a null value in an editable field.
				'  plug a default value in if we can.
				'
117:
				'FieldType = pTargetFeature.Fields.Field(lFldIndex).Type
				' Select Case FieldType
				'Case esriFieldTypeSmallInteger, esriFieldTypeInteger, esriFieldTypeSingle,
				                     ' 'esriFieldTypeDouble
121:
				'pTargetFeature.Value(lFldIndex) = 0
				'Case esriFieldTypeString
123:
				'pTargetFeature.Value(lFldIndex) = ""
				'Case esriFieldTypeDate
125:
				'pTargetFeature.Value(lFldIndex) = Date
126:
				'End Select
127:
				'End If         ' Field is editable
128:
				'End If           ' Field is null
129:
				'Next lFldIndex
				
				'  Store the feature and add it to the set we're going to return
				
133:
				'pTargetFeature.Store
134:
				pFeatSet.Add(pTargetFeature)
				
136:
			End If ' Applies to: If pSourceFC.ShapeType = pTargetFC.ShapeType
			'
			' Increment the progressor and get the next feature.
			'
140:
			If pStepProg.Position < pStepProg.MaxRange Then pStepProg.Step()
141:
			pSourceFeature = pEnumFeature.Next()
142:
		Loop
		
144:
		''    pEditor.StopOperation "Copy Features"
		
146:
		'pEditor.Map.ClearSelection
		
148:
		pApp = MXApp
149:
		pMxDoc = pApp.Document
150:
		pMxDoc.ActiveView.Refresh()
151:
		pStepProg.Hide()
		'
		' Return the ISet.
		'
155:
		Return pFeatSet
		
ErrorHandler:
		HandleError(False, "CopySelFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
160:
		''  pEditor.AbortOperation
	End Function


End Module
