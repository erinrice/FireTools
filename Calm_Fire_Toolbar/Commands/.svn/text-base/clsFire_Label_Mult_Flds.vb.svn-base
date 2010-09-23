' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Label_Mult_Flds")> _ 
<VB6Object("CALM_FireTools.clsFire_Label_Mult_Flds")> _
Public Class clsFire_Label_Mult_Flds
	Implements IDisposable
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Label_Mult_Flds' member isn't used anywhere in current application.

	'  Interfaces implemented by this command
	
	'  ArcMap objects referenced by this command
	Private m_pApp As ESRI.ArcGIS.Framework.IApplication
	
	'  Member variables controlled by this class
	Private mBitmap As Image

    'saeid
    '    Private WithEvents frmMultFields As New frmMultFldLabel

    Private WithEvents frmMultFields As frmMultFldLabel

    '  Constants
	' UPGRADE_INFO (#0701): The 'BITMAP_NAME' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Const BITMAP_NAME As String = "MULTFLDLABEL"
	' UPGRADE_INFO (#0561): The 'TOOL_NAME' symbol was defined without an explicit "As" clause.
	Private Const TOOL_NAME As String = "Label Multiple Fields"
	' UPGRADE_INFO (#0561): The 'TOOL_MSG' symbol was defined without an explicit "As" clause.
	Private Const TOOL_MSG As String = "Create ""stacked"" labels from multiple attribute fields."
	Private Const c_sModuleFileName As String = "Incident Tools\MultFldLabelCmd.cls"
	
	Private p_polyfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_linefeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_pointfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_divfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	
	Private Sub Class_Initialize_VB6()
		On Error Resume Next
		
22:
		'Set mBitmap = LoadResPicture(BITMAP_NAME, vbResBitmap)
		
	End Sub

	Private Sub Class_Terminate_VB6()
		On Error Resume Next
		
29:
		mBitmap = Nothing
30:
		m_pApp = Nothing
31:
		frmMultFields = Nothing
		
	End Sub

	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
		Get
			On Error Resume Next
			
38:
			Return GetPictureHandle6(frmImages.multilabel.Picture)
			
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error Resume Next
			
45:
			Return "Label Mulitple Fields"
			
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
		Get
			On Error Resume Next
			
52:
			'ICommand_Category = TOOL_CATEGORY
			
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Dim bEnabled As Boolean = False
			
			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
			bEnabled = Incident_Workspace_There()
			If g_Username = "" Then
				bEnabled = False
			End If
			
			Return bEnabled
			
ErrorHandler:
			HandleError(True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
		Get
			' NOT USED
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
		Get
			' NOT USED
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
		Get
			' NOT USED
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error Resume Next
			
88:
			Return TOOL_MSG
			
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
		Get
			On Error Resume Next
			
95:
			Return TOOL_NAME
			
	 	End Get
	End Property

	Private Sub frmMultFields_LabelFeatures(ByRef sCurrentLayerName As String) Handles frmMultFields.LabelFeatures
		On Error GoTo ErrorHandler
		
		' UPGRADE_INFO (#0701): The 'LayerKeys' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim LayerKeys() As Object
		Dim bGotFeatLayer As Boolean
		' UPGRADE_INFO (#0701): The 'iLayerCounter' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim iLayerCounter As Short
		' UPGRADE_INFO (#0701): The 'pField' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pField As ESRI.ArcGIS.Geodatabase.IField
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, modGlobals.g_FIRE_EXTENSION)
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim pIncidentInfo As clsFire_Incident_Info = pExtension
		Dim pMap As ESRI.ArcGIS.Carto.IMap

		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		pMap = thisdocument.FocusMap
		Dim count As Short
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = sCurrentLayerName Then
						bGotFeatLayer = True
						pFeatLayer = pcomplayer.Layer(complayercount)
						Exit For
					End If
				Next
			End If
		Next
		
		' Make sure that there was a match
		If bGotFeatLayer = False Then Exit Sub
		
		' Get AnnotateLayerPropertiesCollection from the selected layer.
		Dim pAnnoLayerPropsColl As ESRI.ArcGIS.Carto.IAnnotateLayerPropertiesCollection = pFeatLayer.AnnotationProperties
		
		' Get the (first) property set in the collection.
		Dim pAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerProperties
		pAnnoLayerPropsColl.QueryItem(0, pAnnoLayerProps, Nothing, Nothing)
		
		pAnnoLayerProps.LabelWhichFeatures = ESRI.ArcGIS.Carto.esriLabelWhichFeatures.esriVisibleFeatures
		
		' Update LabelEngineLayerProperties.
		Dim pLabelEngineLayerProps As ESRI.ArcGIS.Carto.ILabelEngineLayerProperties = pAnnoLayerProps
		
		' Set up the font, Times New Roman, 11pt, Bold.
		Dim pTextSymbol As ESRI.ArcGIS.Display.ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol()
		Dim pFontDisp As Font = New Font("Arial", 8)

		With pFontDisp
			FontChangeName6(pFontDisp, "Tahoma")
			FontChangeSize6(pFontDisp, 11)
			FontChangeBold6(pFontDisp, True)
		End With
		
		FontChangeName6(pTextSymbol.Font, pFontDisp.Name)
		Dim pMask As ESRI.ArcGIS.Display.IMask = pTextSymbol
		pMask.MaskStyle = ESRI.ArcGIS.Display.esriMaskStyle.esriMSHalo
		pMask.MaskSize = 2
		
		pLabelEngineLayerProps.Symbol = pTextSymbol
		
		' Set the expression and turn on the labels.
		Dim iListCount As Short
		Dim sLabelExpression As String = ""
		Dim bGotFirstSelected As Boolean = False
		Dim sFieldName As String = ""
		Dim fields_string As String = ""

		For iListCount = 0 To frmMultFields.lstFieldNames.ListCount - 1
			If frmMultFields.lstFieldNames.Selected(iListCount) Then
				If bGotFirstSelected Then
					
					sFieldName = GetFieldName(pFeatLayer, ByVal6(frmMultFields.lstFieldNames.List(CInt(iListCount))))
					
					If sFieldName = "" Then
						sFieldName = frmMultFields.lstFieldNames.List(CInt(iListCount))
					End If
					
					If frmMultFields.chkIncludeFieldNames.Value Then
						' UPGRADE_INFO (#0571): String concatenation inside a loop. Consider declaring the 'sLabelExpression' variable as a StringBuilder6 object.
						sLabelExpression = sLabelExpression & " & vbcrlf & " & Chr6(34) & frmMultFields.lstFieldNames.List(CInt(iListCount)) & Chr6(34) & Chr6(32) & "&" & Chr6(34) & ": " & Chr6(34) & Chr6(32) & "&" & "[" & sFieldName & "]"
					Else
						sLabelExpression = sLabelExpression & " & vbcrlf & " & "[" & sFieldName & "]"
					End If
				Else
					sFieldName = GetFieldName(pFeatLayer, ByVal6(frmMultFields.lstFieldNames.List(CInt(iListCount))))
					
					If sFieldName = "" Then
						sFieldName = frmMultFields.lstFieldNames.List(CInt(iListCount))
					End If
					
					If frmMultFields.chkIncludeFieldNames.Value Then
						sLabelExpression = Chr6(34) & frmMultFields.lstFieldNames.List(CInt(iListCount)) & Chr6(34) & Chr6(32) & "&" & Chr6(34) & ": " & Chr6(34) & Chr6(32) & "&" & "[" & sFieldName & "]"
						bGotFirstSelected = True
					Else
						sLabelExpression = "[" & sFieldName & "]"
						bGotFirstSelected = True
					End If
				End If 'Applies to: If bGotFirstSelected Then
				If fields_string = "" Then
					fields_string = "[" & sFieldName & "]"
				Else
					' UPGRADE_INFO (#0571): String concatenation inside a loop. Consider declaring the 'fields_string' variable as a StringBuilder6 object.
					fields_string = fields_string & ", [" & sFieldName & "]"
				End If
			End If 'Applies to: If frmMultFields.lstFieldNames.Selected(iListCount) Then
		Next
		
		sLabelExpression = "Function FindLabel (" & fields_string & ", [Eff_To])" & ControlChars.CrLf & "If [Eff_To]  = " & Chr6(34) & "Active" & Chr6(34) & " then" & ControlChars.CrLf & "FindLabel = " & sLabelExpression & ControlChars.CrLf & "End If" & ControlChars.CrLf & "End Function"
		
		pLabelEngineLayerProps.IsExpressionSimple = False
		pLabelEngineLayerProps.Expression = sLabelExpression
		
		pFeatLayer.DisplayAnnotation = True
		
		' Refresh the view.
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pActView As ESRI.ArcGIS.Carto.IActiveView = pMxDoc.ActiveView
		Dim pDocDirty As ESRI.ArcGIS.Framework.IDocumentDirty = pMxDoc

		pDocDirty.SetDirty()
		pActView.Refresh()
		
		pActView = Nothing
		pMxDoc = Nothing
		
		Exit Sub
ErrorHandler:
		HandleError(True, "frmMultFields_LabelFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'pFeatLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'sAlias' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Function GetFieldName(ByRef pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef sAlias As String) As String
		On Error GoTo ErrorHandler
		
		' Given an IFeatureLayer and a field Alias return the associatiated field Name.
		
		Dim lField As Integer
		Dim pField As ESRI.ArcGIS.Geodatabase.IField
		Dim pFields As ESRI.ArcGIS.Geodatabase.IFields = pFeatLayer.FeatureClass.Fields

		lField = pFields.FindFieldByAliasName(sAlias)
		If lField <> -1 Then
			pField = pFields.Field(lField)
			'Return the field name.
			Return pField.Name
		Else
			Return ""
		End If
		
		Exit Function
ErrorHandler:
		HandleError(False, "GetFieldName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	Private Sub frmMultFields_UpdateFieldNames(ByRef sCurrentLayerName As String) Handles frmMultFields.UpdateFieldNames
		On Error GoTo ErrorHandler
		
		' UPGRADE_INFO (#0701): The 'LayerKeys' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim LayerKeys() As Object
		' UPGRADE_INFO (#0701): The 'iLayerCounter' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim iLayerCounter As Short
		' UPGRADE_INFO (#0701): The 'pField' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pField As ESRI.ArcGIS.Geodatabase.IField
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, modGlobals.g_FIRE_EXTENSION)
		' UPGRADE_INFO (#0701): The 'pLayerDict' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pLayerDict As VB6Dictionary
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim pIncidentInfo As clsFire_Incident_Info = pExtension
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		' Make sure that layers are from the current FIMT layer.
		
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		pMap = thisdocument.FocusMap
		Dim count As Short
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = sCurrentLayerName Then
						pFeatLayer = pcomplayer.Layer(complayercount)
						Exit For
					End If
				Next
			End If
		Next
		
		' Label the layer combo box with the first layer name.
		frmMultFields.cmbLayers.Text = sCurrentLayerName
		
		' Populate the list box with the fields associated with this layer.
		PopulateFields(pFeatLayer)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "frmMultFields_UpdateFieldNames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
		
		' UPGRADE_INFO (#0701): The 'LayerKeys' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim LayerKeys() As Object
		' UPGRADE_INFO (#0701): The 'iLayerCounter' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim iLayerCounter As Short
		Dim bGotFirstLayer As Boolean = False
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
		' UPGRADE_INFO (#0701): The 'pLayerDict' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pLayerDict As VB6Dictionary
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim pIncidentInfo As clsFire_Incident_Info
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		pMap = thisdocument.FocusMap
		
		' Make sure that layers are from the current FIMT layer.
		pExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		pIncidentInfo = pExtension
		frmMultFields = New frmMultFldLabel()

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_BOUNDARY_NAME Then
						p_polyfeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_LINE_NAME Then
						p_linefeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_POINT_NAME Then
						p_pointfeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_ASSIGNMENT_NAME Then
						p_divfeatlayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		frmMultFields.cmbLayers.AddItem(p_polyfeatlayer.Name)
		PopulateFields(p_polyfeatlayer)
		
		frmMultFields.cmbLayers.AddItem(p_linefeatlayer.Name)
		frmMultFields.cmbLayers.Text = p_linefeatlayer.Name
		PopulateFields(p_linefeatlayer)
		
		frmMultFields.cmbLayers.AddItem(p_pointfeatlayer.Name)
		frmMultFields.cmbLayers.Text = p_pointfeatlayer.Name
		PopulateFields(p_pointfeatlayer)
		
		frmMultFields.cmbLayers.AddItem(p_divfeatlayer.Name)
		frmMultFields.cmbLayers.Text = p_divfeatlayer.Name
		PopulateFields(p_divfeatlayer)
		
		' Show the form.
        '		frmMultFields.Show(VBRUN.FormShowConstants.vbModal)
        frmMultFields.ShowDialog()
		frmMultFields = Nothing
		
		Exit Sub
ErrorHandler:
		HandleError(False, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'pFeatLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub PopulateFields(ByRef pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer)
		On Error GoTo ErrorHandler
		
		' Populate the list box with the fields associated with this layer.
		Dim pField As ESRI.ArcGIS.Geodatabase.IField
		Dim iFieldCounter As Short
		
		frmMultFields.lstFieldNames.Clear()
		
		For iFieldCounter = 0 To pFeatLayer.FeatureClass.Fields.FieldCount - 1
			pField = pFeatLayer.FeatureClass.Fields.Field(iFieldCounter)
			If pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeDate Or pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeString Or pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeDouble Or pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeSingle Or pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeInteger Or pField.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeSmallInteger Then
				
				If pField.AliasName = "" Then
					frmMultFields.lstFieldNames.AddItem((pField.Name))
				Else
					frmMultFields.lstFieldNames.AddItem((pField.AliasName))
				End If
			End If
		Next
		
		' Refresh
		frmMultFields.lstFieldNames.Refresh()
		
		Exit Sub
ErrorHandler:
		HandleError(False, "PopulateMultLblForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
		On Error GoTo ErrorHandler
		
437:
		m_pApp = hook
		
		Exit Sub
ErrorHandler:
		HandleError(False, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
447:
			Return "Label Fire Layers with multiple fields"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	 	End Get
	End Property

	#Region "Constructor"
	
	'A public default constructor
	Public Sub New()
		Class_Initialize_VB6()
		' Add initialization code here
	End Sub
	
	#End Region

	#Region "Static constructor"
	
	' This static constructor ensures that the VB6 support library be initialized before using current class.
	Shared Sub New()
		EnsureVB6LibraryInitialization()
	End Sub
	
	#End Region
	
	Protected Overrides Sub Finalize()
		Dispose(False)
	End Sub
	
	Public Sub Dispose() Implements System.IDisposable.Dispose
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
	
	Protected Overridable Sub Dispose(ByVal disposing As Boolean)
				Class_Terminate_VB6()
	End Sub

End Class
