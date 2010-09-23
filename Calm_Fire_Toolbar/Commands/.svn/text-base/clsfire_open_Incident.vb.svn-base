' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsfire_open_Incident")> _ 
<VB6Object("CALM_FireTools.clsfire_open_Incident")> _
Public Class clsfire_open_Incident
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsfire_open_Incident' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
			On Error GoTo ErrorHandler
			
			Dim bEnabled As Boolean = False
			
			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
			If g_Username = "" Then
				bEnabled = False
			End If
			
			Return bEnabled
			
ErrorHandler:
			HandleError(True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			' ICommand_Checked =
			
			Exit Property
ErrorHandler:
			HandleError(True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Open Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Open Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Open a Fire Incident Geodatabase"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Open an existing Fire Incident Geodatabase"
			
ErrorHandler:
			HandleError(True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			' ICommand_HelpFile =
			
			Exit Property
ErrorHandler:
			HandleError(True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			' ICommand_HelpContextID =
			
			Exit Property
ErrorHandler:
			HandleError(True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return GetPictureHandle6(frmImages.Open.Picture)
			
ErrorHandler:
			HandleError(True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			' ICommand_Category =
			
			Exit Property
ErrorHandler:
			HandleError(True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
		On Error GoTo ErrorHandler
		
		' TODO: Add your implementation here

		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
		
		'check if an incident has already been opened
		If modGlobals.g_INCIDENT_LAYER_NAME <> "" Then
			MsgBox6("An Incident has already been created using this document, please start a new Map Document before creating or opening another Incident", MsgBoxStyle.Exclamation, "Incident Already Open")
			Exit Sub
		End If
		
		Dim layerdialogresult As String = modCommon_Functions.OpenGeodatabase(m_pApp.hWnd)
		
		If layerdialogresult = "empty" Then
			Exit Sub
		End If
		
		Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.AccessWorkspaceFactory()
		
		Dim pworkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(layerdialogresult, m_pApp.hWnd)
		Dim pWorkspace2 As ESRI.ArcGIS.Geodatabase.IWorkspace2 = pWorkspaceFactory.OpenFromFile(layerdialogresult, m_pApp.hWnd)
		
		If pWorkspace2.NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, "Fire_Boundary") = False And pWorkspace2.NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, "Fire_point") = False And pWorkspace2.NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, "Fire_Line") = False Then
			MsgBox6(layerdialogresult & " is not in the correct Fire Incident Format", MsgBoxStyle.Exclamation, "Fire Incident")
			Exit Sub
		End If
		
		' UPGRADE_INFO (#0701): The 'pFeatClass' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		' Now create appropriate Feature Classes
		' UPGRADE_INFO (#0701): The 'progress' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim progress As Integer
		Dim pointfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pworkspace.OpenFeatureClass(g_FIRE_POINT)
		Dim polyfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pworkspace.OpenFeatureClass(g_FIRE_BOUNDARY)
		Dim Linefeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pworkspace.OpenFeatureClass(g_FIRE_LINE)
		Dim assfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pworkspace.OpenFeatureClass(g_FIRE_ASSIGNMENT)
		Dim annofeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pworkspace.OpenFeatureClass(g_FIRE_ANNO)
		
		'load the featuerclasses into the Map
		' Add feature class as layer to the map
		Dim polyFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayer()
		Dim pointFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayer()
		Dim lineFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayer()
		Dim assFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayer()
		Dim annoFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FDOGraphicsLayer()
		Dim annocomplayer As ESRI.ArcGIS.Carto.ICompositeLayer2 = annoFeatureLayer
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDoc.FocusMap
		
		polyFeatureLayer.FeatureClass = polyfeatclass
		polyFeatureLayer.Name = g_FIRE_BOUNDARY_NAME
		
		lineFeatureLayer.FeatureClass = Linefeatclass
		lineFeatureLayer.Name = g_FIRE_LINE_NAME
		
		pointFeatureLayer.FeatureClass = pointfeatclass
		pointFeatureLayer.Name = g_FIRE_POINT_NAME
		
		assFeatureLayer.FeatureClass = assfeatclass
		assFeatureLayer.Name = g_FIRE_ASSIGNMENT_NAME
		
		annoFeatureLayer.FeatureClass = annofeatclass
		annoFeatureLayer.Name = g_FIRE_ANNO_NAME
		
		'Set Incident Table properties
		Dim inc_number As String = modCommon_Functions.GetFireInfoProp(pworkspace, "INCIDENT_NUMBER")
		Dim inc_region As String = modCommon_Functions.GetFireInfoProp(pworkspace, "INCIDENT_REGION")
		Dim inc_region_ext As String = modCommon_Functions.GetFireInfoProp(pworkspace, "INCIDENT_REGION_EXT")
		
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		myIncidentInfo.IncidentNumber = inc_number
		myIncidentInfo.incidentRegion = inc_region
		myIncidentInfo.IncidentRegionEx = inc_region_ext
		myIncidentInfo.IncidentWorkspace = pworkspace
		
		'create a new grouplayer
		Dim pgrouplayer As ESRI.ArcGIS.Carto.IGroupLayer = New ESRI.ArcGIS.Carto.GroupLayer()
		pgrouplayer.Name = inc_region_ext & "::" & inc_number
		
		modGlobals.g_INCIDENT_LAYER_NAME = inc_region_ext & "::" & inc_number
		
		pMap.AddLayer(pgrouplayer)
		
		pgrouplayer.Add(pointFeatureLayer)
		pgrouplayer.Add(assFeatureLayer)
		pgrouplayer.Add(lineFeatureLayer)
		pgrouplayer.Add(polyFeatureLayer)
		pgrouplayer.Add(annoFeatureLayer)
		
		pgrouplayer.Expanded = True
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
			Exit Sub
		End If
		
		modCommon_Functions.Apply_Layer_File_Symbology(polyFeatureLayer, Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Boundary.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(lineFeatureLayer, Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Line.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(pointFeatureLayer, Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Point.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(assFeatureLayer, Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Assignment.lyr")
		
		modCommon_Functions.expand_layer(polyFeatureLayer, False)
		modCommon_Functions.expand_layer(pointFeatureLayer, False)
		modCommon_Functions.expand_layer(lineFeatureLayer, False)
		modCommon_Functions.expand_layer(assFeatureLayer, False)
		annocomplayer.Expanded = False
		
		modFire_Tools.Update_Log_Window()
		
        pMxDoc.UpdateContents()

        ' it make sures that user has created or opened the incident in capturing part of the application
        g_incidentStatus = True
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	#Region "Constructor"
	
	'A public default constructor
	Public Sub New()
		' Add initialization code here
	End Sub
	
	#End Region

	#Region "Static constructor"
	
	' This static constructor ensures that the VB6 support library be initialized before using current class.
	Shared Sub New()
		EnsureVB6LibraryInitialization()
	End Sub
	
	#End Region

End Class
