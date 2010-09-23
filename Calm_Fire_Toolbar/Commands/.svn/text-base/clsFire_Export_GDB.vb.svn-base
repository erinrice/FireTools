' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Export_GDB")> _ 
<VB6Object("CALM_FireTools.clsFire_Export_GDB")> _
Public Class clsFire_Export_GDB
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Export_GDB' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	
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
			Return "xport_incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			'changed 28-08-2007 by FMS request - GrantB
			Return "Export Shapefiles" ' "Export Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Export Fire Incident Geodatabase to Shapefiles"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Export Fire Incident Geodatabase to Shapefiles"
			
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
			Return GetPictureHandle6(frmImages.exportgdb.Picture)
			
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
		' UPGRADE_INFO (#0701): The 'names1' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim names1() As String

		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
		
		exportGeoShapefile()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub exportGeoShapefile()
		
		' UPGRADE_INFO (#0701): The 'peditor' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim peditor As ESRI.ArcGIS.Editor.IEditor
		' UPGRADE_INFO (#0701): The 'pid' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		Dim p_polyfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim p_linefeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim p_pointfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim p_divfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim p_polyfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim p_linefeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim p_pointfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim p_divfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_BOUNDARY_NAME Then
						p_polyfeatlayer = pcomplayer.Layer(complayercount)
						p_polyfeatclass = p_polyfeatlayer.FeatureClass
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_LINE_NAME Then
						p_linefeatlayer = pcomplayer.Layer(complayercount)
						p_linefeatclass = p_linefeatlayer.FeatureClass
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_POINT_NAME Then
						p_pointfeatlayer = pcomplayer.Layer(complayercount)
						p_pointfeatclass = p_pointfeatlayer.FeatureClass
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_ASSIGNMENT_NAME Then
						p_divfeatlayer = pcomplayer.Layer(complayercount)
						p_divfeatclass = p_divfeatlayer.FeatureClass
					End If
				Next
			End If
		Next
		
		'Create a scratch workspace factory to use for the selection
		' UPGRADE_INFO (#0701): The 'pFeatClass' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pScratchWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pScratchWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IScratchWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.ScratchWorkspaceFactory()
		pScratchWorkspace = pScratchWorkspaceFactory.DefaultScratchWorkspace
		
		'  Get the output folder
		Dim mydate As Date = Now
		' UPGRADE_INFO (#0701): The 'myhourint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myhourint As Short
		' UPGRADE_INFO (#0701): The 'myminuteint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myminuteint As Short
		Dim mydatestring As String = ""
		Dim mytimestring As String = ""
		' UPGRADE_INFO (#0701): The 'MyMonthtext' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim MyMonthtext As Object = Nothing
		' UPGRADE_INFO (#0701): The 'MyDatetext' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim MyDatetext As Object = Nothing
		' UPGRADE_INFO (#0561): The 'myhour' symbol was defined without an explicit "As" clause.
		Dim myhour As Object = DatePart("h", mydate)
		' UPGRADE_INFO (#0561): The 'myminute' symbol was defined without an explicit "As" clause.
		Dim myminute As Object = Nothing
		' UPGRADE_INFO (#0561): The 'MyDay' symbol was defined without an explicit "As" clause.
		Dim MyDay As Object = Nothing
		' UPGRADE_INFO (#0561): The 'MyMonth' symbol was defined without an explicit "As" clause.
		Dim MyMonth As Object = Nothing
		Dim MyYear As String = ""
		If Len6(myhour) = 1 Then
			myhour = "0" & myhour
		End If
		
		myminute = DatePart("n", mydate)
		If Len6(myminute) = 1 Then
			myminute = "0" & myminute
		End If
		
		MyDay = DatePart("d", mydate)
		If Len6(MyDay) = 1 Then
			MyDay = "0" & MyDay
		End If
		
		MyMonth = DatePart("m", mydate)
		If Len6(MyMonth) = 1 Then
			MyMonth = "0" & MyMonth
		End If
		
		MyYear = DatePart("yyyy", mydate)
		mydatestring = MyYear & MyMonth & MyDay 'MyDay & MyMonth & MyYear ' GrantB 13-11-2007 by FMS request
		mytimestring = myhour & myminute
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		Dim filename_poly As String = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_" & "FireBoundary" & "_" & mydatestring & "_" & mytimestring
		Dim filename_line As String = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_" & "FireLine" & "_" & mydatestring & "_" & mytimestring
		Dim filename_point As String = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_" & "FirePoint" & "_" & mydatestring & "_" & mytimestring
		Dim filename_div As String = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_" & "FireAssignment" & "_" & mydatestring & "_" & mytimestring
		
		' GrantB 13-11-2007 by FMS request
		
		Dim pSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		pQueryFilter.WhereClause = "[OBJECTID] >-1"
		
		Dim Outpath As String = Trim(modCommon_Functions.GetOutputFolder())
		If Len6(Outpath) > 0 Then
			
			pSelSet = p_polyfeatclass.Select(pQueryFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
			ExportToShapeFile(p_polyfeatlayer, Outpath, filename_poly & ".shp", pSelSet)
			
			pSelSet = p_linefeatclass.Select(pQueryFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
			ExportToShapeFile(p_linefeatlayer, Outpath, filename_line & ".shp", pSelSet)
			
			pSelSet = p_pointfeatclass.Select(pQueryFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
			ExportToShapeFile(p_pointfeatlayer, Outpath, filename_point & ".shp", pSelSet)
			
			pSelSet = p_divfeatclass.Select(pQueryFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
			ExportToShapeFile(p_divfeatlayer, Outpath, filename_div & ".shp", pSelSet)
			
		End If

	End Sub

	' UPGRADE_INFO (#0701): The 'GetOutputFolder' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Function GetOutputFolder() As String
		' EXCLUDED: On Error GoTo ErrorHandler

		' EXCLUDED: Dim pGXDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
		' UPGRADE_INFO (#0701): The 'pObjectFilter' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pObjectFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		' EXCLUDED: Dim pEnumSelect As ESRI.ArcGIS.Catalog.IEnumGxObject
		' EXCLUDED: Dim pGXObject As ESRI.ArcGIS.Catalog.IGxObject

		' EXCLUDED: With pGXDialog
			' EXCLUDED: .RememberLocation = True
			' EXCLUDED: .AllowMultiSelect = False
			' EXCLUDED: .Title = "Destination Folder"
			' EXCLUDED: .ButtonCaption = "Open"
		' EXCLUDED: End With
		
		' EXCLUDED: If pGXDialog.DoModalOpen(0, pEnumSelect) Then
			' EXCLUDED: pGXObject = pEnumSelect.Next()
			' EXCLUDED: Return pGXObject.FullName
		' EXCLUDED: Else
			' EXCLUDED: Return ""
		' EXCLUDED: End If
		
		' EXCLUDED: Exit Function
		
		' EXCLUDED: ErrorHandler:
		' EXCLUDED: HandleError(False, "GetOutputFolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	' EXCLUDED: End Function
	
	' UPGRADE_INFO (#0551): The 'Kaartlaag' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Locatie' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Shapefilename' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'myselection' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ExportToShapeFile(ByRef Kaartlaag As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Locatie As String, ByRef Shapefilename As String, ByRef myselection As ESRI.ArcGIS.Geodatabase.ISelectionSet)
		
		' UPGRADE_INFO (#0701): The 'blnAllSelected' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim blnAllSelected As Boolean
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDoc.FocusMap
		Dim pBlockLayer As ESRI.ArcGIS.Carto.IFeatureLayer = Kaartlaag
		Dim pBlkTbl As ESRI.ArcGIS.Carto.IDisplayTable = pBlockLayer
		Dim pBlkClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pBlkTbl.DisplayTable
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset = pBlkClass
		Dim pInDSname As ESRI.ArcGIS.Geodatabase.IDatasetName = pDataset.FullName
		Dim pOutWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory()
		Dim pOutWS As ESRI.ArcGIS.Geodatabase.IWorkspace = pOutWSFact.OpenFromFile(Locatie, m_pApp.hWnd)
		Dim pDData As ESRI.ArcGIS.Geodatabase.IDataset = pOutWS
		Dim pFeatureWorkspacename As ESRI.ArcGIS.Geodatabase.IWorkspaceName = pDData.FullName
		Dim pOutDSname As ESRI.ArcGIS.Geodatabase.IDatasetName = New ESRI.ArcGIS.Geodatabase.FeatureClassName()
		Dim pExpOp As ESRI.ArcGIS.GeoDatabaseUI.IExportOperation = New ESRI.ArcGIS.GeoDatabaseUI.ExportOperation()

		'Dim pFeatureSelection As IFeatureSelection
		' UPGRADE_INFO (#0701): The 'pSelectionSet' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet

		pOutDSname.Name = Shapefilename
		
		pOutDSname.WorkspaceName = pFeatureWorkspacename
		
		pExpOp.ExportFeatureClass(pInDSname, Nothing, myselection, Nothing, pOutDSname, 0)

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
