' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports VB = Microsoft.VisualBasic

Friend Module modFire_Tools

	' UPGRADE_INFO (#0561): The 'ONLY_FEATURES_IN_EXTENT' symbol was defined without an explicit "As" clause.
	Private Const ONLY_FEATURES_IN_EXTENT As Boolean = True
	' change this to toggle between all features in data, or only those in extent
	' UPGRADE_INFO (#0561): The 'MAX_LAYER_COUNT' symbol was defined without an explicit "As" clause.
	Private Const MAX_LAYER_COUNT As Short = 100
	' increase this if you have more than 100 layers in your map
	'Private pOrigUVRend As IUniqueValueRenderer
	Public layers_added As New Collection
	
	' UPGRADE_INFO (#0551): The 'layername' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'SubType' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub SetCurrentLayer(ByRef layername As String, ByRef SubType As Integer)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Subroutine to set the current layer based on a name and subtype index
		'Called By
		'Calls: None
		'Accepts
		'Returns: None
		'******************************************************************************************************************
		
		Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim count As Short
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		pEditLayers = peditor
		
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		pMap = thisdocument.FocusMap
		
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim complayercount As Short
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			'Account for group layers
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To (pcomplayer.Count - 1) 'fixed 22/09/06 didn't have -1 was looking beyond extents of loop
					If pcomplayer.Layer(complayercount).Name = layername Then
						If pEditLayers.IsEditable(pcomplayer.Layer(complayercount)) Then
							pFeatLayer = pcomplayer.Layer(complayercount)
							pEditLayers.SetCurrentLayer(pFeatLayer, SubType)
							Exit Sub
						Else
							MsgBox6("This layer is not editable")
							Exit Sub
						End If
					End If
				Next
			ElseIf pMap.Layer(count).Name = layername Then
				'Make sure the layer is editable
				If pEditLayers.IsEditable(pMap.Layer(count)) Then
					pFeatLayer = pMap.Layer(count)
					pEditLayers.SetCurrentLayer(pFeatLayer, SubType)
					Exit Sub
				Else
					MsgBox6("This layer is not editable")
					Exit Sub
				End If
			End If
		Next
		
		'MsgBox "after loop", vbInformation, "Debug"
	End Sub

	' UPGRADE_INFO (#0551): The 'showReturn' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub UpdateMeasures(ByRef showReturn As Boolean)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Subroutine to update area fields for all of the Fire Layers
		'Called By
		'Calls: None
		'Accepts
		'Returns: None
		'******************************************************************************************************************
		
		On Error GoTo errhandler
		
		Dim polyFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim Linefeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pointfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim polyfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim Linefeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pointfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim complayercount As Short
		Dim count As Short
		Dim pPolyFeatCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		Dim pLineFeatCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		Dim pPolyFeature As ESRI.ArcGIS.Geodatabase.IFeature
		Dim pLineFeature As ESRI.ArcGIS.Geodatabase.IFeature
		
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		pMap = thisdocument.FocusMap
		
		'Reference the Fire layers
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = g_FIRE_BOUNDARY_NAME Then
						polyFeatLayer = pcomplayer.Layer(complayercount)
						polyfeatclass = polyFeatLayer.FeatureClass
					End If
					If pcomplayer.Layer(complayercount).Name = g_FIRE_LINE_NAME Then
						Linefeatlayer = pcomplayer.Layer(complayercount)
						Linefeatclass = Linefeatlayer.FeatureClass
					End If
					If pcomplayer.Layer(complayercount).Name = g_FIRE_POINT_NAME Then
						pointfeatlayer = pcomplayer.Layer(complayercount)
						pointfeatclass = pointfeatlayer.FeatureClass
					End If
				Next
			End If
		Next
		
		'loop through poly features
		'Dim polyfeatcount As Integer
		
		If polyfeatclass IsNot Nothing Then
			pPolyFeatCursor = polyfeatclass.Search(Nothing, False)
			pPolyFeature = pPolyFeatCursor.NextFeature()
			While pPolyFeature IsNot Nothing
				UpdateFeature_Poly(pPolyFeature)
				pPolyFeature = pPolyFeatCursor.NextFeature()
			End While
			'
			'For polyfeatcount = 1 To polyfeatclass.FeatureCount(Nothing)
			'    UpdateFeature_Poly polyfeatclass.GetFeature(polyfeatcount)
			'Next polyfeatcount
		End If
		
		'loop through Line features
		'Dim linefeatcount As Integer
		
		If Linefeatclass IsNot Nothing Then
			pLineFeatCursor = Linefeatclass.Search(Nothing, False)
			pLineFeature = pLineFeatCursor.NextFeature()
			While pLineFeature IsNot Nothing
				UpdateFeature_Line(pLineFeature)
				pLineFeature = pLineFeatCursor.NextFeature()
			End While
			'
			'For linefeatcount = 1 To Linefeatclass.FeatureCount(Nothing)
			'    UpdateFeature_Line Linefeatclass.GetFeature(linefeatcount)
			'Next linefeatcount
		End If
		
		If showReturn Then
			MsgBox6("Measurements updated successfully!", MsgBoxStyle.Information, "Success") 'GrantB 02-10-2007 - part of FMS Change Request
		End If
		
		Exit Sub
errhandler:
		MsgBox6(Err.Description)
		
	End Sub

	' UPGRADE_INFO (#0551): The 'pFeat' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub UpdateFeature_Poly(ByRef pFeat As ESRI.ArcGIS.Geodatabase.IFeature)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Create the Area values for features
		'Called By: Update_Measurements
		'Calls: Convert2Albers
		'Accepts: Features
		'Returns: None
		'******************************************************************************************************************
		
		On Error GoTo ErrorHandler
		
		Dim pArea As ESRI.ArcGIS.Geometry.IArea
		Dim pPoly As ESRI.ArcGIS.Geometry.IPolygon
		Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeat.Class
		Dim pspatref As ESRI.ArcGIS.Geometry.ISpatialReference
		
		pGeoDataset = pFeatClass
		pspatref = pGeoDataset.SpatialReference
		pPoly = pFeat.ShapeCopy
		pArea = pPoly
		
		Dim hafield As Short = pFeat.Fields.FindField(g_HECTARE_FIELD)
		Dim perimfield As Short = pFeat.Fields.FindField(g_PERIM_FIELD)

		' Project the shape into a localized Albers projection and
		' get its area and perimeter
		
		Dim pPolyAlbers As ESRI.ArcGIS.Geometry.IPolygon
		Dim pGeomAlbers As ESRI.ArcGIS.Geometry.IGeometry = ConvertToAlbers(pFeat.ShapeCopy)
		
		'Call ConverttoAlbers to calculate area's in albers projection
		pPolyAlbers = pGeomAlbers
		pArea = pPolyAlbers
		
		pFeat.Value(perimfield) = pPolyAlbers.Length
		pFeat.Value(hafield) = Abs6(pArea.Area / 10000)
		pFeat.Store()
		
		Exit Sub
		
ErrorHandler:
		'HandleError True, "UpdateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1
	End Sub
	
	' UPGRADE_INFO (#0551): The 'pFeat' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub UpdateFeature_Line(ByRef pFeat As ESRI.ArcGIS.Geodatabase.IFeature)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Create the Area values for features
		'Called By: Update_Measurements
		'Calls: Convert2Albers
		'Accepts: Features
		'Returns: None
		'******************************************************************************************************************
		
		On Error GoTo ErrorHandler
		
		' UPGRADE_INFO (#0701): The 'pArea' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pArea As ESRI.ArcGIS.Geometry.IArea
		Dim pPolyline As ESRI.ArcGIS.Geometry.IPolyline
		Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeat.Class
		Dim pspatref As ESRI.ArcGIS.Geometry.ISpatialReference
		
		pGeoDataset = pFeatClass
		pspatref = pGeoDataset.SpatialReference
		pPolyline = pFeat.ShapeCopy
		
		Dim hafield As Short = pFeat.Fields.FindField(g_HECTARE_FIELD)
		Dim perimfield As Short = pFeat.Fields.FindField(g_PERIM_FIELD)

		' Project the shape into a localized Albers projection and
		' get its area and perimeter
		Dim pPolylineAlbers As ESRI.ArcGIS.Geometry.IPolyline
		Dim pGeomAlbers As ESRI.ArcGIS.Geometry.IGeometry = ConvertToAlbers(pFeat.ShapeCopy)
		
		pPolylineAlbers = pGeomAlbers
		
		pFeat.Value(perimfield) = pPolylineAlbers.Length
		pFeat.Store()
		
		Exit Sub
		
ErrorHandler:
		'HandleError True, "UpdateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1
	End Sub
	
	' UPGRADE_INFO (#0551): The 'pgeom' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function ConvertToAlbers(ByRef pgeom As ESRI.ArcGIS.Geometry.IGeometry) As ESRI.ArcGIS.Geometry.IGeometry
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Project a Geometry into standard albers projection
		'Called By: UpdateFeature*
		'Calls: None
		'Accepts: IGeometry
		'Returns: Albers Projected Geometry
		'******************************************************************************************************************
		
		On Error GoTo ErrorHandler
		
		' UPGRADE_INFO (#0701): The 'pPoint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pPoint As ESRI.ArcGIS.Geometry.IPoint
		' UPGRADE_INFO (#0701): The 'pArea' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pArea As ESRI.ArcGIS.Geometry.IArea
		' UPGRADE_INFO (#0701): The 'spatialreftype' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim spatialreftype As String = ""
		' UPGRADE_INFO (#0701): The 'spatialref' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim spatialref As String = ""
		Dim spatialrefstring As String = ""
		Dim projlocation As String = ""
		' UPGRADE_INFO (#0701): The 'geolocation' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim geolocation As String = ""
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Your ARCGISHOME Environment Variable has not been set. Contact the GIS Section on 9334 0158 for help", MsgBoxStyle.Exclamation, "Error in Setup")
			Exit Function
		End If
		
		projlocation = Environ("ARCGISHOME") & "\Coordinate Systems\Projected Coordinate Systems\National Grids\Australia\"
		spatialrefstring = projlocation & "Albers_Equal_Conic_Area_GDA_Western_Australia.prj"
		
		Dim pSpatReffact As ESRI.ArcGIS.Geometry.ISpatialReferenceFactory2 = New ESRI.ArcGIS.Geometry.SpatialReferenceEnvironment()
		
		Dim pspatref As ESRI.ArcGIS.Geometry.ISpatialReference = pSpatReffact.CreateESRISpatialReferenceFromPRJFile(spatialrefstring)
		
		'Project the geometry to Albers
		pgeom.Project(pspatref)
		Return pgeom

ErrorHandler:
		'HandleError False, "ConvertToAlbers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1
	End Function
	
	' UPGRADE_INFO (#0551): The 'filename' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ExportActiveView_jpg(ByRef filename As String)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Modified by Nathan Eaton, CALM
		'Description: Export the ActiveView to a JPG
		'Called By: clsFire_Email_View
		'Calls
		'Accepts: JPG full filename
		'Returns: None
		'******************************************************************************************************************
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView = pMxDoc.FocusMap
		Dim pexport As ESRI.ArcGIS.Output.IExport = New ESRI.ArcGIS.Output.ExportJPEG()
		Dim pPixelBoundsEnv As ESRI.ArcGIS.Geometry.IEnvelope
		Dim exportRECT As ESRI.ArcGIS.Display.tagRECT
		Dim iOutputResolution As Short = 200
		Dim iScreenResolution As Short = 96
		Dim hDC As Integer

		'set properties to help jpeg resolution issue from using JpegExport interface
		modFire_Tools.SetOutputQuality(pActiveView, 1)
		
		pexport.ExportFileName = filename
		
		'Because we are exporting to a resolution that differs from screen resolution, we should
		' assign the two values to variables for use in our sizing calculations
		'default screen resolution is usually 96dpi
		pexport.Resolution = iOutputResolution

		'The ExportFrame property gives us the dimensions appropriate for an export at screen resolution.
		' Because we are exporting at a higher resolution (more pixels), we must multiply each dimesion
		' by the ratio of OutputResolution to ScreenResolution. Instead of assigning the entire
		' ExportFrame directly to the exportRECT, let's bring the values across one at a time and multiply
		' the dimensions.
		With exportRECT
			.left = 0
			.top = 0
			.right = pActiveView.ExportFrame.right * (iOutputResolution / iScreenResolution)
			.bottom = pActiveView.ExportFrame.bottom * (iOutputResolution / iScreenResolution)
		End With
		
		Dim exportenv As ESRI.ArcGIS.Geometry.IEnvelope = New ESRI.ArcGIS.Geometry.Envelope()
		exportenv = pActiveView.ScreenDisplay.DisplayTransformation.VisibleBounds
		
		Dim xpix As Double = (exportenv.XMax - exportenv.XMin) / (exportRECT.right - exportRECT.left)
		Dim ypix As Double = -(Abs6(exportenv.YMax) - Abs6(exportenv.YMin)) / (exportRECT.top - exportRECT.bottom)
		
		xpix = Format6(xpix, "0.000000000000000")
		ypix = Format6(ypix, "0.000000000000000")
		
		'create a jgw worldfile
		FileOpen6(100, Replace(filename, ".jpg", ".jgw"), OpenMode.Output, OpenAccess.Default, OpenShare.Default, -1)
		
		'Print instead of Write (Write adds quote marks)
		FilePrintLine6(100, xpix)
		FilePrintLine6(100, "0")
		FilePrintLine6(100, "0")
		FilePrintLine6(100, ypix)
		FilePrintLine6(100, exportenv.XMin)
		FilePrintLine6(100, exportenv.YMax)
		FileClose6(100)
		
		'Set up the PixelBounds envelope to match the exportRECT
		pPixelBoundsEnv = New ESRI.ArcGIS.Geometry.Envelope()
		pPixelBoundsEnv.PutCoords(exportRECT.left, exportRECT.top, exportRECT.right, exportRECT.bottom)
		pexport.PixelBounds = pPixelBoundsEnv
		pexport.StepProgressor = m_pApp.StatusBar.ProgressBar
		
		hDC = pexport.StartExporting()
		pActiveView.Output(hDC, pexport.Resolution, exportRECT, Nothing, Nothing)
		pexport.FinishExporting()
		pexport.Cleanup()
		
	End Sub
	
	' UPGRADE_INFO (#0551): The 'filename' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'MyRes' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'MyQual' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ExportLayout_jpg(ByRef filename As String, Optional ByRef MyRes As Double = 0, Optional ByRef MyQual As Double = 0)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Modified by Nathan Eaton, CALM
		'Description: Export the Layout to a JPG
		'Called By: clsFire_Export_Map, clsFire_Email_Layout
		'Calls
		'Accepts: JPG full filename, JPG Resolution as double
		'Returns: None
		'******************************************************************************************************************
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView = pMxDoc.PageLayout
		Dim pexport As ESRI.ArcGIS.Output.IExport = New ESRI.ArcGIS.Output.ExportJPEG()
		Dim pPixelBoundsEnv As ESRI.ArcGIS.Geometry.IEnvelope
		Dim exportRECT As ESRI.ArcGIS.Display.tagRECT
		Dim iOutputResolution As Short
		Dim iScreenResolution As Short = 96
		' UPGRADE_INFO (#0701): The 'iQuality' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim iQuality As Short
		Dim hDC As Integer

		'set properties to help jpeg resolution issue from using JpegExport interface
		modFire_Tools.SetOutputQuality(pActiveView, 1)
		
		pexport.ExportFileName = filename
		'Because we are exporting to a resolution that differs from screen resolution, we should
		' assign the two values to variables for use in our sizing calculations
		'default screen resolution is usually 96dpi
		
		'If a resolution is specified then use that instead of the default 200
		If MyRes > 0 Then
			iOutputResolution = MyRes
		Else
			iOutputResolution = 200
		End If
		pexport.Resolution = iOutputResolution
		
		'Specify Quality
		Dim pExportJPEG As ESRI.ArcGIS.Output.IExportJPEG
		If MyQual > 0 Then
			iOutputResolution = MyRes
			pExportJPEG = pexport '* QI to IExportJPEG
			pExportJPEG.Quality = MyQual '* set the quality
		End If
		
		'The ExportFrame property gives us the dimensions appropriate for an export at screen resolution.
		' Because we are exporting at a higher resolution (more pixels), we must multiply each dimesion
		' by the ratio of OutputResolution to ScreenResolution. Instead of assigning the entire
		' ExportFrame directly to the exportRECT, let's bring the values across one at a time and multiply
		' the dimensions.
		With exportRECT
			.left = 0
			.top = 0
			.right = pActiveView.ExportFrame.right * (iOutputResolution / iScreenResolution)
			.bottom = pActiveView.ExportFrame.bottom * (iOutputResolution / iScreenResolution)
		End With
		
		'Set up the PixelBounds envelope to match the exportRECT
		pPixelBoundsEnv = New ESRI.ArcGIS.Geometry.Envelope()
		pPixelBoundsEnv.PutCoords(exportRECT.left, exportRECT.top, exportRECT.right, exportRECT.bottom)
		pexport.PixelBounds = pPixelBoundsEnv
		pexport.StepProgressor = m_pApp.StatusBar.ProgressBar
		
		hDC = pexport.StartExporting()
		pActiveView.Output(hDC, pexport.Resolution, exportRECT, Nothing, Nothing)
		pexport.FinishExporting()
		pexport.Cleanup()
		
	End Sub


	Public Function Get_Fire_Line_Layer() As ESRI.ArcGIS.Carto.IFeatureLayer
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Access the Fire Line Layer
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: Fire Line Layer as FeatureLayer
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Line"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						pFeatLayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		Return pFeatLayer
		
	End Function

	Public Function Get_Fire_Poly_Layer() As ESRI.ArcGIS.Carto.IFeatureLayer
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Access the Fire Poly Layer
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: Fire Poly Layer as FeatureLayer
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Boundary"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						pFeatLayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		Return pFeatLayer
		
	End Function
	
	Public Function Get_Fire_Ass_Layer() As ESRI.ArcGIS.Carto.IFeatureLayer
		'******************************************************************************************************************
		'Date: 10 / 10 / 06
		'Author: Jack Green, DEC
		'Description: Access the Fire Assignment Layer
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: Fire Assignment Layer as FeatureLayer
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Assignment"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						pFeatLayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		Return pFeatLayer
		
	End Function
	
	Public Function Get_Fire_Point_Layer() As ESRI.ArcGIS.Carto.IFeatureLayer
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Access the Fire Point Layer
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: Fire Point Layer as FeatureLayer
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Point"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						pFeatLayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		Return pFeatLayer
		
	End Function
	
	Public Function Get_Fire_Anno_Layer() As ESRI.ArcGIS.Carto.IAnnotationLayer
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Access the Fire Anno Layer
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: Fire Anno Layer as FeatureLayer
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Annotation"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						pFeatLayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
		Return pFeatLayer
		
	End Function

    Public Function Fire_Poly_Exists() As Boolean
        '******************************************************************************************************************
        'Date: 06 / 06 / 06
        'Author: Nathan Eaton, CALM
        'Description: Check to see if the Fire_Poly Layer Exists in the current Focus Map
        'Called By: Many
        'Calls
        'Accepts: Nada
        'Returns: True/False Layer exists
        '******************************************************************************************************************

        Dim layername As String = "Fire Boundary"
        Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
        ' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
        Dim count As Short
        Dim complayercount As Short
        Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
        Fire_Poly_Exists = False

        'Loop through all of the maps layers to find the desired one
        For count = 0 To pMap.LayerCount - 1
            If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
                pcomplayer = pMap.Layer(count)
                For complayercount = 0 To pcomplayer.Count - 1
                    If pcomplayer.Layer(complayercount).Name = layername Then
                        Fire_Poly_Exists = True
                    End If
                Next
            End If
        Next

    End Function
	
	Public Function Fire_Line_Exists() As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if the Fire_Line Layer Exists in the current Focus Map
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True/False Layer exists
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Line"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Fire_Line_Exists = False
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						Fire_Line_Exists = True
					End If
				Next
			End If
		Next
		
	End Function
	
	Public Function Fire_Point_Exists() As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if the Fire_Point Layer Exists in the current Focus Map
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True/False Layer exists
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Point"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Fire_Point_Exists = False
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						Fire_Point_Exists = True
					End If
				Next
			End If
		Next
		
	End Function
	
	Public Function Fire_Div_Exists() As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if the Fire_Div Layer Exists in the current Focus Map
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True/False Layer exists
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Assignment"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Fire_Div_Exists = False
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						Fire_Div_Exists = True
					End If
				Next
			End If
		Next
		
	End Function
	
	Public Function Fire_Anno_Exists() As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if the Fire_Anno Layer Exists in the current Focus Map
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True/False Layer exists
		'******************************************************************************************************************
		
		Dim layername As String = "Fire Annotation"
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Fire_Anno_Exists = False
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = layername Then
						Fire_Anno_Exists = True
					End If
				Next
			End If
		Next
		
	End Function

	' UPGRADE_INFO (#0561): The 'Incident_StartEditing' symbol was defined without an explicit "As" clause.
	Public Function Incident_StartEditing() As Object
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Start Editing the Fire Incident
		'Called By: Many
		'Calls
		'Accepts
		'Returns
		'******************************************************************************************************************
		On Error GoTo errhandler
		
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView
		'Dim pSelectionSet As esriGeoDatabase.ISelectionSet
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		pMap = peditor.Map
		pActiveView = pMap
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		'Create an edit operation enabling undo/redo
		If peditor.EditState <> ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
			peditor.StartEditing(myIncidentInfo.IncidentWorkspace)
			'  Set pWorkspaceedit = myIncidentInfo.IncidentWorkspace
			'  pWorkspaceedit.StartEditing True
			'  pWorkspaceedit.StartEditOperation
		ElseIf peditor.EditWorkspace.PathName <> myIncidentInfo.IncidentWorkspace.PathName Then
			peditor.StartEditing(myIncidentInfo.IncidentWorkspace)
			'  Set pWorkspaceedit = myIncidentInfo.IncidentWorkspace
			'  pWorkspaceedit.StartEditing True
			'  pWorkspaceedit.StartEditOperation
		End If
		
		Exit Function
errhandler:
		MsgBox6(Err.Description)
	End Function

	' UPGRADE_INFO (#0561): The 'Incident_StopEditing' symbol was defined without an explicit "As" clause.
	Public Function Incident_StopEditing() As Object
		'******************************************************************************************************************
		'Date: 26 / 09 / 06
		'Author: Jack Green, DEC
		'Description: Stop Editing the Fire Incident and save any edits
		'Called By: Many
		'Calls
		'Accepts
		'Returns
		'******************************************************************************************************************
		On Error GoTo errhandler
		
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView
		' UPGRADE_INFO (#0701): The 'pSelectionSet' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		pMap = peditor.Map
		pActiveView = pMap
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		'Create an edit operation enabling undo/redo
		If peditor.EditState = ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
			
			'sp
			'    Set pWorkspaceedit = peditor.EditWorkspace
			'    pWorkspaceedit.StopEditOperation
			'    pWorkspaceedit.StopEditing True

			peditor.StopEditing(True)
		End If
		
		Exit Function
		
errhandler:
		MsgBox6(Err.Description)
	End Function

	Public Function Incident_Workspace_There() As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if the Incident Workspace has been instantiated
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True/False Workspace Exists exists
		'******************************************************************************************************************
		
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = Nothing
		'GrantB - 26-10-2007 - checking object since it has caused bugs
		pExt = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		If pExt Is Nothing Then
			Return False
		End If
		
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		If myIncidentInfo.IncidentWorkspace Is Nothing Then
			Return False
		Else
			Return True
		End If
		
	End Function

	' UPGRADE_INFO (#0701): The 'Show_Log_Window' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Public Function Show_Log_Window() As Object
		' EXCLUDED: '******************************************************************************************************************
		' EXCLUDED: 'Date: 06 / 06 / 06
		' EXCLUDED: 'Author: Nathan Eaton, CALM
		' EXCLUDED: 'Description: Not used in the project
		' EXCLUDED: 'Called By: Many
		' EXCLUDED: 'Calls
		' EXCLUDED: 'Accepts: Nada
		' EXCLUDED: 'Returns
		' EXCLUDED: '******************************************************************************************************************
		
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Clear()
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add(, , "R", "Root")
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add("R", MSComctlLib.TreeRelationshipConstants.tvwChild, "C1", "Child 1")
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add("R", MSComctlLib.TreeRelationshipConstants.tvwChild, "C2", "Child 2")
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add("R", MSComctlLib.TreeRelationshipConstants.tvwChild, "C3", "Child 3")
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add("C3", MSComctlLib.TreeRelationshipConstants.tvwChild, "C3a", "Child 3a")
		' EXCLUDED: frmFire_Log_Details.TreeView1.Nodes.Add("R", MSComctlLib.TreeRelationshipConstants.tvwChild, "C4", "Child 4")
		
	' EXCLUDED: End Function

	Public Sub Update_Log_Window()
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Update the contents of teh Fire Log Dockable Window
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns
		'******************************************************************************************************************
		
		frmFire_Log_Details.TreeView1.Nodes.Clear()
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		Dim pfeatworkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = myIncidentInfo.IncidentWorkspace
		
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable = pfeatworkspace.OpenTable("Fire_Log")
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
		
		Dim typefld As Short
		Dim titlefld As Short
		Dim commentsfld As Short
		Dim resfld As Short
		Dim scalefld As Short
		Dim sizefld As Short
		Dim orientfld As Short
		Dim idfld As Short
		Dim usernamefld As Short
		Dim deptfld As Short
		Dim datefld As Short
		Dim timefld As Short
		Dim filenamefld As Short
		Dim logtypefld As Short
		Dim emailtypefld As Short
		Dim y As Short
		If pTable.RowCount(Nothing) > 0 Then
			typefld = pTable.FindField("Log_Map_Type")
			titlefld = pTable.FindField("Log_Map_Title")
			commentsfld = pTable.FindField("Log_Comments")
			resfld = pTable.FindField("Log_Map_Resolution")
			scalefld = pTable.FindField("Log_Map_Scale")
			sizefld = pTable.FindField("Log_Page_Size")
			orientfld = pTable.FindField("Log_Page_Orientation")
			idfld = pTable.FindField("Log_Fire_ID")
			usernamefld = pTable.FindField("Log_Username")
			deptfld = pTable.FindField("Log_User_Dept")
			datefld = pTable.FindField("Log_Date")
			timefld = pTable.FindField("Log_Time")
			filenamefld = pTable.FindField("Log_Filename")
			logtypefld = pTable.FindField("Log_Type")
			emailtypefld = pTable.FindField("Log_Email_Type")
			
			For y = 1 To pTable.RowCount(Nothing) ' - 1
				
				pRow = pTable.GetRow(y)
				If GetDefaultMember6(pRow.Value(logtypefld)) = "Export Map" Then
					frmFire_Log_Details.TreeView1.Nodes.Add(, , "R" & y, GetDefaultMember6(pRow.Value(typefld)) & " " & GetDefaultMember6(pRow.Value(datefld)) & " " & GetDefaultMember6(pRow.Value(timefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C1_" & y, "Map Type: " & GetDefaultMember6(pRow.Value(typefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C2_" & y, "Map Title: " & GetDefaultMember6(pRow.Value(titlefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C3_" & y, "Comments: " & GetDefaultMember6(pRow.Value(commentsfld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C4_" & y, "Page")
					frmFire_Log_Details.TreeView1.Nodes.Add("C4_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C4_a" & y, "Page Size: " & GetDefaultMember6(pRow.Value(sizefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C4_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C4_b" & y, "Page Orientation: " & GetDefaultMember6(pRow.Value(orientfld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C4_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C4_c" & y, "Map Scale: " & GetDefaultMember6(pRow.Value(scalefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C5_" & y, "Image")
					frmFire_Log_Details.TreeView1.Nodes.Add("C5_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C5_a" & y, "Resolution: " & GetDefaultMember6(pRow.Value(resfld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C5_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C5_b" & y, "Filename: " & GetDefaultMember6(pRow.Value(filenamefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_" & y, "Details")
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_a" & y, "Date: " & GetDefaultMember6(pRow.Value(datefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_b" & y, "Time: " & GetDefaultMember6(pRow.Value(timefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_c" & y, "Fire ID: " & GetDefaultMember6(pRow.Value(idfld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_d" & y, "Username: " & GetDefaultMember6(pRow.Value(usernamefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_e" & y, "User Dept: " & GetDefaultMember6(pRow.Value(deptfld)))
				ElseIf GetDefaultMember6(pRow.Value(logtypefld)) = "Email" Then
					frmFire_Log_Details.TreeView1.Nodes.Add(, , "R" & y, GetDefaultMember6(pRow.Value(logtypefld)) & " " & GetDefaultMember6(pRow.Value(datefld)) & " " & GetDefaultMember6(pRow.Value(timefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C1_" & y, "Email Type: " & GetDefaultMember6(pRow.Value(emailtypefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C5_" & y, "File")
					frmFire_Log_Details.TreeView1.Nodes.Add("C5_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C5_a" & y, "Filename: " & GetDefaultMember6(pRow.Value(filenamefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("R" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_" & y, "Details")
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_a" & y, "Date: " & GetDefaultMember6(pRow.Value(datefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_b" & y, "Time: " & GetDefaultMember6(pRow.Value(timefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_c" & y, "Fire ID: " & GetDefaultMember6(pRow.Value(idfld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_d" & y, "Username: " & GetDefaultMember6(pRow.Value(usernamefld)))
					frmFire_Log_Details.TreeView1.Nodes.Add("C6_" & y, MSComctlLib.TreeRelationshipConstants.tvwChild, "C6_e" & y, "User Dept: " & GetDefaultMember6(pRow.Value(deptfld)))
				End If
			Next
		End If
	End Sub

	' UPGRADE_INFO (#0551): The 'filename' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'mydate' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'mytime' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'emailtype' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub Update_Log_Email(ByRef filename As String, ByRef mydate As String, ByRef mytime As String, ByRef emailtype As String)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Not used in the project
		'Called By: clsFire_Email*
		'Calls
		'Accepts: filename attached to email, date and time of email, type of email
		'Returns
		'******************************************************************************************************************
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		Dim pfeatworkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = myIncidentInfo.IncidentWorkspace
		
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable = pfeatworkspace.OpenTable("Fire_Log")
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow = pTable.CreateRow()
		
		Dim deptstring As String = ""
		If modGlobals.g_UserDept = g_Fire_Dept.DEC Then
			deptstring = "DEC"
		ElseIf modGlobals.g_UserDept = g_Fire_Dept.FESA Then
			deptstring = "FESA"
		End If
		
		'Set Log Table properties
		modCommon_Functions.SetTableValue(pRow, "Log_Fire_ID", myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber)
		modCommon_Functions.SetTableValue(pRow, "Log_Username", ByVal6(modGlobals.g_Username))
		modCommon_Functions.SetTableValue(pRow, "Log_User_Dept", deptstring)
		modCommon_Functions.SetTableValue(pRow, "Log_Date", mydate)
		modCommon_Functions.SetTableValue(pRow, "Log_Time", mytime)
		modCommon_Functions.SetTableValue(pRow, "Log_Filename", filename)
		modCommon_Functions.SetTableValue(pRow, "Log_Type", "Email")
		modCommon_Functions.SetTableValue(pRow, "Log_Email_Type", emailtype)
		
		'Update Log Window
		modFire_Tools.Update_Log_Window()
	End Sub

    '	Public Sub StartTrimUniqueValues()
    '		'******************************************************************************************************************
    '		'Date: 06 / 06 / 06
    '		'Author: Nathan Eaton, CALM
    '		'Description: Main Module for performing the Legend Limiter functionality
    '		'Called By: Many
    '		'Calls
    '		'Accepts: Nada
    '		'Returns
    '		'******************************************************************************************************************

    '		On Error GoTo errhandler

    '		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
    '		Dim player As ESRI.ArcGIS.Carto.IDataLayer2
    '		Dim pGeoFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
    '		' UPGRADE_INFO (#0701): The 'myOrigUVRend' member has been removed because it isn't used anywhere in current application.
    '		' EXCLUDED: Dim myOrigUVRend As ESRI.ArcGIS.Carto.IUniqueValueRenderer
    '		' UPGRADE_INFO (#0701): The 'continue' member has been removed because it isn't used anywhere in current application.
    '		' EXCLUDED: Dim [continue] As Boolean
    '		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
    '		Dim complayercount As Short
    '		Dim j As Short

    '		For j = 0 To pMxDoc.FocusMap.LayerCount - 1
    '			If TypeOf pMxDoc.FocusMap.Layer(j) Is ESRI.ArcGIS.Carto.IGroupLayer Then
    '				pcomplayer = pMxDoc.FocusMap.Layer(j)
    '				For complayercount = 0 To pcomplayer.Count - 1
    '					player = pcomplayer.Layer(complayercount)
    '					If TypeOf pcomplayer.Layer(complayercount) Is ESRI.ArcGIS.Carto.IGeoFeatureLayer Then
    '						pGeoFeatLayer = pcomplayer.Layer(complayercount)
    '						If TypeOf pGeoFeatLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
    '							If MyLayerAdded(player) = True Then
    '								TrimUniqueValues(pGeoFeatLayer)
    '							Else
    '								Save_Renderer_2_Blob(pGeoFeatLayer, player)
    '								layers_added.Add((GetPath(player)))
    '								TrimUniqueValues(pGeoFeatLayer)
    '							End If
    '						End If
    '					End If
    '				Next
    '			Else
    '				player = pMxDoc.FocusMap.Layer(j)
    '				If TypeOf pMxDoc.FocusMap.Layer(j) Is ESRI.ArcGIS.Carto.IGeoFeatureLayer Then
    '					pGeoFeatLayer = pMxDoc.FocusMap.Layer(j)
    '					If TypeOf pGeoFeatLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
    '						If MyLayerAdded(player) = True Then
    '							TrimUniqueValues(pGeoFeatLayer)
    '						Else
    '							Save_Renderer_2_Blob(pGeoFeatLayer, player)
    '							layers_added.Add((GetPath(player)))
    '							TrimUniqueValues(pGeoFeatLayer)
    '						End If
    '					End If
    '				End If
    '			End If

    '		Next

    '		Exit Sub

    'errhandler:
    '		MsgBox6(Err.Description & "starttrim")
    '	End Sub
	
    'Public Sub FinishTrimUniqueValues()
    '	'******************************************************************************************************************
    '	'Date: 06 / 06 / 06
    '	'Author: Nathan Eaton, CALM
    '	'Description: Undo the Legend Limiting functionality in the project
    '	'Called By: clsTrimUnique, when turning legend limiting off
    '	'Calls
    '	'Accepts: Nada
    '	'Returns
    '	'******************************************************************************************************************

    '	Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
    '	Dim player As ESRI.ArcGIS.Carto.IDataLayer2
    '	Dim pGeoFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
    '	' UPGRADE_INFO (#0701): The 'myOrigUVRend' member has been removed because it isn't used anywhere in current application.
    '	' EXCLUDED: Dim myOrigUVRend As ESRI.ArcGIS.Carto.IUniqueValueRenderer
    '	' UPGRADE_INFO (#0701): The 'continue' member has been removed because it isn't used anywhere in current application.
    '	' EXCLUDED: Dim [continue] As Boolean
    '	Dim pOrigUVRend As ESRI.ArcGIS.Carto.IUniqueValueRenderer
    '	Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
    '	Dim complayercount As Short
    '	Dim j As Short

    '	For j = 0 To pMxDoc.FocusMap.LayerCount - 1
    '		If TypeOf pMxDoc.FocusMap.Layer(j) Is ESRI.ArcGIS.Carto.IGroupLayer Then
    '			pcomplayer = pMxDoc.FocusMap.Layer(j)
    '			For complayercount = 0 To pcomplayer.Count - 1
    '				player = pcomplayer.Layer(complayercount)
    '				If TypeOf pcomplayer.Layer(complayercount) Is ESRI.ArcGIS.Carto.IGeoFeatureLayer Then
    '					pGeoFeatLayer = pcomplayer.Layer(complayercount)
    '					If TypeOf pGeoFeatLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
    '						If MyLayerAdded(player) = True Then
    '							pOrigUVRend = Load_Blob_2_Renderer(pGeoFeatLayer, player)
    '							pGeoFeatLayer.Renderer = pOrigUVRend
    '						End If
    '					End If
    '				End If
    '			Next
    '		Else
    '			player = pMxDoc.FocusMap.Layer(j)
    '			If TypeOf pMxDoc.FocusMap.Layer(j) Is ESRI.ArcGIS.Carto.IGeoFeatureLayer Then
    '				pGeoFeatLayer = pMxDoc.FocusMap.Layer(j)
    '				If TypeOf pGeoFeatLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
    '					If MyLayerAdded(player) = True Then
    '						pOrigUVRend = Load_Blob_2_Renderer(pGeoFeatLayer, player)
    '						pGeoFeatLayer.Renderer = pOrigUVRend
    '					End If
    '				End If
    '			End If
    '		End If
    '	Next

    'End Sub
	
	' UPGRADE_INFO (#0551): The 'player' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function MyLayerAdded(ByRef player As ESRI.ArcGIS.Carto.IDataLayer2) As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Check to see if a layer has already been added to the TrimUniqueValues process
		'Called By: Many
		'Calls
		'Accepts: Nada
		'Returns: True if added, False if the layer has not yet been added
		'******************************************************************************************************************
		
		MyLayerAdded = False
		Dim i As Short
		
		If layers_added.Count() > 0 Then
			For i = 1 To layers_added.Count() ' - 1
				If GetDefaultMember6(layers_added(i)) = GetPath(player) Then
					MyLayerAdded = True
				End If
			Next
		End If
		
	End Function
	
    '' UPGRADE_INFO (#0551): The 'pgflayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    '' UPGRADE_INFO (#0551): The 'pdlayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    'Public Sub Save_Renderer_2_Blob(ByRef pgflayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer, ByRef pdlayer As ESRI.ArcGIS.Carto.IDataLayer2)
    '	'******************************************************************************************************************
    '	'Date: 06 / 06 / 06
    '	'Author: Nathan Eaton, CALM
    '	'Description: Save a layer's renderer to a blob file to be accessed and used by the trim unique value functionality
    '	'             to keep a reference to the intial renderer associated with the layer
    '	'Called By: StartTrimUniqueValues
    '	'Calls
    '	'Accepts: geofeaturelayer and datalayer
    '	'Returns
    '	'******************************************************************************************************************

    '	Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
    '	Dim pPropSet As ESRI.ArcGIS.esriSystem.IPropertySet = New ESRI.ArcGIS.esriSystem.PropertySet()
    '	pPropSet.SetProperty(GetPath(pdlayer), pgflayer.Renderer)

    '	Dim pPS As ESRI.ArcGIS.esriSystem.IPersistStream = pPropSet
    '	Dim pMBS As ESRI.ArcGIS.esriSystem.IMemoryBlobStream = New ESRI.ArcGIS.esriSystem.MemoryBlobStream()
    '	pPS.Save(pMBS, 0)
    '	Dim FSO As New VB6FileSystemObject
    '	Dim filename As String = Replace(GetPath(pdlayer), VB.Right(GetPath(pdlayer), 3), "blb")
    '	Dim blbfilename As String = Environ("TEMP") & "\" & Extract_filename(filename)

    '	If FSO.FileExists(blbfilename) = False Then
    '		pMBS.SaveToFile(blbfilename)
    '	Else
    '		FSO.DeleteFile(blbfilename, True)
    '		pMBS.SaveToFile(blbfilename)
    '	End If

    'End Sub
	
	' UPGRADE_INFO (#0551): The 'pgflayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'pdlayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function Load_Blob_2_Renderer(ByRef pgflayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer, ByRef pdlayer As ESRI.ArcGIS.Carto.IDataLayer2) As ESRI.ArcGIS.Carto.IUniqueValueRenderer
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Load a blob file as a renderer to a layer
		'Called By: Many
		'Calls
		'Accepts: GeoFeatureLayer, Datalayer
		'Returns: UniqueValueRenderer
		'******************************************************************************************************************
		
		Dim pMBS As ESRI.ArcGIS.esriSystem.IMemoryBlobStream = New ESRI.ArcGIS.esriSystem.MemoryBlobStream()
		
		Dim filename As String = Replace(GetPath(pdlayer), VB.Right(GetPath(pdlayer), 3), "blb")
		Dim blbfilename As String = Environ("TEMP") & "\" & Extract_filename(filename)
		
		pMBS.LoadFromFile(blbfilename)
		
		Dim pPS As ESRI.ArcGIS.esriSystem.IPersistStream = New ESRI.ArcGIS.esriSystem.PropertySet()
		pPS.Load(pMBS)
		
		Dim pPropSet As ESRI.ArcGIS.esriSystem.IPropertySet = pPS
		
		Return pPropSet.GetProperty(GetPath(pdlayer))
		
	End Function
	
	' UPGRADE_INFO (#0551): The 'pGeoFeatLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub TrimUniqueValues(ByRef pGeoFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Modified by Nathan Eaton, CALM
		'Description: The actual Legend Limiter functionality that limits renderers to view extent
		'Called By: StartTrimUnique
		'Calls
		'Accepts: GeoFeatureLayer
		'Returns
		'******************************************************************************************************************
		
		On Error GoTo ErrHand
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim player As ESRI.ArcGIS.Carto.ILayer = pGeoFeatLayer

		Dim pOrigUVRend As ESRI.ArcGIS.Carto.IUniqueValueRenderer = Load_Blob_2_Renderer(pGeoFeatLayer, player)
		
		Dim pNewUVRend As ESRI.ArcGIS.Carto.IUniqueValueRenderer = New ESRI.ArcGIS.Carto.UniqueValueRenderer()
		
		' --------------------------------------------
		
		Dim iFieldCount As Short = pOrigUVRend.FieldCount
		pNewUVRend.FieldCount = iFieldCount
		
		Dim strField() As String
		ReDim strField(iFieldCount)
		Dim i As Short
		For i = 0 To iFieldCount - 1
			strField(i) = pOrigUVRend.Field(i)
			pNewUVRend.Field(i) = strField(i)
		Next
		
		' carry over a few other properties
		' must do default symbol stuff after setting fields and field count, but before adding any values
		pNewUVRend.DefaultSymbol = pOrigUVRend.DefaultSymbol
		pNewUVRend.UseDefaultSymbol = pOrigUVRend.UseDefaultSymbol ' must set this after setting default symbol
		pNewUVRend.DefaultLabel = pOrigUVRend.DefaultLabel
		pNewUVRend.ColorScheme = pOrigUVRend.ColorScheme
		
		Dim newRotRend As ESRI.ArcGIS.Carto.IRotationRenderer = pNewUVRend
		Dim porigRotRend As ESRI.ArcGIS.Carto.IRotationRenderer = pOrigUVRend
		newRotRend.RotationField = porigRotRend.RotationField
		newRotRend.RotationType = porigRotRend.RotationType
		
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pGeoFeatLayer.FeatureClass
		Dim iFieldIndex() As Short
		ReDim iFieldIndex(iFieldCount)
		
		For i = 0 To iFieldCount - 1
			iFieldIndex(i) = pFeatClass.FindField(pNewUVRend.Field(i))
		Next
		
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		Dim pFeatCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		
		Dim pFLDef As ESRI.ArcGIS.Carto.IFeatureLayerDefinition
		Dim j As Short
		Dim pSym As ESRI.ArcGIS.Display.ISymbol
		Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
		Dim pRendSym() As ESRI.ArcGIS.Display.ISymbol
		Dim strRendVal() As String
		Dim c As Short
		Dim origfeatselset As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim origselset As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim tempselset As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim pFeatSelect As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim pSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		If Not ONLY_FEATURES_IN_EXTENT Then
			' handle possible definition query on layer
			pFLDef = pGeoFeatLayer
			
			pQueryFilter.WhereClause = pFLDef.DefinitionExpression
			pFeatCursor = pFeatClass.Search(pQueryFilter, False)
			
			i = 0
			
			c = 0
			
			 Else ' trim to only features within current extent
			'Dim tempfeatlayer As IFeatureLayer
			'Set tempfeatlayer = New FeatureLayer
			'Set tempfeatlayer = pGeoFeatLayer
			origfeatselset = pGeoFeatLayer 'tempfeatlayer
			origselset = origfeatselset.SelectionSet
			tempselset = origselset
			
			SelectLayerFeaturesInCurrentExtent(pGeoFeatLayer)
			
			pFeatSelect = pGeoFeatLayer
			pSelSet = pFeatSelect.SelectionSet
			
			' create query filter based on selection
			pSelSet.Search(pQueryFilter, True, pFeatCursor)
			'pFeatSelect.Clear
			
			pFeatSelect.SelectionSet = tempselset 'origselset
			
		End If

		' --------------------------------------------
		' loop thru features, storing info for new renderer in two temp arrays
		pFeat = pFeatCursor.NextFeature()
		Dim strValue() As String
		Dim strValueString As String = ""
		ReDim strValue(iFieldCount)
		Dim bExpectedErr As Boolean
		Dim bAddValue As Boolean
		Dim bExpectedErr2 As Boolean
		Dim bHasReferenceValue As Boolean
		Dim strRefValueString As String = ""
		Dim bNoValueFound As Boolean
		Dim bValueFound As Boolean
		Dim bNoRefValueFound As Boolean
		Dim k As Short
		
		Do While pFeat IsNot Nothing
			
			For i = 0 To iFieldCount - 1
				If pFeat.Fields.Field(i).IsNullable Then
					If Not IsNull6(GetDefaultMember6(pFeat.Value(iFieldIndex(i)))) Then
						strValue(i) = GetDefaultMember6(pFeat.Value(iFieldIndex(i)))
					Else
						strValue(i) = "<Null>"
					End If
				Else
					strValue(i) = GetDefaultMember6(pFeat.Value(iFieldIndex(i)))
				End If
			Next
			
			' use BuildValueString to make sure we're correct when rendering based on one, two, or three fields
			strValueString = BuildValueString(strValue)
			
			bNoValueFound = True
			' search array to see if we've already added this value
			For j = 0 To c - 1
				If strRendVal(j) = strValueString Then
					bNoValueFound = False
					Exit For
				End If
			Next
			
			If bNoValueFound Then
				
				' get symbol from old renderer
				pSym = Nothing
				bExpectedErr = True
				
				pSym = pOrigUVRend.Symbol(strValueString)
				bExpectedErr = False
				bAddValue = False
				If pSym IsNot Nothing Then
					' add value if found in old renderer
					bAddValue = True
				Else
					' we may still add if value is found on old renderer, but with a reference value
					bHasReferenceValue = True
					bExpectedErr2 = True
					strRefValueString = pOrigUVRend.ReferenceValue(strValueString) ' throws error if no ref val
					bExpectedErr2 = False
					If bHasReferenceValue Then
						bAddValue = True
						' add ref value (if not already added)
						bNoRefValueFound = True
						For k = 0 To c - 1
							If strRendVal(k) = strValueString Then
								bNoRefValueFound = False
								Exit For
							End If
						Next
						If bNoRefValueFound Then
							Debug.WriteLine(strRefValueString & " as ref val")
							' add symbol and ref value string to temporary, paired arrays
							c += 1
							ReDim Preserve pRendSym(c)
							ReDim Preserve strRendVal(c)
							pRendSym(c - 1) = pOrigUVRend.Symbol(strRefValueString)
							strRendVal(c - 1) = strRefValueString
						End If
					End If
				End If
				
				If bAddValue Then
					
					' add symbol and value string to temporary, paired arrays
					c += 1
					ReDim Preserve pRendSym(c)
					ReDim Preserve strRendVal(c)
					pRendSym(c - 1) = pSym
					strRendVal(c - 1) = strValueString
				End If
				
			End If
			pFeat = pFeatCursor.NextFeature()
		Loop
		
		Debug.WriteLine(" ")
		pMxDoc.ActiveView.Refresh()
		
		' --------------------------------------------
		' loop thru old renderer values, search for value string in temp array.
		'   array.  if found, add to new renderer under correct Heading and with correct Label
		
		For i = 0 To pOrigUVRend.ValueCount - 1
			Debug.WriteLine(pOrigUVRend.Value(i))
			
			strValueString = pOrigUVRend.Value(i)
			bValueFound = False
			
			For j = 0 To c - 1
				If strRendVal(j) = strValueString Then
					bValueFound = True
					pSym = pRendSym(j)
					Exit For
				End If
			Next
			
			If bValueFound Then
				
				bExpectedErr2 = True
				bHasReferenceValue = True
				
				strRefValueString = pOrigUVRend.ReferenceValue(strValueString)
				
				bExpectedErr2 = False
				If Not bHasReferenceValue Then
					
					pNewUVRend.AddValue(strValueString, pOrigUVRend.Heading(strValueString), pSym)
					
				Else
					
					pNewUVRend.AddReferenceValue(strValueString, strRefValueString)
					
				End If
				
				pNewUVRend.Label(strValueString) = pOrigUVRend.Label(strValueString)
			End If
			
		Next
		
		' --------------------------------------------
		' set new renderer to layer and refresh map
		pGeoFeatLayer.Renderer = pNewUVRend
		pMxDoc.UpdateContents()
		pMxDoc.ActiveView.Refresh()

		Exit Sub
ErrHand:
		' in two special cases where errors are expected, we clear the error and resume execution
		If bExpectedErr Then
			Err.Clear()
			Resume Next
		ElseIf bExpectedErr2 Then
			Err.Clear()
			bHasReferenceValue = False
			Resume Next
		Else
			' re-raise if we have a real error
			Err.Raise(Err.Number, Err.Source, Err.Description)
		End If
	End Sub

	' UPGRADE_INFO (#0551): The 'ListOfValues' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Function BuildValueString(ByRef ListOfValues() As String) As String
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Modified by Nathan Eaton, CALM
		'Description: builds unique values renderer "value string" based on one, two, or three fields
		'Called By: Many
		'Calls
		'Accepts: Array of values
		'Returns
		'******************************************************************************************************************
		
		Dim strRetVal As String = ""
		Dim iNumFields As Short = UBound6(ListOfValues)
		
		' create value string based on known number of fields on renderer
		If iNumFields = 1 Then 'Format: "Value"
			strRetVal = ListOfValues(0)
		ElseIf iNumFields = 2 Then 'Format: "Value1, Value2"
			strRetVal = ListOfValues(0) & ", " & ListOfValues(1)
		Else 'Format: "Value1, Value2, Value3"
			strRetVal = ListOfValues(0) & ", " & ListOfValues(1) & ", " & ListOfValues(2)
		End If
		
		Return strRetVal
		
	End Function

	' UPGRADE_INFO (#0551): The 'selectLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub SelectLayerFeaturesInCurrentExtent(ByRef selectLayer As ESRI.ArcGIS.Carto.IFeatureLayer)
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Modified by Nathan Eaton, CALM
		'Description: selects the featrues in the current view extent
		'Called By: Many
		'Calls
		'Accepts: selectedlayer
		'Returns
		'******************************************************************************************************************
		
		Dim pMxDocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		
		' --------------------------------------------------------------------------
		' for each layer: preserve whether it is selectable and make it non-selectable
		Dim i As Short
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pLayer2 As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim SelectableLayers(MAX_LAYER_COUNT) As Boolean
		Dim SelSets(MAX_LAYER_COUNT) As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim pFeatSel As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim complayercount As Short
		
		For i = 0 To pMxDocument.FocusMap.LayerCount - 1
			player = pMxDocument.FocusMap.Layer(i)
			If TypeOf player Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = player
				For complayercount = 0 To pcomplayer.Count - 1
					pLayer2 = pcomplayer.Layer(complayercount)
					If TypeOf pLayer2 Is ESRI.ArcGIS.Carto.IFeatureLayer Then
						pFeatLayer = pLayer2
						pFeatSel = pFeatLayer
						SelSets(i) = pFeatSel.SelectionSet
						SelectableLayers(i) = pFeatLayer.Selectable
						pFeatLayer.Selectable = False
					End If
				Next
			ElseIf TypeOf player Is ESRI.ArcGIS.Carto.IFeatureLayer Then
				pFeatLayer = player
				pFeatSel = pFeatLayer
				SelSets(i) = pFeatSel.SelectionSet
				SelectableLayers(i) = pFeatLayer.Selectable
				pFeatLayer.Selectable = False
			End If
		Next
		' make only our layer selectable
		selectLayer.Selectable = True
		
		' --------------------------------------------------------------------------
		' select features within current extent
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView = pMxDocument.FocusMap
		Dim pDisplayTransform As ESRI.ArcGIS.Display.IDisplayTransformation = pActiveView.ScreenDisplay.DisplayTransformation
		Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope = pDisplayTransform.FittedBounds
		
		' perform the selection
		Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDocument.FocusMap
		Dim pMxApp As ESRI.ArcGIS.ArcMapUI.IMxApplication = m_pApp
		pMap.SelectByShape(pEnvelope, pMxApp.SelectionEnvironment, False) ' problem here
		
		' --------------------------------------------------------------------------
		' return selectable layers list to original state
		
		For i = 0 To pMxDocument.FocusMap.LayerCount - 1
			
			player = pMxDocument.FocusMap.Layer(i)
			If TypeOf player Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = player
				For complayercount = 0 To pcomplayer.Count - 1
					pLayer2 = pcomplayer.Layer(complayercount)
					If TypeOf pLayer2 Is ESRI.ArcGIS.Carto.IFeatureLayer Then
						pFeatLayer = pLayer2
						If pFeatLayer.Name <> selectLayer.Name Then
							pFeatSel = pFeatLayer
							pFeatSel.SelectionSet = SelSets(i)
						End If
						pFeatLayer.Selectable = SelectableLayers(i)
					End If
				Next
			ElseIf TypeOf player Is ESRI.ArcGIS.Carto.IFeatureLayer Then
				pFeatLayer = player
				If pFeatLayer.Name <> selectLayer.Name Then
					pFeatSel = pFeatLayer
					pFeatSel.SelectionSet = SelSets(i)
				End If
				pFeatLayer.Selectable = SelectableLayers(i)
			End If
		Next
		
	End Sub
	
	' UPGRADE_INFO (#0551): The 'selectLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub makeOnlySelectableLayer(ByRef selectLayer As ESRI.ArcGIS.Carto.IFeatureLayer)
		'******************************************************************************************************************
		'Date: 10 / 10 / 06
		'Author: Modified by Jack Green, DEC
		'Description: Set the selectLayer as the only selectable layer
		'Called By: Many
		'Calls
		'Accepts: selectLayer
		'Returns
		'******************************************************************************************************************
		
		Dim pMxDocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		
		Dim i As Short
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pLayer2 As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		' UPGRADE_INFO (#0701): The 'pFeatSel' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatSel As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim complayercount As Short
		
		For i = 0 To pMxDocument.FocusMap.LayerCount - 1
			player = pMxDocument.FocusMap.Layer(i)
			If TypeOf player Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = player
				For complayercount = 0 To pcomplayer.Count - 1
					pLayer2 = pcomplayer.Layer(complayercount)
					If TypeOf pLayer2 Is ESRI.ArcGIS.Carto.IFeatureLayer Then
						pFeatLayer = pLayer2
						pFeatLayer.Selectable = False
					End If
				Next
			ElseIf TypeOf player Is ESRI.ArcGIS.Carto.IFeatureLayer Then
				pFeatLayer = player
				pFeatLayer.Selectable = False
			End If
		Next
		' make only our layer selectable
		selectLayer.Selectable = True
		
	End Sub
	
	' UPGRADE_INFO (#0701): The 'LabelVisibleFeatures' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Public Sub LabelVisibleFeatures(ByRef player As ESRI.ArcGIS.Carto.IGeoFeatureLayer)
		' EXCLUDED: '******************************************************************************************************************
		' EXCLUDED: 'Date: 06 / 06 / 06
		' EXCLUDED: 'Author: Nathan Eaton, CALM
		' EXCLUDED: 'Description: Not used in Project
		' EXCLUDED: 'Called By: Many
		' EXCLUDED: 'Calls
		' EXCLUDED: 'Accepts
		' EXCLUDED: 'Returns
		' EXCLUDED: '******************************************************************************************************************
		' EXCLUDED: On Error GoTo ErrorHandler
		
		' EXCLUDED: 'Get the focus map
		' EXCLUDED: Dim pMxDocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' EXCLUDED: Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDocument.FocusMap
		' EXCLUDED: Dim pAnnoLayerPropsColl As ESRI.ArcGIS.Carto.IAnnotateLayerPropertiesCollection = player.AnnotationProperties
		' EXCLUDED: Dim pAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerProperties
		' EXCLUDED: Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView = pMap

		' EXCLUDED: pAnnoLayerPropsColl.QueryItem(0, pAnnoLayerProps, Nothing, Nothing)
		' EXCLUDED: pAnnoLayerProps.LabelWhichFeatures = ESRI.ArcGIS.Carto.esriLabelWhichFeatures.esriVisibleFeatures

		' EXCLUDED: Exit Sub
		
		' EXCLUDED: ErrorHandler:
		' EXCLUDED: MsgBox6("An error has occured within LabelFeatures." & ControlChars.Cr & ControlChars.Cr & "Error Details : " & Err.Description, MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton1, "Error")
	' EXCLUDED: End Sub

	Public Function GetEmails() As Collection
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Creates a collection of emails based on the work centre specified by the user when creating/opening the incident
		'Called By: clsFIre_Email*
		'Calls
		'Accepts: Nada
		'Returns: Collection of required Emails
		'******************************************************************************************************************
		
		Dim getemailslist As New Collection
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		getemailslist.Add(("state1_sdo@calm.wa.gov.au"))
		
		If myIncidentInfo.incidentRegion = "Goldfields" Then
			getemailslist.Add(("goldfields_rdo@calm.wa.gov.au"))
			getemailslist.Add(("goldfields_ddo@calm.wa.gov.au"))
			getemailslist.Add(("kalg1_MgmtSupprt@calm.wa.gov.au"))
		ElseIf myIncidentInfo.incidentRegion = "Kimberley" Then
			getemailslist.Add(("kimberley_rdo@calm.wa.gov.au"))
			getemailslist.Add(("kimberley_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Broome" Then
				getemailslist.Add(("broo1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Kununurra" Then
				getemailslist.Add(("kunu1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Midwest" Then
			getemailslist.Add(("midwest_rdo@calm.wa.gov.au"))
			getemailslist.Add(("midwest_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Denham" Then
				getemailslist.Add(("denh1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Geraldton" Then
				getemailslist.Add(("gera1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Jurien" Then
				getemailslist.Add(("juri1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Pilbara" Then
			getemailslist.Add(("pilbara_rdo@calm.wa.gov.au"))
			getemailslist.Add(("pilbara_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Exmouth" Then
				getemailslist.Add(("exmo1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Karratha" Then
				getemailslist.Add(("karr1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Sth Coast" Then
			getemailslist.Add(("southcoast_rdo@calm.wa.gov.au"))
			getemailslist.Add(("southcoast_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Albany" Then
				getemailslist.Add(("alba1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Esperance" Then
				getemailslist.Add(("espe1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Sth West" Then
			getemailslist.Add(("southwest_rdo@calm.wa.gov.au"))
			getemailslist.Add(("southwest_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Bunbury" Then
				getemailslist.Add(("bunb1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Busselton" Then
				getemailslist.Add(("buss1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Collie" Then
				getemailslist.Add(("coll1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Harvey" Then
				getemailslist.Add(("harv1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Kirup" Then
				getemailslist.Add(("kiru1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Margaret River" Then
				getemailslist.Add(("marg1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Swan" Then
			getemailslist.Add(("swan_rdo@calm.wa.gov.au"))
			getemailslist.Add(("swan_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Dwellingup" Then
				getemailslist.Add(("dwel1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Jarrahdale" Then
				getemailslist.Add(("jarr1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Mundaring" Then
				getemailslist.Add(("mund1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Wanneroo" Then
				getemailslist.Add(("wann1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Yanchep" Then
				getemailslist.Add(("yanc1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Warren" Then
			getemailslist.Add(("warren_rdo@calm.wa.gov.au"))
			getemailslist.Add(("warren_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Manjimup" Then
				getemailslist.Add(("manj1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Pemberton" Then
				getemailslist.Add(("pemb1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Walpole" Then
				getemailslist.Add(("walp1_MgmtSupprt@calm.wa.gov.au"))
			End If
		ElseIf myIncidentInfo.incidentRegion = "Wheatbelt" Then
			getemailslist.Add(("wheatbelt_rdo@calm.wa.gov.au"))
			getemailslist.Add(("wheatbelt_ddo@calm.wa.gov.au"))
			If myIncidentInfo.IncidentWorkCentre = "Katanning" Then
				getemailslist.Add(("kata1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Merredin" Then
				getemailslist.Add(("merr1_MgmtSupprt@calm.wa.gov.au"))
			ElseIf myIncidentInfo.IncidentWorkCentre = "Narrogin" Then
				getemailslist.Add(("narr1_MgmtSupprt@calm.wa.gov.au"))
			End If
		End If
		Return getemailslist
		
	End Function


	' UPGRADE_INFO (#0551): The 'pActiveView' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'iResampleRatio' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub SetOutputQuality(ByRef pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByRef iResampleRatio As Integer)
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pGraphicsContainer As ESRI.ArcGIS.Carto.IGraphicsContainer
		Dim pElement As ESRI.ArcGIS.Carto.IElement
		Dim pOutputRasterSettings As ESRI.ArcGIS.Display.IOutputRasterSettings
		Dim pMapFrame As ESRI.ArcGIS.Carto.IMapFrame
		Dim pTmpActiveView As ESRI.ArcGIS.Carto.IActiveView
		
		If TypeOf pActiveView Is ESRI.ArcGIS.Carto.IMap Then
			pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
			pOutputRasterSettings.ResampleRatio = iResampleRatio
		ElseIf TypeOf pActiveView Is ESRI.ArcGIS.Carto.IPageLayout Then
			
			'assign ResampleRatio for PageLayout
			pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
			pOutputRasterSettings.ResampleRatio = iResampleRatio
			
			'and assign ResampleRatio to the Maps in the PageLayout
			pGraphicsContainer = pActiveView
			pGraphicsContainer.Reset()
			pElement = pGraphicsContainer.Next()
			Do While pElement IsNot Nothing
				If TypeOf pElement Is ESRI.ArcGIS.Carto.IMapFrame Then
					pMapFrame = pElement
					pTmpActiveView = pMapFrame.Map
					pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
					pOutputRasterSettings.ResampleRatio = iResampleRatio
				End If
				DoEvents6()
				pElement = pGraphicsContainer.Next()
			Loop
			pMap = Nothing
			pMapFrame = Nothing
			pGraphicsContainer = Nothing
			pTmpActiveView = Nothing
		End If
		pOutputRasterSettings = Nothing
		
	End Sub

	Public Sub restoreLegend()
		'******************************************************************************************************************
		'Date: 10 / 10 / 06
		'Author: Jack Green, DEC
		'Description: Restore the legend for all feature classes in the geodatabase
		'Called By: Many
		'Calls
		'Accepts
		'Returns
		'******************************************************************************************************************
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		modCommon_Functions.Apply_Layer_File_Symbology(modFire_Tools.Get_Fire_Poly_Layer(), Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Boundary.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(modFire_Tools.Get_Fire_Line_Layer(), Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Line.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(modFire_Tools.Get_Fire_Point_Layer(), Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Point.lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(modFire_Tools.Get_Fire_Ass_Layer(), Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Assignment.lyr")
		
		pMxDoc.UpdateContents()
		
	End Sub

	Public Function getUserName() As String
		'******************************************************************************************************************
		'Date: 24 / 10 / 06
		'Author: Jack Green, DEC
		'Description: Return the username. If one doesn't exist prompts the user to enter it
		'Called By: Many
		'Calls
		'Accepts: Nothing
		'Returns: Username as string
		'******************************************************************************************************************
		
		Dim myUserName As String = g_Username
		
		If myUserName = "" Then
			myUserName = InputBox6("Enter your username:", "You are not logged in")
			g_Username = myUserName
		End If
		
		Return myUserName
	End Function

End Module
