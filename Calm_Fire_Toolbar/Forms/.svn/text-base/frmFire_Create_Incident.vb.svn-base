' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Create_Incident

	Private Sub btnNewGDB_Click() Handles btnNewGDB.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Specify the location of the new Fire Geodatabase
		'Called By: None
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		Dim dest_string As String = modCommon_Functions.NewGeodatabase(Me.hWnd)
		
		'(modCommon_Functions.m_pApp.Hwnd)
		If dest_string = "empty" Then
			Exit Sub
		End If
		
		Me.txtNewGDB.Text = dest_string
		
	End Sub

	Private Sub btnOK_Click() Handles btnOK.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Create and Load the Fire Geodatabase as a Group Layer
		'Called By: None
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		On Error GoTo ErrorHandler
		
		'Ensure a location has been specified for the Geodatabase
		If Me.txtNewGDB.Text = "" Then
			MsgBox6("Please specify a new fire incident geodatabase", MsgBoxStyle.Critical, "Geodatabase")
			Me.txtNewGDB.SetFocus()
			Exit Sub
		End If
		
		'changed 28-08-2007 by FMS request - GrantB
		'construct a query to make sure the fire Number does not contain a string
		'If Me.txtInc_Number.Text Like "*a*" Or Me.txtInc_Number.Text Like "*b*" Or Me.txtInc_Number.Text Like "*c*" Or
		' ' Me.txtInc_Number.Text Like "*d*" Or Me.txtInc_Number.Text Like "*e*" Or Me.txtInc_Number.Text Like "*f*" Or
		' '  Me.txtInc_Number.Text Like "*g*" Or Me.txtInc_Number.Text Like "*h*" Or Me.txtInc_Number.Text Like "*i*" Or
		' '   Me.txtInc_Number.Text Like "*k*" Or Me.txtInc_Number.Text Like "*l*" Or Me.txtInc_Number.Text Like "*m*" Or
		' '    Me.txtInc_Number.Text Like "*m*" Or Me.txtInc_Number.Text Like "*n*" Or Me.txtInc_Number.Text Like "*o*" Or
		' '     Me.txtInc_Number.Text Like "*p*" Or Me.txtInc_Number.Text Like "*q*" Or Me.txtInc_Number.Text Like "*r*" Or
		' '      Me.txtInc_Number.Text Like "*s*" Or Me.txtInc_Number.Text Like "*t*" Or Me.txtInc_Number.Text Like "*u*" Or
		' '       Me.txtInc_Number.Text Like "*v*" Or Me.txtInc_Number.Text Like "*w*" Or Me.txtInc_Number.Text Like "*x*" Or
		' '        Me.txtInc_Number.Text Like "*y*" Or Me.txtInc_Number.Text Like "*z*" Then
		'                MsgBox "Please specify an integer fire number that doesnt contain any letters", vbCritical, "Fire Number"
		'                Exit Sub
		'changed 28-08-2007 by FMS request - GrantB
		If IsNumeric6(Me.txtInc_Number.Text) Then
			If Len6(Me.txtInc_Number.Text) <> 3 Then
				MsgBox6("Please type in only three numbers for the Fire Number. eg 012", MsgBoxStyle.Exclamation, "Fire Number")
				Exit Sub
			End If
		Else
			MsgBox6("Please type in three numbers for the Fire Number. eg 012", MsgBoxStyle.Exclamation, "Fire Number")
			Exit Sub
		End If

		Dim firenum As Short = Me.txtInc_Number.Text
		
		'Reference the Fire Incident Info Class
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		'Set Incident Parameters
		myIncidentInfo.IncidentNumber = firenum
		myIncidentInfo.incidentRegion = Me.cbInc_Region.Text
		myIncidentInfo.IncidentRegionEx = Me.cbInc_Region_ext.Text
		myIncidentInfo.IncidentWorkCentre = Me.cbWorkCentre.Text

		Dim fsoSysObj As VB6FileSystemObject = New VB6FileSystemObject()
		
		'Ensure Default Fire mdb has been instaleld correctly
		Dim progdir As String = modCommon_Functions.Get_ProgramFiles_Folder("DEC\GIS\ArcGIS9\Fire")

		If Fileexists(progdir & "\DEC\GIS\ArcGIS9\Fire\Fire_Incident.mdb") = False Then
			MsgBox6("Your Fire Tool has not been setup correctly, the " & progdir & "\DEC\GIS\ArcGIS9\Fire\Fire_Incident.mdb " & "file cannot be found under your program files directory", MsgBoxStyle.Critical, "Fire Tools Error")
			Exit Sub
		End If
		
		'Copy the Default Fire incident mdb to the new location
        fsoSysObj.CopyFile(progdir & "\DEC\GIS\ArcGIS9\Fire\Fire_Incident.mdb", Me.txtNewGDB.Text, True)
        Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.AccessWorkspaceFactory()
		Dim pworkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(Me.txtNewGDB.Text, modCommon_Functions.m_pApp.hWnd)
		
		myIncidentInfo.IncidentWorkspace = pworkspace
		
		'Set Incident Table properties
		modCommon_Functions.SetFireInfoProp(pworkspace, "INCIDENT_NUMBER", ByVal6(Me.txtInc_Number.Text))
		modCommon_Functions.SetFireInfoProp(pworkspace, "INCIDENT_REGION", ByVal6(Me.cbInc_Region.Text))
		modCommon_Functions.SetFireInfoProp(pworkspace, "INCIDENT_REGION_EXT", ByVal6(Me.cbInc_Region_ext.Text))
		modCommon_Functions.SetFireInfoProp(pworkspace, "INCIDENT_WORKCENTRE", ByVal6(Me.cbWorkCentre.Text))

		'load the incident feature classes into the map
		Dim pWorkspace2 As ESRI.ArcGIS.Geodatabase.IWorkspace2 = pWorkspaceFactory.OpenFromFile(Me.txtNewGDB.Text, m_pApp.hWnd)
		
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
		
		'create a new grouplayer
		Dim pgrouplayer As ESRI.ArcGIS.Carto.IGroupLayer = New ESRI.ArcGIS.Carto.GroupLayer()
		pgrouplayer.Name = myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber
		
		modGlobals.g_INCIDENT_LAYER_NAME = myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber
		
		pMap.AddLayer(pgrouplayer)
		
		pgrouplayer.Add(annoFeatureLayer)
		pgrouplayer.Add(pointFeatureLayer)
		pgrouplayer.Add(assFeatureLayer)
		pgrouplayer.Add(lineFeatureLayer)
		pgrouplayer.Add(polyFeatureLayer)
		
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
		
		pMxDoc.UpdateContents()
		
		If pMap.SpatialReference Is Nothing Then
			Call Set_DataFRame_Coord_GDA()
		End If

        Unload6(Me)
        ' it make sures that user has created or opened the incident in capturing part of the application
        g_incidentStatus = True
		Exit Sub
ErrorHandler:
        MsgBox6("Please specify an integer fire number - " & GetDefaultMember6(Err.Description), MsgBoxStyle.Critical, "Fire Number")
	End Sub

	Private Sub cbInc_Region_Click() Handles cbInc_Region.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Update all Cmboxes to reflect the Region Selection
		'Called By: None
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		Me.cbInc_Region_ext.Clear()
		Me.cbWorkCentre.Clear()
		
		'added some work centres 1/11/06
		
		If Me.cbInc_Region.Text = "Goldfields" Then
			Me.cbInc_Region_ext.AddItem(("GFD"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Kalgoorlie"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Kimberley" Then
			Me.cbInc_Region_ext.AddItem(("KMB"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Broome"))
			Me.cbWorkCentre.AddItem(("Kununurra"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Midwest" Then
			Me.cbInc_Region_ext.AddItem(("GAS"))
			Me.cbInc_Region_ext.AddItem(("GER"))
			Me.cbInc_Region_ext.AddItem(("MRA"))
			Me.cbInc_Region_ext.AddItem(("SBY"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Denham"))
			Me.cbWorkCentre.AddItem(("Geraldton"))
			Me.cbWorkCentre.AddItem(("Jurien"))
			Me.cbWorkCentre.AddItem(("Carnarvon"))
			Me.cbWorkCentre.AddItem(("Cervantes"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Pilbara" Then
			Me.cbInc_Region_ext.AddItem(("EXM"))
			Me.cbInc_Region_ext.AddItem(("PIL"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Exmouth"))
			Me.cbWorkCentre.AddItem(("Karratha"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Sth Coast" Then
			Me.cbInc_Region_ext.AddItem(("ALB"))
			Me.cbInc_Region_ext.AddItem(("ESP"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Albany"))
			Me.cbWorkCentre.AddItem(("Esperance"))
			Me.cbWorkCentre.AddItem(("Ravensthorpe"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Sth West" Then
			Me.cbInc_Region_ext.AddItem(("BWD"))
			Me.cbInc_Region_ext.AddItem(("WTN"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Bunbury"))
			Me.cbWorkCentre.AddItem(("Busselton"))
			Me.cbWorkCentre.AddItem(("Collie"))
			Me.cbWorkCentre.AddItem(("Harvey"))
			Me.cbWorkCentre.AddItem(("Kirup"))
			Me.cbWorkCentre.AddItem(("Margaret River"))
			Me.cbWorkCentre.AddItem(("Nannup"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Swan" Then
			Me.cbInc_Region_ext.AddItem(("PHL"))
			Me.cbInc_Region_ext.AddItem(("SCT"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Dwellingup"))
			Me.cbWorkCentre.AddItem(("Jarrahdale"))
			Me.cbWorkCentre.AddItem(("Mundaring"))
			Me.cbWorkCentre.AddItem(("Wanneroo"))
			Me.cbWorkCentre.AddItem(("Yanchep"))
			Me.cbWorkCentre.AddItem(("Kensington"))
			Me.cbWorkCentre.AddItem(("Fremantle"))
			Me.cbWorkCentre.AddItem(("Serpentine"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Warren" Then
			Me.cbInc_Region_ext.AddItem(("DON"))
			Me.cbInc_Region_ext.AddItem(("FKL"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Manjimup"))
			Me.cbWorkCentre.AddItem(("Pemberton"))
			Me.cbWorkCentre.AddItem(("Walpole"))
			Me.cbWorkCentre.AddItem(("Northcliffe"))
			Me.cbWorkCentre.ListIndex = 0
		ElseIf Me.cbInc_Region.Text = "Wheatbelt" Then
			Me.cbInc_Region_ext.AddItem(("KTG"))
			Me.cbInc_Region_ext.AddItem(("MDN"))
			Me.cbInc_Region_ext.AddItem(("NGN"))
			Me.cbInc_Region_ext.ListIndex = 0
			Me.cbWorkCentre.AddItem(("Katanning"))
			Me.cbWorkCentre.AddItem(("Merredin"))
			Me.cbWorkCentre.AddItem(("Narrogin"))
			Me.cbWorkCentre.ListIndex = 0
		End If
		
	End Sub
	
	Private Sub btnCancel_Click() Handles btnCancel.Click
		
		Unload6(Me)
		
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		Me.cbInc_Region.ListIndex = 0
		Me.cbInc_Region_ext.ListIndex = 0
	End Sub

	Private Sub Set_DataFRame_Coord_GDA()
		
		' UPGRADE_INFO (#0701): The 'spatialreftype' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim spatialreftype As String = ""
		' UPGRADE_INFO (#0701): The 'spatialref' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim spatialref As String = ""
		Dim spatialrefstring As String = ""
		Dim projlocation As String = ""
		Dim geolocation As String = ""

		Dim pIMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pMap As ESRI.ArcGIS.Carto.IMap = pIMxDoc.FocusMap
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Your ARCGISHOME Environment Variable has not been set. Contact the GIS Section on 9334 0158 for help", MsgBoxStyle.Exclamation, "Error in Setup")
			Exit Sub
		End If
		
		geolocation = Environ("ARCGISHOME") & "\Coordinate Systems\Geographic Coordinate Systems\Australia and New Zealand\"
		projlocation = Environ("ARCGISHOME") & "\Coordinate Systems\Projected Coordinate Systems\National Grids\Australia\"
		
		If pIMxDoc.ActiveView.Extent.XMin < 1000 Then
			spatialrefstring = geolocation & "Geocentric Datum of Australia 1994.prj"
			pMap.DistanceUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
		Else
			spatialrefstring = projlocation & "GDA 1994 MGA Zone 50.prj"
			pMap.DistanceUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMeters
		End If
		
		If modCommon_Functions.Fileexists(spatialrefstring) = False Then
			MsgBox6("Projection file " & spatialrefstring & " does not exist", MsgBoxStyle.Exclamation, "PROJECTION PARAMETER FAILED")
			Exit Sub
		End If
		
		pMap.DelayDrawing((True))
		
		Dim pSpatReffact As ESRI.ArcGIS.Geometry.ISpatialReferenceFactory2 = New ESRI.ArcGIS.Geometry.SpatialReferenceEnvironment()
		
		Dim pspatref As ESRI.ArcGIS.Geometry.ISpatialReference = pSpatReffact.CreateESRISpatialReferenceFromPRJFile(spatialrefstring)
		
		Dim pActView As ESRI.ArcGIS.Carto.IActiveView
		If pspatref IsNot Nothing Then
			pMap.SpatialReference = pspatref
			pActView = pMap
			pActView.Refresh()
		End If

		pMap.DelayDrawing((False))
		
	End Sub

End Class
