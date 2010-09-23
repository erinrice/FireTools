' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Email_Boundary")> _ 
<VB6Object("CALM_FireTools.clsFire_Email_Boundary")> _
Public Class clsFire_Email_Boundary
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Email_Boundary' member isn't used anywhere in current application.

	'changed 28-08-2007 by FMS request - GrantB
	'commented out to test impact on app
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Dim bEnabled As Boolean = False
			
			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
			If g_Username = "" Then
				bEnabled = False
			End If
			If modFire_Tools.Fire_Poly_Exists() Then
				bEnabled = True 'modFire_Tools.FCIsEditable(modFire_Tools.Get_Fire_Line_Layer)
			Else
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
			Return "Email_Boundary"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Boundary Shapefile"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email Fire Boundary Shapefile"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email Fire Boundary Shapefile"
			
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
			Return GetPictureHandle6(frmImages.email_boundary.Picture)
			
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
		
		CreateEmail()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub CreateEmail()
		' UPGRADE_INFO (#05B1): The 'GetIncNo' variable wasn't declared explicitly.
		Dim GetIncNo As Object = Nothing
		
		' UPGRADE_INFO (#0701): The 'peditor' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim peditor As ESRI.ArcGIS.Editor.IEditor
		' UPGRADE_INFO (#0701): The 'pid' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		
		pFeatLayer = modFire_Tools.Get_Fire_Poly_Layer()
		
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeatLayer.FeatureClass
		
		'Create a scratch workspace factory to use for the selection
		Dim pScratchWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pScratchWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IScratchWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.ScratchWorkspaceFactory()
		pScratchWorkspace = pScratchWorkspaceFactory.DefaultScratchWorkspace
		
		Dim pSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		
		pQueryFilter.WhereClause = "[OBJECTID] >-1"
		
		pSelSet = pFeatClass.Select(pQueryFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
		
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
		mydatestring = MyDay & MyMonth & MyYear
		mytimestring = myhour & myminute
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		Dim foldername As String = Replace(myIncidentInfo.IncidentWorkspace.PathName, modCommon_Functions.Extract_filename(myIncidentInfo.IncidentWorkspace.PathName), "") & myIncidentInfo.IncidentRegionEx & myIncidentInfo.IncidentNumber & "_Output\"
		Dim fsoSysObj As VB6FileSystemObject = New VB6FileSystemObject()
		
		If modCommon_Functions.Direxists(foldername) = False Then
			fsoSysObj.CreateFolder((foldername))
		End If
		
		Dim filename As String = myIncidentInfo.IncidentRegionEx & myIncidentInfo.IncidentNumber & "_" & mydatestring & "_" & mytimestring & "_Boundary"
		
		Dim Outpath As String = foldername
		
		Dim extcol As New Collection
		extcol.Add((".dbf"))
		extcol.Add((".prj"))
		extcol.Add((".sbn"))
		extcol.Add((".sbx"))
		extcol.Add((".shp"))
		extcol.Add((".shx"))
		
		Dim extcolvar As Object = Nothing
		For Each extcolvar In extcol
			If Fileexists(Outpath & "\" & filename & extcolvar) Then
				fsoSysObj.DeleteFile((Outpath & "\" & filename & extcolvar))
			End If
		Next
		
		If Len6(Outpath) > 0 Then
			ExportToShapeFile(pFeatLayer, Outpath, filename & ".shp", pSelSet)
		End If
		
		Dim objOL As New Outlook.ApplicationClass
		Dim objMail As Outlook.MailItem
		objOL = New Outlook.Application()
		objMail = objOL.CreateItem(Outlook.OlItemType.olMailItem)
		Dim sIncNo As String = GetIncNo
		
		modFire_Tools.Update_Log_Email(filename & ".shp", MyDay & "/" & MyMonth & "/" & MyYear, myhour & "/" & myminute, "Incident Boundary Shapefile")
		
		Dim k As Short
		Dim emailstring As String = ""
		With objMail
			For k = 1 To GetEmails().Count()
				' UPGRADE_INFO (#0571): String concatenation inside a loop. Consider declaring the 'emailstring' variable as a StringBuilder6 object.
				emailstring = emailstring & GetDefaultMember6(GetEmails()(k)) & "; "
			Next
			.To = emailstring '"nathane@calm.wa.gov.au" 'sNames(0)
			.Subject = myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber & " Fire Boundary - " & mydatestring & "_" & mytimestring
			.Body = "Please find attached data from incident " & myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber & Chr6(13) & Chr6(13) & "Regards, " & modGlobals.g_Username & Chr6(13) & Chr6(13) & "The Dept. of Environment and Conservation does not guarantee that this map is without flaw of any kind and disclaims" & Chr6(13) & "all liability for any errors, loss or other consequence which may arise from relying on any information depicted" & Chr6(13) & Chr6(13)
		End With
		
		' Add email attachments
		If Environ("ARCGISHOME") = "" Then
			'MsgBox "Some Necessary Environment Variables are missing, try restarting ArcMap", vbExclamation, "Incorrect Configuration"
			objMail.Attachments.Add(Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\Fire_Boundary.lyr")
		End If
		
		objMail.Attachments.Add(Outpath & "\" & filename & ".dbf")
		objMail.Attachments.Add(Outpath & "\" & filename & ".shp")
		objMail.Attachments.Add(Outpath & "\" & filename & ".prj")
		objMail.Attachments.Add(Outpath & "\" & filename & ".shx")
		
		objMail.Display()
		
		objMail = Nothing
		objOL = Nothing
		
	End Sub

	' UPGRADE_INFO (#0701): The 'GetOutputFolder' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Function GetOutputFolder() As String
		' EXCLUDED: On Error GoTo ErrorHandler

		' EXCLUDED: Dim pGXDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
		' UPGRADE_INFO (#0701): The 'pObjectFilter' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pObjectFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		' EXCLUDED: Dim pEnumSelect As ESRI.ArcGIS.Catalog.IEnumGxObject
		' EXCLUDED: Dim pGXObject As ESRI.ArcGIS.Catalog.IGxObject
		
		' EXCLUDED: 136:
		' EXCLUDED: pGXDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
		
		' EXCLUDED: 139:
		' EXCLUDED: With pGXDialog
			' EXCLUDED: 140:
			' EXCLUDED: .RememberLocation = True
			' EXCLUDED: 142:
			' EXCLUDED: .AllowMultiSelect = False
			' EXCLUDED: 143:
			' EXCLUDED: .Title = "Destination Folder"
			' EXCLUDED: 144:
			' EXCLUDED: .ButtonCaption = "Open"
			' EXCLUDED: 145:
		' EXCLUDED: End With
		
		' EXCLUDED: 147:
		' EXCLUDED: If pGXDialog.DoModalOpen(0, pEnumSelect) Then
			' EXCLUDED: 148:
			' EXCLUDED: pGXObject = pEnumSelect.Next()
			' EXCLUDED: 149:
			' EXCLUDED: GetOutputFolder = pGXObject.FullName
			' EXCLUDED: 150:
		' EXCLUDED: Else
			' EXCLUDED: 151:
			' EXCLUDED: GetOutputFolder = ""
			' EXCLUDED: 152:
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
