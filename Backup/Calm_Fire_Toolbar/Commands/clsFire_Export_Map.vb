' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Export_Map")> _ 
<VB6Object("CALM_FireTools.clsFire_Export_Map")> _
Public Class clsFire_Export_Map
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Export_Map' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	Public papplication As ESRI.ArcGIS.Framework.IApplication
	Public m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
	Private WithEvents m_wpDoc As ESRI.ArcGIS.ArcMapUI.MxDocument
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
			Return "Fire_Export_Map"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Export Map to JPG"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Export Fire Map to JPG"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Export a Fire Map to a JPG file"
			
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
			Return GetPictureHandle6(frmImages.export_Map.Picture)
			
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
		papplication = hook
		m_pDoc = papplication.Document
		m_wpDoc = papplication.Document
		
		modCommon_Functions.m_pApp = papplication
		' TODO: Add your implementation here

		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
		
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView
		Dim mydoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pPageLayout As ESRI.ArcGIS.Carto.IPageLayout
		
		pActiveView = mydoc.ActiveView
		
		'set properties to help jpeg resolution issue from using JpegExport interface
		modFire_Tools.SetOutputQuality(pActiveView, 1)
		
		pPageLayout = mydoc.PageLayout
		pMap = mydoc.FocusMap
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
        Dim myForm As New frmFire_Export_Map

        myForm.txtRegion_Ex.Text = myIncidentInfo.IncidentRegionEx
        myForm.txtFireNumber.Text = myIncidentInfo.IncidentNumber
        myForm.txtUsername.Text = getUserName()
		
		If modGlobals.g_UserDept = g_Fire_Dept.DEC Then
            myForm.txtDept.Text = "DEC"
		ElseIf modGlobals.g_UserDept = g_Fire_Dept.FESA Then
            myForm.txtDept.Text = "FESA"
		End If
		
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
		mydatestring = MyDay & "/" & MyMonth & "/" & MyYear
		' GrantB 13-11-2007 - by FMS request
		' UPGRADE_INFO (#0561): The 'myFileDateString' symbol was defined without an explicit "As" clause.
		Dim myFileDateString As Object = MyYear & MyMonth & MyDay
		mytimestring = myhour & ":" & myminute
        myForm.txtDate.Text = mydatestring
        myForm.txtTime.Text = mytimestring
		
        myForm.txtMapScale.Text = Round6(pMap.MapScale, 0)
		
		pPageLayout.Page.Units = ESRI.ArcGIS.esriSystem.esriUnits.esriCentimeters
		Dim orientationstring As String = ""
		Dim Pagesize As String = ""
		Dim pageheight As Double
		Dim Pagewidth As Double
		pPageLayout.Page.QuerySize(Pagewidth, pageheight)
		
		If pPageLayout.Page.Orientation = 1 Then
			orientationstring = "Portrait"
			If pageheight > 110 Then
				Pagesize = "A0"
			ElseIf pageheight > 80 Then
				Pagesize = "A1"
			ElseIf pageheight > 50 Then
				Pagesize = "A2"
			ElseIf pageheight > 40 Then
				Pagesize = "A3"
			ElseIf pageheight > 20 Then
				Pagesize = "A4"
			Else
				Pagesize = "Unknown"
			End If
		ElseIf pPageLayout.Page.Orientation = 2 Then
			orientationstring = "Landscape"
			If Pagewidth > 110 Then
				Pagesize = "A0"
			ElseIf Pagewidth > 80 Then
				Pagesize = "A1"
			ElseIf Pagewidth > 50 Then
				Pagesize = "A2"
			ElseIf Pagewidth > 40 Then
				Pagesize = "A3"
			ElseIf Pagewidth > 20 Then
				Pagesize = "A4"
			Else
				Pagesize = "Unknown"
			End If
		End If
		
        myForm.txtPage_Size.Text = Pagesize
        myForm.txtOrientation.Text = orientationstring
        myForm.txtfolder.Text = Replace(myIncidentInfo.IncidentWorkspace.PathName, modCommon_Functions.Extract_filename(myIncidentInfo.IncidentWorkspace.PathName), "") & myIncidentInfo.IncidentRegionEx & myIncidentInfo.IncidentNumber & "_Output\"
        myForm.txtfilename.Text = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_" & Replace(myForm.cbMap_Type.Text, " ", "") & "_" & Pagesize & "_" & myFileDateString & "_" & Replace(mytimestring, ":", "") & ".jpg"
		
		If Pagesize = "A0" Then
            myForm.cbResolution.Text = "125"
		ElseIf Pagesize = "A1" Then
            myForm.cbResolution.Text = "175"
		ElseIf Pagesize = "A2" Then
            myForm.cbResolution.Text = "225"
		ElseIf Pagesize = "A3" Then
            myForm.cbResolution.Text = "300"
		ElseIf Pagesize = "A1" Then
            myForm.cbResolution.Text = "300"
		End If
		
        myForm.ShowDialog()
        If myForm.Did_Cancel = True Then
            Exit Sub
        End If
		
        'open loading form
        '        frmLoading.Show(VBRUN.FormShowConstants.vbModeless)
        Dim fsoSysObj As VB6FileSystemObject
        If modCommon_Functions.Direxists(ByVal6(myForm.txtfolder.Text)) = False Then
            fsoSysObj = New VB6FileSystemObject()
            fsoSysObj.CreateFolder((myForm.txtfolder.Text))
        End If
		
		' GrantB 13-11-2007 by FMS request
        ExportLayout_jpg(myForm.txtfolder.Text & myForm.txtfilename.Text, ByVal6(myForm.cbResolution.Text), ByVal6(myForm.cbQuality.Text))
		
		'GrantB 13-11-2007 - Removed by FMS request not to use log file
		'Dim pfeatworkspace As esriGeoDatabase.IFeatureWorkspace
		'Set pfeatworkspace = myIncidentInfo.IncidentWorkspace
		'Dim pTable As esriGeoDatabase.ITable
		'Set pTable = pfeatworkspace.OpenTable("Fire_Log")
		'Dim pRow As esriGeoDatabase.IRow
		'Set pRow = pTable.CreateRow
		'
		''Set Log Table properties
        'modCommon_Functions.SetTableValue pRow, "Log_Map_Type", myForm .cbMap_Type.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Map_Title", myForm .txtMap_Title.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Comments", myForm .txtComments.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Map_Resolution", myForm .cbResolution.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Map_Scale", myForm .txtMapScale.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Page_Size", myForm .txtPage_Size.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Page_Orientation", myForm .txtOrientation.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Fire_ID", myForm .txtRegion_Ex.Text & "::" & myForm .txtFireNumber.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Username", myForm .txtUserName.Text
        'modCommon_Functions.SetTableValue pRow, "Log_User_Dept", myForm .txtDept.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Date", myForm .txtDate.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Time", myForm .txtTime.Text
        'modCommon_Functions.SetTableValue pRow, "Log_Filename", myForm .txtfolder & myForm .txtfilename
		'modCommon_Functions.SetTableValue pRow, "Log_Type", "Export Map"
		'
		'modFire_Tools.Update_Log_Window
        '        frmLoading.Hide()
        MsgBox("Map exported")
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        'frmLoading.Hide()
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
