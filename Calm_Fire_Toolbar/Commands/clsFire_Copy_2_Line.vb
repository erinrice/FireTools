' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Copy_2_Line")> _ 
<VB6Object("CALM_FireTools.clsFire_Copy_2_Line")> _
Public Class clsFire_Copy_2_Line
	Implements IDisposable
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Copy_2_Line' member isn't used anywhere in current application.

	'  Interfaces implemented by this command
	
	'  ArcMap objects referenced by this command
	
	Private mMxApp As ESRI.ArcGIS.ArcMapUI.IMxApplication
	Private mEditor As ESRI.ArcGIS.Editor.IEditor
	Private m_datestring As String

	'  FIMT application objects referenced by this command
	
	'  Member variables controlled by this class
	
	Private mBitmap As Image
	
	' Constants
	
	' UPGRADE_INFO (#0701): The 'BITMAP_NAME' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Const BITMAP_NAME As String = "COPYTOPERIM"
	' UPGRADE_INFO (#0561): The 'c_ModuleFileName' symbol was defined without an explicit "As" clause.
	Private Const c_ModuleFileName As String = "clsFire_Copy_2_Line"
	' UPGRADE_INFO (#0561): The 'TOOL_NAME' symbol was defined without an explicit "As" clause.
	Private Const TOOL_NAME As String = "Copy to Fire Line Layer"
	' UPGRADE_INFO (#0561): The 'TOOL_MSG' symbol was defined without an explicit "As" clause.
	Private Const TOOL_MSG As String = "Copy selected lines from any layer into the Fire Line layer."
	
	Private Sub Class_Initialize_VB6()
		On Error Resume Next
28:
		'Set mBitmap = LoadResPicture(BITMAP_NAME, vbResBitmap)
	End Sub

	Private Sub Class_Terminate_VB6()
		On Error Resume Next
		
34:
		mBitmap = Nothing
35:
		mMxApp = Nothing
36:
		mEditor = Nothing
		
	End Sub

	'  ICommand interface methods
	
	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
		Get
			On Error Resume Next
44:
			Return TOOL_NAME
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
		Get
			On Error Resume Next
49:
			Return GetPictureHandle6(frmImages.copyline.Picture)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error Resume Next
54:
			Return TOOL_NAME
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error Resume Next
59:
			Return "Copy Line(s) to Fire Line Layer"
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
		Get
			' not used
	 	End Get
	End Property

	Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
		Get
			' not used
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error Resume Next
72:
			Return TOOL_MSG
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
		Get
			On Error Resume Next
77:
			'ICommand_Category = TOOL_CATEGORY
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
		Get
			' not used
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
            '			On Error GoTo ErrorHandler

            '			' TODO: Add your implementation here
            '			Dim bEnabled As Boolean = False

            '			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            '			If modFire_Tools.Fire_Line_Exists() Then
            '				bEnabled = True
            '			Else
            '				bEnabled = False
            '			End If
            '			If g_Username = "" Then
            '				bEnabled = False
            '			End If

            '			Return bEnabled

            'ErrorHandler:
            '			HandleError(True, "ICommand_Enabled " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
            '115:
            Return True
	 	End Get
	End Property

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
        If g_incidentStatus = False Then
            MsgBox(modExtension.IncidentMessage)
            Exit Sub
        End If
        If Not modFire_Tools.Fire_Line_Exists() Then
            MsgBox("Fire Line Layer doesn't exist")
            Exit Sub

        End If

		Incident_StartEditing()
		'  Copy the selected features into the perimeter polygon layer, and
		'  create corresponding lines in the sector assignment layer.
		
		Dim bInOperation As Boolean
125:
		bInOperation = False
		
		'  Get the perimeter polygon layer...
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
132:
		pFeatureLayer = modFire_Tools.Get_Fire_Line_Layer()
150:
		mEditor.StartOperation()
		
		Dim mydate As Date = Now
		' UPGRADE_INFO (#0701): The 'myhourint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myhourint As Short
		' UPGRADE_INFO (#0701): The 'myminuteint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myminuteint As Short
		Dim mydatestring As String = ""
		Dim mytimestring As String = ""
		' UPGRADE_INFO (#0701): The 'MyMonthtext' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim MyMonthtext As Object = Nothing
		' UPGRADE_INFO (#0561): The 'MyDatetext' symbol was defined without an explicit "As" clause.
		Dim MyDatetext As Object = Nothing
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
		MyMonth = DatePart("m", mydate)
		MyYear = DatePart("yyyy", mydate)
		mydatestring = MyDay & "/" & MyMonth & "/" & MyYear
		mytimestring = myhour & ":" & myminute
		MyDatetext = myhour & ":" & myminute & " (" & MyDay & "/" & MyMonth & "/" & MyYear & ")"
		
		'need to choose whether or not to append
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		
		pQueryFilter.WhereClause = "Eff_To = 'Active'" 'is Null or Eff_To = ''" '"Eff_to = ''"
		
		Dim myfeatures As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		Dim myfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeatureLayer.FeatureClass
		myfeatures = myfeatclass.Search(pQueryFilter, False)
		
		Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
		Dim h As Short = 1
		Dim r_from As Short
		Dim r_to As Short

        Dim myForm As New frmFire_Effective_From

		If myfeatclass.FeatureCount(Nothing) = 0 Then
            myForm.ckbAppend.Value = 2
            myForm.ckbAppend.Enabled = False
            myForm.AppendToCurrent = False
		Else
            myForm.ckbAppend.Value = 1
            myForm.ckbAppend.Enabled = True
		End If
		
        myForm.txtDate.Text = MyDatetext
        '        myForm.Show(VBRUN.FormShowConstants.vbModal)
        myForm.ShowDialog()
        m_datestring = myForm.txtDate.Text
		
        If myForm.OKButtonClicked = False Then 'added 29/9/06 to check if cancel button pressed
            mEditor.StopOperation(TOOL_NAME)
            Exit Sub
        End If
		
		Dim y As Short
        If myForm.AppendToCurrent = True Then
            pQueryFilter.WhereClause = "[Eff_To] = 'Active'" 'changed 19/09/06
            myfeatures = myfeatclass.Search(pQueryFilter, False)
            pFeat = myfeatures.NextFeature()
            If pFeat Is Nothing Then 'added 19/09/06 to check there are active features
                MsgBox6("There are no active features in the Fire_Point feature class", MsgBoxStyle.Information, "Debug")
                mEditor.StopOperation(TOOL_NAME)
                Exit Sub
            End If
            r_from = pFeat.Fields.FindField("Eff_From")
            m_datestring = GetDefaultMember6(pFeat.Value(r_from))
        Else
            For y = 1 To myfeatclass.FeatureCount(pQueryFilter)
                pFeat = myfeatures.NextFeature()
                r_to = pFeat.Fields.FindField("Eff_to")
                pFeat.Value(r_to) = m_datestring
                pFeat.Store()
            Next
        End If
		'end append
		
151:
		bInOperation = True
		Dim pNewPerimFeatSet As ESRI.ArcGIS.esriSystem.ISet
153:
		pNewPerimFeatSet = CopySelFeatures(mMxApp, pFeatureLayer)
		
		If pNewPerimFeatSet.Count = 0 Then
			MsgBox6("There are no suitable features to copy into the layer", MsgBoxStyle.Critical, "Invalid Feature(s)")
			mEditor.StopOperation(TOOL_NAME)
			Exit Sub
		End If
		
		Dim pNewPerimFeat As ESRI.ArcGIS.Geodatabase.IFeature
		' UPGRADE_INFO (#0701): The 'pNewSectorSet' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pNewSectorSet As ESRI.ArcGIS.esriSystem.ISet
		' UPGRADE_INFO (#0701): The 'pNewSectorFeat' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pNewSectorFeat As ESRI.ArcGIS.Geodatabase.IFeature
		' UPGRADE_INFO (#0701): The 'SectorCount' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim SectorCount As Integer
		
168:
		pNewPerimFeatSet.Reset()
169:
		pNewPerimFeat = pNewPerimFeatSet.Next()
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If

        Dim myFormLine As New frmFire_Line
        myFormLine.txtRegion_Ex.Text = myIncidentInfo.IncidentRegionEx
        myFormLine.txtFireNumber.Text = myIncidentInfo.IncidentNumber
        myFormLine.txtRegion.Text = myIncidentInfo.incidentRegion
        myFormLine.txtUserName.Text = getUserName()
		
		If modGlobals.g_UserDept = g_Fire_Dept.DEC Then
            myFormLine.txtDept.Text = "DEC"
		ElseIf modGlobals.g_UserDept = g_Fire_Dept.FESA Then
            myFormLine.txtDept.Text = "FESA"
		End If
		
        myFormLine.txtDate.Text = mydatestring
        myFormLine.txtTime.Text = mytimestring
        '        myFormLine.Show(VBRUN.FormShowConstants.vbModal)
        myFormLine.ShowDialog()

        If myFormLine.Did_Cancel = True Then
            mEditor.StopOperation(TOOL_NAME)
            Exit Sub
        End If
		
        Dim feature_string As String = myFormLine.Feature_type_cb.Text
		
		r_from = pNewPerimFeat.Fields.FindField("Eff_From")
		r_to = pNewPerimFeat.Fields.FindField("Eff_to")
		
		Dim featfield As Short = pNewPerimFeat.Fields.FindField(g_FEATURE_FIELD)
		
		Dim regfield As Short = pNewPerimFeat.Fields.FindField(g_REGION_FIELD)
		
		Dim regexfield As Short = pNewPerimFeat.Fields.FindField(g_REGIONEX_FIELD)
		
		Dim firenumfield As Short = pNewPerimFeat.Fields.FindField(g_FIRE_NUM_FIELD)
		
		Dim usernamefield As Short = pNewPerimFeat.Fields.FindField(g_USERNAME_FIELD)
		
		Dim userdeptfield As Short = pNewPerimFeat.Fields.FindField(g_USER_DEPT_FIELD)
		
		Dim captdatefield As Short = pNewPerimFeat.Fields.FindField(g_CAPT_DATE_FIELD)
		
		Dim capttimefield As Short = pNewPerimFeat.Fields.FindField(g_CAPT_TIME_FIELD)
		
		Dim captmethodfield As Short = pNewPerimFeat.Fields.FindField(g_CAPT_METHOD_FIELD)
		
        pNewPerimFeat.Value(captmethodfield) = myFormLine.cbCapt_Method.Text
        pNewPerimFeat.Value(capttimefield) = myFormLine.txtTime.Text
        pNewPerimFeat.Value(captdatefield) = myFormLine.txtDate.Text
        pNewPerimFeat.Value(userdeptfield) = myFormLine.txtDept.Text
        pNewPerimFeat.Value(usernamefield) = myFormLine.txtUserName.Text
        pNewPerimFeat.Value(regexfield) = myFormLine.txtRegion_Ex.Text
        pNewPerimFeat.Value(regfield) = myFormLine.txtRegion.Text
		pNewPerimFeat.Value(featfield) = feature_string
		pNewPerimFeat.Value(r_from) = m_datestring
        pNewPerimFeat.Value(firenumfield) = myFormLine.txtFireNumber.Text
		pNewPerimFeat.Value(r_to) = "Active"
		
179:
		pNewPerimFeat.Store()
		
190:
		Do Until pNewPerimFeat Is Nothing
			
            pNewPerimFeat.Value(captmethodfield) = myFormLine.cbCapt_Method.Text
            pNewPerimFeat.Value(capttimefield) = myFormLine.txtTime.Text
            pNewPerimFeat.Value(captdatefield) = myFormLine.txtDate.Text
            pNewPerimFeat.Value(userdeptfield) = myFormLine.txtDept.Text
            pNewPerimFeat.Value(usernamefield) = myFormLine.txtUserName.Text
            pNewPerimFeat.Value(regexfield) = myFormLine.txtRegion_Ex.Text
            pNewPerimFeat.Value(regfield) = myFormLine.txtRegion.Text
			pNewPerimFeat.Value(featfield) = feature_string
			pNewPerimFeat.Value(r_from) = m_datestring
            pNewPerimFeat.Value(firenumfield) = myFormLine.txtFireNumber.Text
			pNewPerimFeat.Value(r_to) = "Active"
			
			pNewPerimFeat.Store()
215:
			pNewPerimFeat = pNewPerimFeatSet.Next()
216:
		Loop
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		
		pMxDoc.ActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography, Nothing, Nothing) 'call again to refresh selection
		
217:
		mEditor.StopOperation(TOOL_NAME)
218:
		bInOperation = False

		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
224:
		If bInOperation Then
225:
			mEditor.AbortOperation()
226:
			bInOperation = False
227:
		End If
	End Sub

	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
		On Error GoTo ErrorHandler
		
		'  Save a reference to ArcMap
		
235:
		mMxApp = hook
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		mEditor = m_pApp.FindExtensionByCLSID(pid)
		
		Exit Sub
		
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

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
