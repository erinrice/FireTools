' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_ModifyPly")> _ 
<VB6Object("CALM_FireTools.clsFire_ModifyPly")> _
Public Class clsFire_ModifyPly
	Implements IDisposable
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_ModifyPly' member isn't used anywhere in current application.

	'  Interfaces implemented by this command

	'  ArcMap objects referenced by this command
	
	Private mMxApp As ESRI.ArcGIS.ArcMapUI.IMxApplication
	Private mEditor As ESRI.ArcGIS.Editor.IEditor
	' UPGRADE_INFO (#0711): The 'm_datestring' member has been removed because it is referenced only by members that haven't found to be used in the current application.
	' EXCLUDED: Private m_datestring As String

	'  FIMT application objects referenced by this command
	
	'  Member variables controlled by this class
	
	Private mBitmap As Image
	
	' Constants
	
	' UPGRADE_INFO (#0701): The 'BITMAP_NAME' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Const BITMAP_NAME As String = "COPYTOPERIM"
	' UPGRADE_INFO (#0561): The 'c_ModuleFileName' symbol was defined without an explicit "As" clause.
	Private Const c_ModuleFileName As String = "modpoly"
	' UPGRADE_INFO (#0561): The 'TOOL_NAME' symbol was defined without an explicit "As" clause.
	Private Const TOOL_NAME As String = "Modify Fire Boundary"
	' UPGRADE_INFO (#0561): The 'TOOL_MSG' symbol was defined without an explicit "As" clause.
	Private Const TOOL_MSG As String = "Modify the Existing Fire Burn Boundary"
	
	' UPGRADE_INFO (#0561): The 'getDate' symbol was defined without an explicit "As" clause.
	Public Function getDate() As Object
		Dim mydate As Date = Now
		Dim mydatestring As String = ""
		' UPGRADE_INFO (#0561): The 'MyDay' symbol was defined without an explicit "As" clause.
		Dim MyDay As Object = DatePart("d", mydate)
		' UPGRADE_INFO (#0561): The 'MyMonth' symbol was defined without an explicit "As" clause.
		Dim MyMonth As Object = DatePart("m", mydate)
		Dim MyYear As String = DatePart("yyyy", mydate)
		mydatestring = MyDay & "/" & MyMonth & "/" & MyYear
		
		Return mydatestring
	End Function

	' UPGRADE_INFO (#0561): The 'getDateString' symbol was defined without an explicit "As" clause.
	Private Function getDateString() As Object
		'GrantB 10-09-2007
		Dim mydate As Date = Now
		' UPGRADE_INFO (#0701): The 'myhourint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myhourint As Short
		' UPGRADE_INFO (#0701): The 'myminuteint' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim myminuteint As Short
		Dim mydatestring As String = ""
		Dim mytimestring As String = ""
		' UPGRADE_INFO (#0701): The 'MyMonthtext' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim MyMonthtext As Object = Nothing
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
		Return myhour & ":" & myminute & " (" & MyDay & "/" & MyMonth & "/" & MyYear & ")"
	End Function

	' UPGRADE_INFO (#0711): The 'clearSelectableLayers' member has been removed because it is referenced only by members that haven't found to be used in the current application.
	' EXCLUDED: Private Sub clearSelectableLayers(ByRef pMap As ESRI.ArcGIS.Carto.IMap)
		' EXCLUDED: 'GrantB - changed 24-09-2007 - pMap local var was not valid
		' EXCLUDED: Dim player As ESRI.ArcGIS.Carto.ILayer
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		' EXCLUDED: Dim complayercount As Short
		' EXCLUDED: Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		' EXCLUDED: Dim count As Short
		
		' EXCLUDED: 'clear selectable layers
		' EXCLUDED: For count = 0 To pMap.LayerCount - 1
			' EXCLUDED: player = pMap.Layer(count)
			' EXCLUDED: If TypeOf player Is ESRI.ArcGIS.Carto.IGroupLayer Then
				' EXCLUDED: pcomplayer = pMap.Layer(count)
				' EXCLUDED: For complayercount = 0 To pcomplayer.Count - 1
					' EXCLUDED: If TypeOf pcomplayer.Layer(complayercount) Is ESRI.ArcGIS.Carto.IFeatureLayer Then
						' EXCLUDED: pFeatLayer = pcomplayer.Layer(complayercount)
						' EXCLUDED: pFeatLayer.Selectable = False
					' EXCLUDED: End If
				' EXCLUDED: Next
			' EXCLUDED: ElseIf TypeOf player Is ESRI.ArcGIS.Carto.IFeatureLayer Then ' EXCLUDED: 'edited 23/10/06
				' EXCLUDED: pFeatLayer = player
				' EXCLUDED: pFeatLayer.Selectable = False
			' EXCLUDED: End If
		' EXCLUDED: Next
	' EXCLUDED: End Sub

	' UPGRADE_INFO (#0561): The 'getTime' symbol was defined without an explicit "As" clause.
	Public Function getTime() As Object
		Dim mydate As Date = Now
		Dim mytimestring As String = ""
		' UPGRADE_INFO (#0561): The 'myhour' symbol was defined without an explicit "As" clause.
		Dim myhour As Object = DatePart("h", mydate)
		Dim myminute As String = ""
		If Len6(myhour) = 1 Then
			myhour = "0" & myhour
		End If
		myminute = DatePart("n", mydate)
		If Len6(myminute) = 1 Then
			myminute = "0" & myminute
		End If
		mytimestring = myhour & ":" & myminute
		
		Return mytimestring
	End Function

	Private Sub Class_Initialize_VB6()
		On Error Resume Next
28:
		'Set mBitmap = LoadResPicture(BITMAP_NAME, vbResBitmap)
	End Sub

	Private Sub Class_Terminate_VB6()
		On Error Resume Next
		
35:
		mBitmap = Nothing
36:
		mMxApp = Nothing
37:
		mEditor = Nothing
		
	End Sub

	'  ICommand interface methods
	
	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
		Get
			On Error Resume Next
45:
			Return TOOL_NAME
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
		Get
			On Error Resume Next
50:
			Return GetPictureHandle6(frmImages.modifypoly.Picture)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error Resume Next
55:
			Return TOOL_NAME
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error Resume Next
			'changed 28-08-2007 by FMS request - GrantB
60:
			Return "Modify Current Fire Boundary" ' "Modify Active Fire Boundary"
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
73:
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

            '			'  Enable if...
            '			'  Extension is enabled,
            '			'  Edit session is started,
            '			'  Have FIMT workspace,
            '			'  Perimeter layer is present and editable
            '			'  Something is selected in the current map.

            '			' TODO: Add your implementation here
            '			Dim bEnabled As Boolean = False

            '			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            '			If modFire_Tools.Fire_Poly_Exists() Then
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
            '113:
            Return True
	 	End Get
	End Property

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler

        If g_incidentStatus = False Then
            MsgBox(modExtension.IncidentMessage)
            Exit Sub
        End If
        If Not modFire_Tools.Fire_Poly_Exists() Then
            MsgBox("Fire Boundary Layer doesn't exist")
            Exit Sub

        End If
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDoc.FocusMap
		
		If pMap.SelectionCount < 1 Then
			MsgBox6("Please select a Fire Boundary polygon to modify")
			Exit Sub
		End If
		
		'find active fire poly
		Dim sFeat As ESRI.ArcGIS.Geodatabase.IFeature = Nothing
		Dim pEnumFeat As ESRI.ArcGIS.Geodatabase.IEnumFeature = pMap.FeatureSelection
		Dim pEnumFeatureSetup As ESRI.ArcGIS.Geodatabase.IEnumFeatureSetup = pEnumFeat
		pEnumFeatureSetup.AllFields = True
		Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature = pEnumFeat.Next()
		Dim fID As Integer
		Do While (pFeat IsNot Nothing)
			fID = pFeat.Fields.FindField("Eff_To")
			If fID > 0 Then
				If GetDefaultMember6(pFeat.Value(fID)) Like "Active" Then
					If pFeat.Class.AliasName Like "Fire_Boundary" Then
						If sFeat Is Nothing Then
							sFeat = pFeat
						Else
							MsgBox6("Please select only one active Fire Boundary to modify")
							Exit Sub
						End If
					End If
				End If
			End If
			pFeat = pEnumFeat.Next()
		Loop

        If sFeat Is Nothing Then
            MsgBox6("Please select at least one active Fire Boundary to modify")
            Exit Sub
        End If
        Dim myForm As New frmFire_Poly

		'get the capture details
		' UPGRADE_INFO (#0561): The 'fb' symbol was defined without an explicit "As" clause.
        Dim fb As Object = myForm.Feature_type_cb.Locked
		' UPGRADE_INFO (#0701): The 'cm' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim cm As Boolean
		' UPGRADE_INFO (#0561): The 'txEff' symbol was defined without an explicit "As" clause.
		Dim txEff As Object = Nothing
		' UPGRADE_INFO (#0561): The 'txDate' symbol was defined without an explicit "As" clause.
		Dim txDate As Object = Nothing
		' UPGRADE_INFO (#0561): The 'txTime' symbol was defined without an explicit "As" clause.
		Dim txTime As Object = Nothing
		' UPGRADE_INFO (#0561): The 'txCapt' symbol was defined without an explicit "As" clause.
        Dim txCapt As Object = Nothing

        Dim txRegionEx As String
        Dim txFireNum As String
        Dim txRegion As String
        Dim txUsername As String
        Dim txDept As String
        Dim txFeatureType As String

        txRegionEx = sFeat.Value(sFeat.Fields.FindField(g_REGIONEX_FIELD))
        txFireNum = sFeat.Value(sFeat.Fields.FindField(g_FIRE_NUM_FIELD))
        txRegion = sFeat.Value(sFeat.Fields.FindField(g_REGION_FIELD))
        txUsername = sFeat.Value(sFeat.Fields.FindField(g_USERNAME_FIELD))
        txDept = sFeat.Value(sFeat.Fields.FindField(g_USER_DEPT_FIELD))
        txFeatureType = sFeat.Value(sFeat.Fields.FindField(g_FEATURE_FIELD))

        myForm.Feature_type_cb.Text = txFeatureType
        myForm.txtRegion_Ex.Text = txRegionEx
        myForm.txtFireNumber.Text = txFireNum
        myForm.txtRegion.Text = txRegion
        myForm.txtUserName.Text = txUsername
        myForm.txtDept.Text = txDept


        Dim tfb As String = myForm.Feature_type_cb.Text
		txEff = GetDefaultMember6(getDateString())
		txDate = GetDefaultMember6(getDate())
		txTime = GetDefaultMember6(getTime())
		txCapt = GetDefaultMember6(sFeat.Value(sFeat.Fields.FindField("CAPT_METHOD")))
        myForm.txtDate.Text = txDate
        myForm.txtDate.BackColor = SystemColors.Menu
        myForm.txtTime.Text = txTime
        myForm.txtTime.BackColor = SystemColors.Menu
        myForm.cbCapt_Method.Text = txCapt
        myForm.cbCapt_Method.BackColor = SystemColors.Menu
        myForm.Feature_type_cb.BackColor = SystemColors.ScrollBar
        myForm.Feature_type_cb.Text = ""
        myForm.Feature_type_cb.Enabled = False
        myForm.Feature_type_cb.Locked = True
        myForm.ShowDialog()

        '        myForm.Show(VBRUN.FormShowConstants.vbModal)
        If myForm.Did_Cancel Then
            myForm.Feature_type_cb.BackColor = SystemColors.Menu
            myForm.Feature_type_cb.Enabled = True
            myForm.Feature_type_cb.Locked = fb
            myForm.Feature_type_cb.Text = tfb
            Exit Sub
        End If
        myForm.Feature_type_cb.BackColor = SystemColors.Menu
        myForm.Feature_type_cb.Enabled = True
        myForm.Feature_type_cb.Locked = fb
        myForm.Feature_type_cb.Text = tfb
        txDate = myForm.txtDate.Text
        txTime = myForm.txtTime.Text
        txCapt = myForm.cbCapt_Method.Text
		
		Incident_StartEditing()
		
		modFire_Tools.SetCurrentLayer("Fire Boundary", 0)
		
		mEditor.StartOperation()
		
		'retire the user selected feature
		Dim r_from As Short
		Dim r_to As Short = sFeat.Fields.FindField("Eff_to")
		r_from = sFeat.Fields.FindField("Eff_from")
		sFeat.Value(r_to) = txEff
		sFeat.Store()
		
		'set the layer
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		' UPGRADE_INFO (#0561): The 'c' symbol was defined without an explicit "As" clause.
        Dim c As Short
		Dim i As Short
		For c = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(c) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(c)
				For i = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(i).Name = "Fire Boundary" Then
						pFeatureLayer = pcomplayer.Layer(i)
					End If
				Next
			End If
		Next
		
		'create and populate new feature
		Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeatureLayer.FeatureClass
		Dim newFeat As ESRI.ArcGIS.Geodatabase.IFeature = pFeatureClass.CreateFeature()
		
		Dim f As Short
		For f = 2 To sFeat.Fields.FieldCount - 1
			If sFeat.Fields.Field(f).Name <> "SHAPE_Length" And sFeat.Fields.Field(f).Name <> "SHAPE_Area" Then
				newFeat.Value(f) = sFeat.Value(f)
				newFeat.Shape = sFeat.Shape
			End If
		Next
		
		newFeat.Value(r_from) = txEff
		newFeat.Value(r_to) = "Active"
		newFeat.Value(sFeat.Fields.FindField("CAPT_DATE")) = txDate
		newFeat.Value(sFeat.Fields.FindField("CAPT_TIME")) = txTime
		newFeat.Value(sFeat.Fields.FindField("CAPT_METHOD")) = txCapt
		newFeat.Value(sFeat.Fields.FindField("AUTH_INC_STATUS")) = "DRAFT"
		newFeat.Value(sFeat.Fields.FindField("AUTH_INC_AUTH")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_INC_NAME")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_INC_DATE")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_INC_TIME")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_STA_STATUS")) = "DRAFT"
		newFeat.Value(sFeat.Fields.FindField("AUTH_STA_AUTH")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_STA_NAME")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_STA_DATE")) = ""
		newFeat.Value(sFeat.Fields.FindField("AUTH_STA_TIME")) = ""
		newFeat.Store()
		
		mEditor.StopOperation(TOOL_NAME)
		
		pMap.ClearSelection()
		pMap.SelectFeature(pFeatureLayer, newFeat)
		
		'selecting the right tool
		Dim pEditTask As ESRI.ArcGIS.Editor.IEditTask
		Dim ctask As Short
		If mEditor.CurrentTask.Name <> "Reshape Feature" Then
			For ctask = 0 To mEditor.TaskCount - 1
				pEditTask = mEditor.Task(ctask)
				If pEditTask.Name = "Reshape Feature" Then
					mEditor.CurrentTask = pEditTask
					Exit For
				End If
			Next
		End If
		
		'Activate the sketch tool
		Dim dUID As New ESRI.ArcGIS.esriSystem.UIDClass
		Dim pCmdItem As ESRI.ArcGIS.Framework.ICommandItem
		dUID.Value = "esriEditor.SketchTool"
		pCmdItem = m_pApp.Document.CommandBars.Find(dUID)
		m_pApp.CurrentTool = pCmdItem
		
		pMxDoc.ActiveView.Refresh()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
		
		mEditor.AbortOperation()
	End Sub

	' UPGRADE_INFO (#0701): The 'ICommand_OnClick_OLD_VERSION' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Sub ICommand_OnClick_OLD_VERSION()
		' EXCLUDED: On Error GoTo ErrorHandler
		
		' EXCLUDED: Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		' EXCLUDED: Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDoc.FocusMap
		
		' EXCLUDED: clearSelectableLayers(pMap)
		
		' EXCLUDED: Incident_StartEditing()
		
		' EXCLUDED: modFire_Tools.SetCurrentLayer("Fire Boundary", 0)
		
		' EXCLUDED: Dim bInOperation As Boolean = False
		
		' EXCLUDED: Dim polyFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer = modFire_Tools.Get_Fire_Poly_Layer()
		
		' EXCLUDED: 'make the Fire Boundary Layer the only selectable layer
		' EXCLUDED: polyFeatLayer.Selectable = True
		
		' EXCLUDED: mEditor.StartOperation()
		
		' EXCLUDED: 'GrantB - 10-09-2007 - is this actually used ????????
		' EXCLUDED: '    Dim MyDateText As String
		' EXCLUDED: '    MyDateText = getDateString
		
		' EXCLUDED: 'need to choose whether or not to append
		' EXCLUDED: Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		
		' EXCLUDED: pQueryFilter.WhereClause = "Eff_To = 'Active'" ' EXCLUDED: 'is Null or Eff_To = ''" '"Eff_to = ''"
		
		' EXCLUDED: Dim myfeatures As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		' EXCLUDED: Dim myfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = polyFeatLayer.FeatureClass
		
		' EXCLUDED: myfeatures = myfeatclass.Search(pQueryFilter, False)
		
		' EXCLUDED: Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
		' EXCLUDED: Dim h As Short = 1
		' EXCLUDED: Dim r_from As Short
		' EXCLUDED: Dim r_to As Short
		' EXCLUDED: Dim r_area As Short
		
		' EXCLUDED: m_datestring = GetDefaultMember6(getDateString()) ' EXCLUDED: 'frmFire_Effective_From.txtDate.Text  'GrantB - Reason for having the workaround in Manual when Modify don't work
		' EXCLUDED: Dim y As Short
		' EXCLUDED: Dim j As Short
		' UPGRADE_INFO (#0701): The 'copyFeat' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim copyFeat As ESRI.ArcGIS.Geodatabase.IFeature
		' EXCLUDED: Dim newFeat As ESRI.ArcGIS.Geodatabase.IFeature
		' EXCLUDED: Dim shape_area As Double
		' EXCLUDED: Dim largest_area As Double = 0
		
		' EXCLUDED: For y = 1 To myfeatclass.FeatureCount(pQueryFilter)
			' EXCLUDED: pFeat = myfeatures.NextFeature()
			' EXCLUDED: 'Set copyfeat = New Feature
			' EXCLUDED: 'Set copyfeat = pFeat
			' EXCLUDED: r_to = pFeat.Fields.FindField("Eff_to")
			' EXCLUDED: r_from = pFeat.Fields.FindField("Eff_from")
			' EXCLUDED: r_area = pFeat.Fields.FindField("SHAPE_Area") ' EXCLUDED: 'GrantB - might not need this area shit
			' EXCLUDED: shape_area = GetDefaultMember6(pFeat.Value(r_area))
			' EXCLUDED: If shape_area > largest_area Then
				' EXCLUDED: largest_area = shape_area
			' EXCLUDED: End If
			' EXCLUDED: pFeat.Value(r_to) = m_datestring
			' EXCLUDED: pFeat.Store()
			
			' EXCLUDED: newFeat = myfeatclass.CreateFeature()
			' EXCLUDED: For j = 2 To pFeat.Fields.FieldCount - 1
				' EXCLUDED: If pFeat.Fields.Field(j).Name <> "SHAPE_Length" And pFeat.Fields.Field(j).Name <> "SHAPE_Area" Then
					' EXCLUDED: newFeat.Value(j) = pFeat.Value(j)
					' EXCLUDED: newFeat.Shape = pFeat.Shape
				' EXCLUDED: End If
			' EXCLUDED: Next
			' EXCLUDED: newFeat.Value(r_from) = m_datestring
			' EXCLUDED: newFeat.Value(r_to) = "Active"
			' EXCLUDED: 'GrantB 19-10-2007 - Enforce Authorisation fields
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_INC_STATUS")) = "DRAFT"
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_INC_AUTH")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_INC_NAME")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_INC_DATE")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_INC_TIME")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_STA_STATUS")) = "DRAFT"
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_STA_AUTH")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_STA_NAME")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_STA_DATE")) = ""
			' EXCLUDED: newFeat.Value(pFeat.Fields.FindField("AUTH_STA_TIME")) = ""
			' EXCLUDED: 'MsgBox newfeat.Value(2)
			' EXCLUDED: newFeat.Store()
		' EXCLUDED: Next
		
		' EXCLUDED: mEditor.StopOperation(TOOL_NAME)
		' EXCLUDED: '--------- hmmmm should i worry about selecting something here - the user should click on a feature to be edited
		' EXCLUDED: pQueryFilter.WhereClause = pQueryFilter.WhereClause & " AND SHAPE_Area = " & largest_area
		
		' EXCLUDED: Dim pFeatureSel As ESRI.ArcGIS.Carto.IFeatureSelection = polyFeatLayer
		' EXCLUDED: pFeatureSel.SelectFeatures(pQueryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, False)
		
		' EXCLUDED: Dim pSelEvents As ESRI.ArcGIS.Carto.ISelectionEvents = pMxDoc.FocusMap
		' EXCLUDED: 'Fire SelectionChanged event
		' EXCLUDED: pSelEvents.SelectionChanged()
		' EXCLUDED: '---------
		' EXCLUDED: 'selecting the right tool
		' EXCLUDED: Dim pEditTask As ESRI.ArcGIS.Editor.IEditTask
		' EXCLUDED: Dim ctask As Short
		' EXCLUDED: If mEditor.CurrentTask.Name <> "Reshape Feature" Then
			' EXCLUDED: For ctask = 0 To mEditor.TaskCount - 1
				' EXCLUDED: pEditTask = mEditor.Task(ctask)
				' EXCLUDED: If pEditTask.Name = "Reshape Feature" Then
					' EXCLUDED: mEditor.CurrentTask = pEditTask
					' EXCLUDED: Exit For
				' EXCLUDED: End If
			' EXCLUDED: Next
		' EXCLUDED: End If
		
		' EXCLUDED: 'Activate the sketch tool
		' EXCLUDED: Dim dUID As New ESRI.ArcGIS.esriSystem.UIDClass
		' EXCLUDED: Dim pCmdItem As ESRI.ArcGIS.Framework.ICommandItem
		' EXCLUDED: dUID.Value = "esriEditor.SketchTool"
		' EXCLUDED: pCmdItem = m_pApp.Document.CommandBars.Find(dUID)
		' EXCLUDED: m_pApp.CurrentTool = pCmdItem
		
		' EXCLUDED: pMxDoc.ActiveView.Refresh()
		
		' EXCLUDED: Exit Sub
		' EXCLUDED: ErrorHandler:
		' EXCLUDED: HandleError(True, "ICommand_OnClick " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
		' EXCLUDED: If bInOperation Then
			' EXCLUDED: mEditor.AbortOperation()
			' EXCLUDED: bInOperation = False
		' EXCLUDED: End If
	' EXCLUDED: End Sub

	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
		On Error GoTo ErrorHandler
		
		'  Save a reference to ArcMap
		
290:
		mMxApp = hook
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		'Get a reference to the editor extension
293:
		pid.Value = "esriEditor.Editor"
294:
		mEditor = m_pApp.FindExtensionByCLSID(pid)

		'  To do:  Add anything else you need to here.
		
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
