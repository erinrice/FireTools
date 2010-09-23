' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs

<ComClass(clsedit_task_fire_div.ClassId, clsedit_task_fire_div.InterfaceId, clsedit_task_fire_div.EventsId), _
ProgId("CALM_FireTools.clsedit_task_fire_div")> _
<VB6Object("CALM_FireTools.clsedit_task_fire_div")> _
Public Class clsedit_task_fire_div
    Implements Editor.IEditTask

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d6bbb3c6-4fc4-4728-a497-76476754c48b"
    Public Const InterfaceId As String = "4778b6a0-e838-4377-803c-76e6dd60fe14"
    Public Const EventsId As String = "7fbcb97b-b356-4dc8-8ae6-d733340659ea"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditTasks.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditTasks.Unregister(regKey)

    End Sub

#End Region
#End Region


    ' UPGRADE_INFO (#0501): The 'clsedit_task_fire_div' member isn't used anywhere in current application.

    '******************************************************************************************************************
    'Date: 06 / 06 / 06
    'Author: Nathan Eaton, CALM
    'Description: Specific Edit Task to be included in the Editor to capture Fire Assignment data
    'Called By
    'Calls
    'Accepts
    'Returns
    '******************************************************************************************************************

    Private m_pEditor As ESRI.ArcGIS.Editor.IEditor
    Private m_pEditSketch As ESRI.ArcGIS.Editor.IEditSketch
    Public m_pCurrentLayer As ESRI.ArcGIS.Carto.ILayer
    Public m_datestring As String = ""
    Private m_pEditLayers As ESRI.ArcGIS.Editor.IEditLayers

    Private Sub populateFormFirePoint(ByVal myForm As frmFire_Point)
        Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)

        Dim myIncidentInfo As clsFire_Incident_Info
        If TypeOf pExt Is clsFire_Incident_Info Then
            myIncidentInfo = pExt
        End If


        myForm.txtRegion_Ex.Text = myIncidentInfo.IncidentRegionEx
        myForm.txtFireNumber.Text = myIncidentInfo.IncidentNumber
        myForm.txtRegion.Text = myIncidentInfo.incidentRegion
        myForm.txtUserName.Text = getUserName()

        If modGlobals.g_UserDept = g_Fire_Dept.DEC Then
            myForm.txtDept.Text = "DEC"
        ElseIf modGlobals.g_UserDept = g_Fire_Dept.FESA Then
            myForm.txtDept.Text = "FESA"
        End If

        Dim mydate As Date = Now
        Dim mydatestring As String = ""
        ' UPGRADE_INFO (#0561): The 'MyDay' symbol was defined without an explicit "As" clause.
        Dim MyDay As Object = DatePart("d", mydate)
        ' UPGRADE_INFO (#0561): The 'MyMonth' symbol was defined without an explicit "As" clause.
        Dim MyMonth As Object = DatePart("m", mydate)
        Dim MyYear As String = DatePart("yyyy", mydate)
        If Len6(MyDay) = 1 Then
            MyDay = "0" & MyDay
        End If
        If Len6(MyMonth) = 1 Then
            MyMonth = "0" & MyMonth
        End If
        mydatestring = MyDay & "/" & MyMonth & "/" & MyYear

        'Populate the myForm form with incident and date details for data capture
        myForm.txtDate.Text = mydatestring

        'GrantB 13-11-2007 - changed again by FMS request to have time on form
        Dim myhour As Short = DatePart("h", mydate)
        Dim myminute As Short
        Dim mytimestring As String = ""
        If Len6(myhour) = 1 Then
            myhour = "0" & myhour
        End If
        myminute = DatePart("n", mydate)
        If Len6(myminute) = 1 Then
            myminute = "0" & myminute
        End If
        mytimestring = myhour & myminute

        myForm.txtTime.Text = mytimestring '"-"
        'myForm.txtTime.BackColor = &H80000000 'changed 31-08-2007 by FMS Request - GrantB
        'myForm.txtTime.Locked = True 'changed 31-08-2007 by FMS Request - GrantB
        myForm.Caption = "Fire Boundary Assignment Point"
        myForm.Feature_type_cb.Clear()
        myForm.Feature_type_cb.AddItem(("Division"))
        myForm.Feature_type_cb.AddItem(("Sector"))
        myForm.Feature_type_cb.ListIndex = 0

        '        myForm.Show(VBRUN.FormShowConstants.vbModal)

        myForm.ShowDialog()
    End Sub

    Public Sub IEditTask_Activate(ByVal Editor As ESRI.ArcGIS.Editor.IEditor, ByVal oldTask As ESRI.ArcGIS.Editor.IEditTask) Implements Editor.IEditTask.Activate
        'GrantB 12-09-2007 - Just making sure the tool will work
        '    Dim myDoc As esriArcMapUI.IMxDocument
        '    Dim myScale As Double
        '    Dim mySRef As ISpatialReference
        '    Set myDoc = modCommon_Functions.m_pApp.Document
        '    Set mySRef = Nothing
        '    Set mySRef = myDoc.ActiveView.FocusMap.SpatialReference
        '    If mySRef Is Nothing Then 'see modMAP_Production.LoadTemplate: spatial reference
        '        MsgBox "A suitable projection could not be found. Please load a projected layer from the Corporate Data menu or define a projection"
        '        Call IEditTask_Deactivate
        '        Exit Sub
        '    End If

        'Show the "Append to" form on activation of the edit task
        m_pEditor = Editor
        m_pEditLayers = m_pEditor
        m_pEditSketch = m_pEditor 'QI
        m_pCurrentLayer = m_pEditLayers.CurrentLayer

        m_pEditSketch.GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint

        'input effective times
        Dim mydate As Date = Now
        'Dim myhourint As Integer
        'Dim myminuteint As Integer
        'Dim MyMonthtext As String
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
        If Len6(MyDay) = 1 Then
            MyDay = "0" & MyDay
        End If
        If Len6(MyMonth) = 1 Then
            MyMonth = "0" & MyMonth
        End If
        MyYear = DatePart("yyyy", mydate)
        MyDatetext = myhour & ":" & myminute & " (" & MyDay & "/" & MyMonth & "/" & MyYear & ")"

        Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
        pQueryFilter.WhereClause = "Eff_to = 'Active'" '"Eff_to = ''"

        Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer = m_pEditLayers.CurrentLayer
        Dim myfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass = pFeatLayer.FeatureClass
        Dim myfeatures As ESRI.ArcGIS.Geodatabase.IFeatureCursor = myfeatclass.Search(pQueryFilter, False)

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

        If myForm.OKButtonClicked = False Then 'added 10/10/06 to check if cancel is clicked
            Incident_StopEditing()
            Exit Sub
        End If

        Dim r_from As Short
        Dim r_to As String = ""
        Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
        'Dim h As Integer
        '    h = 1
        If myForm.AppendToCurrent = True Then
            pQueryFilter.WhereClause = "[OBJECTID] =" & myfeatclass.FeatureCount(Nothing)
            myfeatures = myfeatclass.Search(pQueryFilter, False)
            pFeat = myfeatures.NextFeature()
            r_from = pFeat.Fields.FindField("Eff_From")
            m_datestring = GetDefaultMember6(pFeat.Value(r_from))
            '       pFeat.Store 'added 31-08-2007 - GrantB - why?
        Else
            Do
                pFeat = myfeatures.NextFeature()
                r_to = pFeat.Fields.FindField("Eff_to")
                pFeat.Value(r_to) = m_datestring
                pFeat.Store()
            Loop Until pFeat Is Nothing
        End If
    End Sub

    Public Sub IEditTask_Deactivate() Implements Editor.IEditTask.Deactivate
        'Remove Snapping settings
        Dim pSnapEnvironment As ESRI.ArcGIS.Editor.ISnapEnvironment = m_pEditor
        pSnapEnvironment.ClearSnapAgents()

        'selecting the right tool
        Dim pEditTask As ESRI.ArcGIS.Editor.IEditTask
        Dim taskcount As Short
        For taskcount = 0 To m_pEditor.TaskCount - 1
            pEditTask = m_pEditor.Task(taskcount)
            If pEditTask.Name = "Create New Feature" Then
                Exit For
            End If
        Next
    End Sub

    Public ReadOnly Property IEditTask_Name() As String Implements Editor.IEditTask.Name
        Get
            Return "CALM Fire Assignment"
        End Get
    End Property

    Private Sub IEditTask_OnDeleteSketch() Implements Editor.IEditTask.OnDeleteSketch

    End Sub

    Public Sub IEditTask_OnFinishSketch() Implements Editor.IEditTask.OnFinishSketch
        On Error GoTo errhandler

        Dim myForm As New frmFire_Point

        populateFormFirePoint(myForm)
        'if the user cancels the form exit the sub
        If myForm.Did_Cancel = True Then
            Exit Sub
        End If

        'Dim pSelEnv As ISelectionEnvironment
        'Dim pMxApp As IMxApplication
        'Dim pGeometry As IGeometry

10:
        m_pEditor.StartOperation()

        'append new shape to feature class
        Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers = m_pEditor
        Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer = pEditLayers.CurrentLayer
        Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature = pFeatLayer.FeatureClass.CreateFeature()
        Dim pMap As ESRI.ArcGIS.Carto.IMap
        Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView
        Dim feature_string As String = ""
        pFeat.Shape = m_pEditSketch.Geometry
        pMap = m_pEditor.Map
        pActiveView = pMap
        feature_string = myForm.Feature_type_cb.Text

        'hittest
        Dim pFeatLayer_fire_bound As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim layername_fire_bound As String = "Fire Boundary"
        Dim pfeatclass_fire_bound As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim count As Short
        Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
        Dim pcomplayercount As Short
        'Loop through all of the maps layers to find the desired one
        For count = 0 To pMap.LayerCount - 1
            If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
                pcomplayer = pMap.Layer(count)
                For pcomplayercount = 0 To pcomplayer.Count - 1
                    If pcomplayer.Layer(pcomplayercount).Name = layername_fire_bound Then
                        pFeatLayer_fire_bound = pcomplayer.Layer(pcomplayercount)
                    End If
                Next
            End If
        Next
        pfeatclass_fire_bound = pFeatLayer_fire_bound.FeatureClass

        'perform the hit test to see where the point was digitised and the line it was attached to
        'so an angle can be created to display the assignment perpendicular to the fire boundary
        ' UPGRADE_INFO (#0701): The 'n' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim n As Short
        Dim pfeat_poly As ESRI.ArcGIS.Geodatabase.IFeature
        Dim myhit As Boolean = False
        Dim myhittest As ESRI.ArcGIS.Geometry.IHitTest
        Dim mypoint As ESRI.ArcGIS.Geometry.IPoint
        Dim myhitpoint As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point()
        Dim hitdist As Double = 0
        Dim partind As Integer = 0
        Dim segind As Integer = 0
        Dim Isrightside As Boolean = False
        Dim mysegcol As ESRI.ArcGIS.Geometry.ISegmentCollection
        Dim deltax As Double
        Dim deltay As Double
        Dim hyp As Double
        Dim Angle As Short
        Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        Dim pLine As ESRI.ArcGIS.Geometry.ILine
        Dim pgeom As ESRI.ArcGIS.Geometry.IGeometry = m_pEditSketch.Geometry

        'Saeid: the line below should be commented out otherwise myhittest.HitTest
        ' will create false when the geometry has diiferent coordinate than arcmap
        'framework
        'pgeom.Project pMap.SpatialReference
        mypoint = pgeom
        pFeatureCursor = pfeatclass_fire_bound.Search(Nothing, False)
        pfeat_poly = pFeatureCursor.NextFeature()

        While pfeat_poly IsNot Nothing
            '    'this line is where the error occurs
            '    '************************
            '    Set pfeat_poly = pfeatclass_fire_bound.GetFeature(n)
            '    '**********************
            myhittest = pfeat_poly.Shape

            mysegcol = pfeat_poly.Shape

            myhit = myhittest.HitTest(mypoint, 2, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary, myhitpoint, hitdist, partind, segind, Isrightside)

            'MsgBox "mypoint=" & mypoint.X & "," & mypoint.y
            'MsgBox "myhitpoint=" & myhitpoint.X & "," & myhitpoint.y

            If myhit = True Then
                deltax = (mysegcol.Segment(segind).FromPoint.X - mysegcol.Segment(segind).ToPoint.X)
                deltay = (mysegcol.Segment(segind).FromPoint.Y - mysegcol.Segment(segind).ToPoint.Y)
                hyp = ((deltax) ^ 2 + (deltay) ^ 2) ^ 0.5

                pLine = New ESRI.ArcGIS.Geometry.Line()
                ' pLine = New VB6Line()
                pLine.PutCoords(mysegcol.Segment(segind).FromPoint, mysegcol.Segment(segind).ToPoint)

                Angle = getLineAngle(pLine)
                '        MsgBox "deltax = " & deltax & vbNewLine
                ' '        & "deltay = " & deltay & vbNewLine
                ' '        & "hyp = " & hyp, vbInformation, "Debug"
                '        If hyp = 0 Then
                '            Angle = 0
                '        Else
                '            If (deltax > 0 And deltay > 0) Or (deltax < 0 And deltay < 0) Then
                ''                MsgBox "deltax or deltay not 0", vbInformation, "Debug"
                '                Angle = 180 - (Acos(Abs(deltax) / hyp)) ' - 90) '(180 - (Acos(deltax)) / hyp) * 2))
                '            Else
                ''                MsgBox "deltax or y is 0", vbInformation, "Debug"
                ''                MsgBox "Abs(deltax) = " & Abs(deltax) & vbNewLine
                ' ''                & "Acos(Abs(deltax) = " & Acos(Abs(deltax)) & vbNewLine
                ' ''                & "hyp = " & hyp, vbInformation, "Debug"
                '                Angle = (Acos(Abs(deltax) / hyp)) + 180
                ''                MsgBox "Angle = " & Angle, vbInformation, "Debug"
                '            End If
                '        End If
            End If
            pfeat_poly = pFeatureCursor.NextFeature()
        End While

        Dim r1 As Short = pFeat.Fields.FindField("Angle")
        Dim r_from As Short = pFeat.Fields.FindField("Eff_From")
        Dim r_to As Short = pFeat.Fields.FindField("Eff_to")
        Dim r_feat As Short = pFeat.Fields.FindField("Feature")

        'attribute the created feature
        pFeat.Value(r1) = getLineAngle(pLine) 'Angle

        pFeat.Value(r_from) = m_datestring

        pFeat.Value(r_to) = "Active"

        feature_string = myForm.Feature_type_cb.Text

        pFeat.Value(r_feat) = feature_string

        Dim featfield As Short = pFeat.Fields.FindField(g_FEATURE_FIELD)
        pFeat.Value(featfield) = feature_string

        Dim regfield As Short = pFeat.Fields.FindField(g_REGION_FIELD)
        pFeat.Value(regfield) = myForm.txtRegion.Text

        Dim regexfield As Short = pFeat.Fields.FindField(g_REGIONEX_FIELD)
        pFeat.Value(regexfield) = myForm.txtRegion_Ex.Text

        Dim firenumfield As Short = pFeat.Fields.FindField(g_FIRE_NUM_FIELD)
        pFeat.Value(firenumfield) = myForm.txtFireNumber.Text

        Dim usernamefield As Short = pFeat.Fields.FindField(g_USERNAME_FIELD)
        pFeat.Value(usernamefield) = myForm.txtUserName.Text

        Dim userdeptfield As Short = pFeat.Fields.FindField(g_USER_DEPT_FIELD)
        pFeat.Value(userdeptfield) = myForm.txtDept.Text

        Dim captdatefield As Short = pFeat.Fields.FindField(g_CAPT_DATE_FIELD)
        pFeat.Value(captdatefield) = myForm.txtDate.Text

        Dim captmethodfield As Short = pFeat.Fields.FindField(g_CAPT_METHOD_FIELD)
        pFeat.Value(captmethodfield) = myForm.cbCapt_Method.Text

        pFeat.Store()

        pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography, Nothing, Nothing) 'call again to refresh selection

20:
        m_pEditor.StopOperation("CALM Fire Assignment")

        Exit Sub

errhandler:
        MsgBox6(Err.Description & " " & Err.Number)
    End Sub

    ' UPGRADE_INFO (#0701): The 'Acos' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Public Function Acos(ByVal X As Single) As Single
    ' EXCLUDED: '******************************************************************************************************************
    ' EXCLUDED: 'Date: 06 / 06 / 06
    ' EXCLUDED: 'Author: Nathan Eaton, CALM
    ' EXCLUDED: 'Description: ACOS maths function
    ' EXCLUDED: 'Called By
    ' EXCLUDED: 'Calls
    ' EXCLUDED: 'Accepts
    ' EXCLUDED: 'Returns
    ' EXCLUDED: '******************************************************************************************************************
    ' EXCLUDED: '    MsgBox "start Acos", vbInformation, "Debug"
    ' EXCLUDED: '    MsgBox "-X = " & -X & vbNewLine
    ' EXCLUDED: ' & "Sqr(-X * X + 1) = " & Sqr(-X * X + 1), vbInformation, "Debug"
    ' EXCLUDED: '    If X = 1 Then
    ' EXCLUDED: '        Acos = 0
    ' EXCLUDED: '    Else
    ' EXCLUDED: Acos = Math.Atan(-X / Math.Sqrt(-X * X + 1)) + 2 * Math.Atan(1)
    ' EXCLUDED: Return Acos * (180 / 3.141592654)
    ' EXCLUDED: '    endif
    ' EXCLUDED: '    MsgBox "Acos = " & Acos, vbInformation, "Debug"
    ' EXCLUDED: '    MsgBox "end Acos", vbInformation, "Debug"

    ' EXCLUDED: End Function

    ' UPGRADE_INFO (#0551): The 'pLine' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    Private Function getLineAngle(ByRef pLine As ESRI.ArcGIS.Geometry.ILine) As Single
        Dim myAngle As Double
        Dim Pi As Double = 4 * Math.Atan(1)
        myAngle = (180 * pLine.Angle) / Pi

        If Abs6(myAngle) < 90 Then
            myAngle = 180 - myAngle
        ElseIf Abs6(myAngle) > 90 Then
            myAngle = 180 - myAngle
        Else
            myAngle = 90
        End If

        Return Int(Abs6(myAngle))

    End Function

    ' UPGRADE_INFO (#0701): The 'SetCurrentLayer' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private Sub SetCurrentLayer(ByRef layername As String, ByRef SubType As Integer)
    ' EXCLUDED: 'Subroutine to set the current layer based on a name and subtype index
    ' EXCLUDED: Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers = m_pEditor
    ' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    ' EXCLUDED: Dim pMap As ESRI.ArcGIS.Carto.IMap = m_pEditor.Map
    ' EXCLUDED: Dim count As Short
    ' EXCLUDED: 'Loop through all of the maps layers to find the desired one
    ' EXCLUDED: For count = 0 To pMap.LayerCount - 1
    ' EXCLUDED: If pMap.Layer(count).Name = layername Then
    ' EXCLUDED: 'Make sure the layer is editable
    ' EXCLUDED: If pEditLayers.IsEditable(pMap.Layer(count)) Then
    ' EXCLUDED: pFeatLayer = pMap.Layer(count)
    ' EXCLUDED: pEditLayers.SetCurrentLayer(pFeatLayer, SubType)
    ' EXCLUDED: 'MsgBox "set the layer"
    ' EXCLUDED: Exit Sub
    ' EXCLUDED: Else
    ' EXCLUDED: MsgBox6("This layer is not editable")
    ' EXCLUDED: Exit Sub
    ' EXCLUDED: End If
    ' EXCLUDED: End If
    ' EXCLUDED: Next
    ' EXCLUDED: End Sub

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
