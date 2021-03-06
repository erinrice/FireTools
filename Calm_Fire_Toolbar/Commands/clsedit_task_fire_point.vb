﻿' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs

<ComClass(clsedit_task_fire_point.ClassId, clsedit_task_fire_point.InterfaceId, clsedit_task_fire_point.EventsId), _
ProgId("CALM_FireTools.clsedit_task_fire_point")> _
<VB6Object("CALM_FireTools.clsedit_task_fire_point")> _
Public Class clsedit_task_fire_point
    Implements Editor.IEditTask

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "bc8837fc-e216-4dd9-9145-7a40bf5d29af"
    Public Const InterfaceId As String = "22ba5b13-c257-4b4f-8d5e-c4b6b29f41d7"
    Public Const EventsId As String = "f89ddf0b-68e5-4613-ae57-c6c5a6af5a15"
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

    ' UPGRADE_INFO (#0501): The 'clsedit_task_fire_point' member isn't used anywhere in current application.

    '******************************************************************************************************************
    'Date: 06 / 06 / 06
    'Author: Nathan Eaton, CALM
    'Description: Edit task for capturing Fire Points
    'Called By
    'Calls
    'Accepts
    'Returns
    '******************************************************************************************************************

    Private m_pEditor As ESRI.ArcGIS.Editor.IEditor
    Private m_pEditSketch As ESRI.ArcGIS.Editor.IEditSketch
    Public m_pCurrentLayer As ESRI.ArcGIS.Carto.ILayer
    Private m_pEditLayers As ESRI.ArcGIS.Editor.IEditLayers
    Public m_datestring As String = ""

    Public Sub IEditTask_Activate(ByVal Editor As ESRI.ArcGIS.Editor.IEditor, ByVal oldTask As ESRI.ArcGIS.Editor.IEditTask) Implements Editor.IEditTask.Activate
        ' UPGRADE_INFO (#05B1): The 'r_to' variable wasn't declared explicitly.
        Dim r_to As Object = Nothing

        m_pEditor = Editor
        m_pEditLayers = m_pEditor
        m_pEditSketch = m_pEditor 'QI
        m_pCurrentLayer = m_pEditLayers.CurrentLayer

        'Set the geometry type of the edit sketch to polyline
        m_pEditSketch.GeometryType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint

        'input effective times
        Dim mydate As Date = Now
        ' UPGRADE_INFO (#0701): The 'myhourint' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim myhourint As Short
        ' UPGRADE_INFO (#0701): The 'myminuteint' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim myminuteint As Short
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
        If Len6(MyDay) = 1 Then
            MyDay = "0" & MyDay
        End If
        If Len6(MyMonth) = 1 Then
            MyMonth = "0" & MyMonth
        End If

        MyYear = DatePart("yyyy", mydate)
        MyDatetext = myhour & ":" & myminute & " (" & MyDay & "/" & MyMonth & "/" & MyYear & ")"

        Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()

        pQueryFilter.WhereClause = "Eff_To = 'Active'" 'is Null or Eff_To = ''" '"Eff_to = ''"

        Dim myfeatures As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        Dim myfeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer = m_pEditLayers.CurrentLayer
        myfeatclass = pFeatLayer.FeatureClass
        myfeatures = myfeatclass.Search(pQueryFilter, False)

        Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
        Dim h As Short = 1
        Dim r_from As Short

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

        If myForm.AppendToCurrent = True Then
            pQueryFilter.WhereClause = "[OBJECTID] =" & myfeatclass.FeatureCount(Nothing)
            myfeatures = myfeatclass.Search(pQueryFilter, False)
            pFeat = myfeatures.NextFeature()
            r_from = pFeat.Fields.FindField("Eff_From")
            m_datestring = GetDefaultMember6(pFeat.Value(r_from))
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
            Return "CALM Fire Point"

        End Get
    End Property

    Public Sub IEditTask_OnDeleteSketch() Implements Editor.IEditTask.OnDeleteSketch

    End Sub

    Public Sub IEditTask_OnFinishSketch() Implements Editor.IEditTask.OnFinishSketch

        ' UPGRADE_INFO (#0701): The 'pSelEnv' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim pSelEnv As ESRI.ArcGIS.Carto.ISelectionEnvironment
        ' UPGRADE_INFO (#0701): The 'pMxApp' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim pMxApp As ESRI.ArcGIS.ArcMapUI.IMxApplication
        Dim pMap As ESRI.ArcGIS.Carto.IMap = m_pEditor.Map
        ' UPGRADE_INFO (#0701): The 'pGeometry' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim pGeometry As ESRI.ArcGIS.Geometry.IGeometry
        Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView = pMap

        Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
        Dim myIncidentInfo As clsFire_Incident_Info

        If TypeOf pExt Is clsFire_Incident_Info Then
            myIncidentInfo = pExt
        End If
        Dim myForm As New frmFire_Point
        myForm.populatePointCB()
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
        MyMonth = DatePart("m", mydate)
        MyYear = DatePart("yyyy", mydate)
        mydatestring = MyDay & "/" & MyMonth & "/" & MyYear
        mytimestring = myhour & ":" & myminute
        myForm.txtDate.Text = mydatestring
        myForm.txtTime.Text = mytimestring
        myForm.Feature_type_cb.Text = "Fire Origin"
        '        myForm.Show(VBRUN.FormShowConstants.vbModal)
        myForm.ShowDialog()
        If myForm.Did_Cancel = True Then
            Exit Sub
        End If

        m_pEditor.StartOperation()

        'append new shape to feature class
        Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
        Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers = m_pEditor
        pFeatLayer = pEditLayers.CurrentLayer
        pFeat = pFeatLayer.FeatureClass.CreateFeature()
        pFeat.Shape = m_pEditSketch.Geometry
        pFeat.Store()

        Dim r_from As Short = pFeat.Fields.FindField("Eff_From")
        pFeat.Value(r_from) = m_datestring

        Dim r_to As Short = pFeat.Fields.FindField("Eff_to")
        pFeat.Value(r_to) = "Active"

        Dim feature_string As String = myForm.Feature_type_cb.Text

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

        Dim capttimefield As Short = pFeat.Fields.FindField(g_CAPT_TIME_FIELD)
        pFeat.Value(capttimefield) = myForm.txtTime.Text

        Dim captmethodfield As Short = pFeat.Fields.FindField(g_CAPT_METHOD_FIELD)
        pFeat.Value(captmethodfield) = myForm.cbCapt_Method.Text

        pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography, Nothing, Nothing) 'call again to refresh selection

        m_pEditor.StopOperation("CALM Fire Point")

    End Sub

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
