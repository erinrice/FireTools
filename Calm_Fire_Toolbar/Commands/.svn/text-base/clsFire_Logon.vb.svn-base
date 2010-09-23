' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports VB = Microsoft.VisualBasic

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Logon")> _ 
<VB6Object("CALM_FireTools.clsFire_Logon")> _
Public Class clsFire_Logon
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Logon' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	Private WithEvents m_wpDoc As ESRI.ArcGIS.ArcMapUI.MxDocument
	Private m_pEditor As ESRI.ArcGIS.Editor.IEditor
	' UPGRADE_ISSUE (#0338): Unable to translate next statement. Please fix this code manually.
    Dim WithEvents myeditor As Editor.Editor
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
            'Dim bEnabled As Boolean = False
            '         bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            '         Return bEnabled

            Return True
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
			Return "Logon"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Logon"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Logon the Fire Mapping Tool"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Logon the Fire Mapping Tool"
			
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
			Return GetPictureHandle6(frmImages.Logon.Picture)
			
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
		' UPGRADE_INFO (#05B1): The 'papplication' variable wasn't declared explicitly.
		Dim papplication As Object = Nothing
		On Error GoTo ErrorHandler
		
		' TODO: Add your implementation here
		modCommon_Functions.m_pApp = hook
		papplication = hook
		
		m_wpDoc = papplication.Document
		
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		myeditor = papplication.FindExtensionByCLSID(pid)
		m_pEditor = papplication.FindExtensionByCLSID(pid)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick

        If (Not ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)) Then
            MsgBox(modExtension.LoginMessage)
            Exit Sub
        End If

        'On Error GoTo ErrorHandler
        'Dim myForm As New frmFire_Logon
        'myForm.Show()
        frmFire_Logon.Show()

        Exit Sub
        'ErrorHandler
        'MsgBox Err.Description
        'HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
    End Sub


	' UPGRADE_INFO (#0701): The 'm_wpDoc_BeforeCloseDocument' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Function m_wpDoc_BeforeCloseDocument() As Boolean
		
		' EXCLUDED: 'ensure legend limiter gets turned off before closing - Leave out for the minute, as there will be problems with saving and all that jazz
		
		' EXCLUDED: '    Dim pUID As New UID
		' EXCLUDED: '    Dim i As Integer
		' EXCLUDED: '    Dim f As Integer
		' EXCLUDED: '    Dim activemap As IMap
		' EXCLUDED: '    pUID.Value = "CALM_FireTools.clsTrim_Unique"
		' EXCLUDED: '    Dim pCI As ICommandItem
		' EXCLUDED: '    Set pCI = m_pApp.Document.CommandBars.find(pUID)
		' EXCLUDED: '
		' EXCLUDED: '    If Not pCI Is Nothing Then
		' EXCLUDED: '        Dim pMyCommand1 As CALM_FireTools.clsTrim_Unique
		' EXCLUDED: '        Set pMyCommand1 = pCI.Command
		' EXCLUDED: '    Else
		' EXCLUDED: '        Exit Sub
		' EXCLUDED: '    End If
		' EXCLUDED: '
		' EXCLUDED: '    If pMyCommand1.m_checked = True Then
		' EXCLUDED: '        Set activemap = m_wpDoc.ActiveView.FocusMap
		' EXCLUDED: '        For i = 0 To m_pMXDoc.Maps.count - 1
		' EXCLUDED: '            Set m_pMXDoc.ActiveView = m_pMXDoc.Maps.Item(i)
		' EXCLUDED: '            modFire_Tools.FinishTrimUniqueValues
		' EXCLUDED: '        Next i
		' EXCLUDED: '
		' EXCLUDED: '        Set m_pMXDoc.ActiveView = activemap
		' EXCLUDED: '
		' EXCLUDED: '        For f = 1 To layers_added.count
		' EXCLUDED: '            layers_added.REMOVE (1)
		' EXCLUDED: '        Next f
		' EXCLUDED: '        m_pMXDoc.UpdateContents
		' EXCLUDED: '    End If
		
	' EXCLUDED: End Function

	' UPGRADE_INFO (#0701): The 'm_wpDoc_OpenDocument' member has been removed because it isn't used anywhere in current application.
    'Private Function m_wpDoc_OpenDocument() As Boolean




    ' EXCLUDED: 'If a document is opened with an incident in it, then recognise the incident porperties and set
    ' EXCLUDED: 'the incident info

    ' EXCLUDED: Dim layername As String = "Fire Boundary"
    ' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    ' EXCLUDED: Dim pMap As ESRI.ArcGIS.Carto.IMap
    ' EXCLUDED: Dim count As Short
    ' EXCLUDED: Dim complayercount As Short
    ' EXCLUDED: Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
    ' EXCLUDED: Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
    ' EXCLUDED: pMap = thisdocument.FocusMap
    ' EXCLUDED: Dim incidentName As String = ""
    ' EXCLUDED: Dim incidentRegion As String = ""

    ' EXCLUDED: 'Loop through all of the maps layers to find the desired one
    ' EXCLUDED: For count = 0 To pMap.LayerCount - 1
    ' EXCLUDED: If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer And pMap.Layer(count).Name Like "*::*" Then
    ' EXCLUDED: pcomplayer = pMap.Layer(count)
    ' EXCLUDED: incidentName = pMap.Layer(count).Name
    ' EXCLUDED: For complayercount = 0 To pcomplayer.Count - 1
    ' EXCLUDED: If pcomplayer.Layer(complayercount).Name = layername Then
    ' EXCLUDED: pFeatLayer = pcomplayer.Layer(complayercount)
    ' EXCLUDED: End If
    ' EXCLUDED: Next
    ' EXCLUDED: End If
    ' EXCLUDED: Next

    ' EXCLUDED: Dim mydataset As ESRI.ArcGIS.Geodatabase.IDataset
    ' EXCLUDED: Dim pEditor As ESRI.ArcGIS.Editor.IEditor
    ' EXCLUDED: Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
    ' EXCLUDED: Dim pExt As ESRI.ArcGIS.esriSystem.IExtension
    ' EXCLUDED: Dim myIncidentInfo As clsFire_Incident_Info
    ' EXCLUDED: If pFeatLayer IsNot Nothing Then

    ' EXCLUDED: 'Make sure something has been selected
    ' EXCLUDED: mydataset = pFeatLayer.FeatureClass

    ' EXCLUDED: 'Get a reference to the editor extension
    ' EXCLUDED: pid.Value = "esriEditor.Editor"
    ' EXCLUDED: pEditor = m_pApp.FindExtensionByCLSID(pid)
    ' EXCLUDED: pExt = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)

    ' EXCLUDED: If TypeOf pExt Is clsFire_Incident_Info Then
    ' EXCLUDED: myIncidentInfo = pExt
    ' EXCLUDED: End If

    ' EXCLUDED: myIncidentInfo.IncidentWorkspace = mydataset.Workspace
    ' EXCLUDED: myIncidentInfo.IncidentNumber = Replace(incidentName, VB.Left(incidentName, 5), "")
    ' EXCLUDED: myIncidentInfo.IncidentRegionEx = VB.Left(incidentName, 3)

    ' EXCLUDED: Select Case myIncidentInfo.IncidentRegionEx
    ' EXCLUDED: Case "GFD"
    ' EXCLUDED: incidentRegion = "Goldfields"
    ' EXCLUDED: Case "KMB"
    ' EXCLUDED: incidentRegion = "Kimberley"
    ' EXCLUDED: Case "EXM"
    ' EXCLUDED: incidentRegion = "Pilbara"
    ' EXCLUDED: Case "PIL"
    ' EXCLUDED: incidentRegion = "Pilbara"
    ' EXCLUDED: Case "ALB"
    ' EXCLUDED: incidentRegion = "Sth Coast"
    ' EXCLUDED: Case "ESP"
    ' EXCLUDED: incidentRegion = "Sth Coast"
    ' EXCLUDED: Case "BWD"
    ' EXCLUDED: incidentRegion = "Sth West"
    ' EXCLUDED: Case "WTN"
    ' EXCLUDED: incidentRegion = "Sth West"
    ' EXCLUDED: Case "PHL"
    ' EXCLUDED: incidentRegion = "Swan"
    ' EXCLUDED: Case "SCT"
    ' EXCLUDED: incidentRegion = "Swan"
    ' EXCLUDED: Case "DON"
    ' EXCLUDED: incidentRegion = "Warren"
    ' EXCLUDED: Case "FKL"
    ' EXCLUDED: incidentRegion = "Warren"
    ' EXCLUDED: Case "KTG"
    ' EXCLUDED: incidentRegion = "Wheatbelt"
    ' EXCLUDED: Case "MDN"
    ' EXCLUDED: incidentRegion = "Wheatbelt"
    ' EXCLUDED: Case "NGN"
    ' EXCLUDED: incidentRegion = "Wheatbelt"
    ' EXCLUDED: End Select

    ' EXCLUDED: myIncidentInfo.incidentRegion = incidentRegion

    ' EXCLUDED: End If

    'End Function

    Private Sub myeditor_OnStopEditing(ByVal save As Boolean) Handles myeditor.OnStopEditing
        'Ensure that no Fire Edit Task remains the selected edit task after editing has finished
        'otherwise the user will be prompted to "Append Shapes" when next starting an edit session

        Dim pEditTask As ESRI.ArcGIS.Editor.IEditTask
        Dim taskcount As Short

        For taskcount = 0 To m_pEditor.TaskCount - 1
            pEditTask = m_pEditor.Task(taskcount)
            If pEditTask.Name = "Create New Feature" Then
                Exit For
            End If
        Next

        m_pEditor.CurrentTask = pEditTask

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
