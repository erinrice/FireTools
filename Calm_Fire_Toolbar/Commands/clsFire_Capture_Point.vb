' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Capture_Point")> _ 
<VB6Object("CALM_FireTools.clsFire_Capture_Point")> _
Public Class clsFire_Capture_Point
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Capture_Point' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
		Get
            '			On Error GoTo ErrorHandler

            '			' TODO: Add your implementation here
            '			Dim bEnabled As Boolean = False

            '			bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            '			If modFire_Tools.Fire_Point_Exists() Then
            '				bEnabled = True 'odFire_Tools.FCIsEditable(modFire_Tools.Get_Fire_Point_Layer)
            '			Else
            '				bEnabled = False
            '			End If
            '			If g_Username = "" Then
            '				bEnabled = False
            '			End If

            '			Return bEnabled

            'ErrorHandler:
            '			HandleError(True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
            Return True
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
			Return "Fire Point"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Fire Point"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Create Fire Points"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Create a Fire Point feature"
			
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
			Return GetPictureHandle6(frmImages.new_point.Picture)
			
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
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		
		On Error GoTo ErrorHandler
        If g_incidentStatus = False Then
            MsgBox(modExtension.IncidentMessage)
            Exit Sub
        End If
        If Not modFire_Tools.Fire_Point_Exists() Then
            MsgBox("Fire Point Layer doesn't Exist")
            Exit Sub

        End If
		Dim peditor As ESRI.ArcGIS.Editor.IEditor
		Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim layername As String = "Fire Point"
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap = thisdocument.FocusMap
		' UPGRADE_INFO (#0701): The 'count' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim count As Short
		' UPGRADE_INFO (#0701): The 'complayercount' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim complayercount As Short
		' UPGRADE_INFO (#0701): The 'pcomplayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer

		pFeatLayer = modFire_Tools.Get_Fire_Point_Layer()
		
		'Get a reference to the editor extension
		pid.Value = "esriEditor.Editor"
		peditor = m_pApp.FindExtensionByCLSID(pid)
		
		'Create an edit operation enabling undo/redo
		modFire_Tools.Incident_StartEditing()
		
		Dim pdoc As ESRI.ArcGIS.Framework.IDocument = modCommon_Functions.m_pApp.Document
		
		modFire_Tools.SetCurrentLayer(layername, 0)
		
		'selecting the right tool
		Dim pEditTask As ESRI.ArcGIS.Editor.IEditTask
		Dim ctask As Short
		For ctask = 0 To peditor.TaskCount - 1
			pEditTask = peditor.Task(ctask)
			If pEditTask.Name = "CALM Fire Point" Then
				peditor.CurrentTask = pEditTask
				Exit For
			End If
		Next

        ' Saeid: i removed this to the task_fire_point
        '  Dim myForm As New frmFire_Point
        ' myForm.populatePointCB()
		
		Dim dUID As New ESRI.ArcGIS.esriSystem.UIDClass
		Dim pCmdItem As ESRI.ArcGIS.Framework.ICommandItem
		dUID.Value = "esriEditor.SketchTool"
		pCmdItem = m_pApp.Document.CommandBars.Find(dUID)
		m_pApp.CurrentTool = pCmdItem
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
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
