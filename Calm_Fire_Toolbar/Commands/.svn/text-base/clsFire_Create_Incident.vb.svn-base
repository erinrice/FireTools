' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Create_Incident")> _ 
<VB6Object("CALM_FireTools.clsFire_Create_Incident")> _
Public Class clsFire_Create_Incident
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Create_Incident' member isn't used anywhere in current application.

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
			Return "Create Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Create Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Create a Fire Incident Geodatabase"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Create a new Fire Incident Geodatabase"
			
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
			Return GetPictureHandle6(frmImages.wizard.Picture)
			
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
		CreateArcGISHOME_Environ()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
		
		If modGlobals.g_INCIDENT_LAYER_NAME <> "" Then
			MsgBox6("An Incident has already been created using this document, please start a new Map Document before creating or opening another Incident", MsgBoxStyle.Exclamation, "Incident Already Open")
			Exit Sub
		End If
		
		frmFire_Create_Incident.Show(VBRUN.FormShowConstants.vbModal)

	End Sub

	Public Sub CreateArcGISHOME_Environ()
		' UPGRADE_INFO (#05B1): The 'oEnvSystem' variable wasn't declared explicitly.
		Dim oEnvSystem As Object = Nothing
		Dim Drivelist As New Collection
		Dim d As Object = Nothing
		If Environ("ARCGISHOME") = "" Then
			oEnvSystem = New MySystem()
			Drivelist.Add(("b"))
			Drivelist.Add(("c"))
			Drivelist.Add(("d"))
			Drivelist.Add(("e"))
			Drivelist.Add(("f"))
			Drivelist.Add(("g"))
			Drivelist.Add(("h"))
			Drivelist.Add(("i"))
			Drivelist.Add(("j"))
			Drivelist.Add(("k"))
			Drivelist.Add(("l"))
			Drivelist.Add(("m"))
			Drivelist.Add(("n"))
			Drivelist.Add(("o"))
			Drivelist.Add(("p"))
			Drivelist.Add(("q"))
			Drivelist.Add(("r"))
			Drivelist.Add(("s"))
			Drivelist.Add(("t"))
			Drivelist.Add(("u"))
			Drivelist.Add(("v"))
			Drivelist.Add(("w"))
			Drivelist.Add(("x"))
			Drivelist.Add(("y"))
			Drivelist.Add(("z"))
			For Each d In Drivelist
				If modCommon_Functions.Direxists(d & ":\Program Files\ArcGIS\Coordinate Systems") Then
					oEnvSystem.Environment("ARCGISHOME") = "C:\Program Files\ArcGIS"
					Exit For
				ElseIf modCommon_Functions.Direxists(d & ":\ArcGIS\Coordinate Systems") Then
					oEnvSystem.Environment("ARCGISHOME") = "C:\ArcGIS"
					Exit For
				End If
			Next
		End If
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
