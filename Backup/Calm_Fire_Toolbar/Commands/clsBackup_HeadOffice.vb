' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsBackup_HeadOffice")> _ 
<VB6Object("CALM_FireTools.clsBackup_HeadOffice")> _
Public Class clsBackup_HeadOffice
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsBackup_HeadOffice' member isn't used anywhere in current application.

	'-------------------------------------------------
	'Module was added to tool on 30-08-2007 by GrantB
	'Changes where requested by FMS
	'-------------------------------------------------

	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	'Const i_BACKUP_PATH As String = "\\calm-kens-54\RSD_Share\RSD_Share\Wildfire 200708\"            ' Constant for Head Office backup path  '\\calm-kens-54\RSD_Share\RSD_Share\Wildfire 200708

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
			Return "Backup_HeadOffice"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			'changed 28-08-2007 by FMS request - GrantB
			Return "Export to Head Office" ' "Backup Incident"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Backup Incident Geodatabase to Head Office"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Backup Incident Geodatabase to Head Office"
			
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
			Return GetPictureHandle6(frmImages.Saveimg.Picture)
			
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
		
		CreateBackup()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub CreateBackup()
		' UPGRADE_INFO (#05B1): The 'fsoSysObj' variable wasn't declared explicitly.
		Dim fsoSysObj As Object = Nothing
		'changed 28-08-2007 by FMS request - GrantB
		
		On Error GoTo ErrorHandler
		
		Incident_StopEditing() 'added 26/09/06
		
		Dim foldername As String = modExtension.GetConfigValue("Backup")
		'    foldername = i_BACKUP_PATH
		
		Dim msgSTR As String = ControlChars.NewLine & foldername & ControlChars.NewLine & "Please ensure the destination exists, write access is allowed to it and there is enough storage space available."
		
		If modCommon_Functions.Direxists(foldername) = False Then
			MsgBox6("The destination folder does not exist or is not accessible!" & msgSTR)
			Exit Sub
		End If
		
		' UPGRADE_INFO (#0701): The 'pid' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass
		' UPGRADE_INFO (#0701): The 'pWorkspaceedit' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pWorkspaceedit As ESRI.ArcGIS.Geodatabase.IWorkspaceEdit
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
		Dim myIncidentInfo As clsFire_Incident_Info
		If TypeOf pExt Is clsFire_Incident_Info Then
			myIncidentInfo = pExt
		End If
		
		Dim filename As String = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_Backup_" & GetDefaultMember6(getFileStamp()) & ".mdb" ' GrantB 13-11-2007 by FMS request
		
		'GrantB - 25-10-2007 - Not allowed to overwrite files
		Dim c As Short = 1
		While Fileexists(foldername & filename) ' GrantB 13-11-2007 by FMS request
			filename = myIncidentInfo.IncidentRegionEx & "_" & myIncidentInfo.IncidentNumber & "_Backup_" & GetDefaultMember6(getFileStamp()) & "(" & c & ").mdb"
			c += 1
		End While
		'    If Fileexists(foldername & filename) Then
		'        Set fsoSysObj = New FileSystemObject
		'        fsoSysObj.DeleteFile foldername & filename
		'    End If
		
		fsoSysObj = New VB6FileSystemObject()
		fsoSysObj.CopyFile(myIncidentInfo.IncidentWorkspace.PathName, foldername & filename)
		
		'GrantB - 17-10-2007 - changed wording
		MsgBox6("Export Successful!" & ControlChars.NewLine & "The Incident GeoDatabase was successfuly exported to Head Office." & ControlChars.NewLine & "Location: " & foldername & filename, MsgBoxStyle.Information, "Export to Head Office")
		Exit Sub
		
ErrorHandler:
		'GrantB - 17-10-2007 - changed wording
		MsgBox6("The export procedure may not have worked properly, you will have to do it manually!!" & ControlChars.NewLine & "Tip: Use the Export Geodatabase Tool until Head Office is available." & msgSTR, MsgBoxStyle.Critical, "Export Incomplete")
		
	End Sub

	' UPGRADE_INFO (#0561): The 'getFileStamp' symbol was defined without an explicit "As" clause.
	Private Function getFileStamp() As Object
		
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
		mydatestring = MyYear & MyMonth & MyDay 'MyDay & MyMonth & MyYear ' GrantB 13-11-2007 by FMS request
		mytimestring = myhour & myminute
		
		Return mydatestring & "_" & mytimestring
		
	End Function
	
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
