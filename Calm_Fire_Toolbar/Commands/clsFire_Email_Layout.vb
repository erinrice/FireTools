' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Email_Layout")> _ 
<VB6Object("CALM_FireTools.clsFire_Email_Layout")> _
Public Class clsFire_Email_Layout
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Email_Layout' member isn't used anywhere in current application.

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
			
			bEnabled = Incident_Workspace_There()
			
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
			Return "Email_Layout"
			
ErrorHandler:
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Layout"
			
ErrorHandler:
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email Layout Jpeg"
			
ErrorHandler:
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email Layout Jpeg"
			
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
			Return GetPictureHandle6(frmImages.email_layout.Picture)
			
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
		'(sNames() As String, sLayers() As String, iSelCount As Integer)
		' From the ESRI Forum - hasn't been implemented yet.
		
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
		
		Dim fsoSysObj As VB6FileSystemObject
		If modCommon_Functions.Direxists(foldername) = False Then
			fsoSysObj = New VB6FileSystemObject()
			fsoSysObj.CreateFolder((foldername))
		End If
		
		Dim filename As String = foldername & myIncidentInfo.IncidentRegionEx & myIncidentInfo.IncidentNumber & "_" & mydatestring & "_" & mytimestring & "_Layout.jpg"
		modFire_Tools.ExportLayout_jpg(filename)

		Dim objOL As New Outlook.ApplicationClass
		Dim objMail As Outlook.MailItem
		objOL = New Outlook.Application()
		objMail = objOL.CreateItem(Outlook.OlItemType.olMailItem)
		Dim sIncNo As String = GetIncNo
		
		modFire_Tools.Update_Log_Email(filename, MyDay & "/" & MyMonth & "/" & MyYear, myhour & "/" & myminute, "Layout")
		
		Dim k As Short
		Dim emailstring As String = ""
		With objMail
			For k = 1 To GetEmails().Count()
				' UPGRADE_INFO (#0571): String concatenation inside a loop. Consider declaring the 'emailstring' variable as a StringBuilder6 object.
				emailstring = emailstring & GetDefaultMember6(GetEmails()(k)) & "; "
			Next
			
			.To = emailstring
			.Subject = myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber & " Layout - " & mydatestring & "_" & mytimestring
			.Body = "Please find attached data from incident " & myIncidentInfo.IncidentRegionEx & "::" & myIncidentInfo.IncidentNumber & Chr6(13) & Chr6(13) & "Regards, " & modGlobals.g_Username & Chr6(13) & Chr6(13) & "The Dept. of Environment and Conservation does not guarantee that this map is without flaw of any kind and disclaims" & Chr6(13) & "all liability for any errors, loss or other consequence which may arise from relying on any information depicted" & Chr6(13) & Chr6(13)
		End With
		
		objMail.Attachments.Add(filename)
		
		objMail.Display()
		
		objMail = Nothing
		objOL = Nothing
		
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
