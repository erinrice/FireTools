' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmDriveChoice
' UPGRADE_INFO (#0511): The 'frmDriveChoice' member is referenced only by members that haven't found to be used in the current application.

	Public isOK As Boolean

	Private Sub cancel_Click() Handles CANCEL.Click
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This just hides the Drivechoice form
		'Called By: Many
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		isOK = False
		Unload6(Me)
		
	End Sub

	Private Sub Drivelist_DblClick() Handles Drivelist.DblClick
		OK_Click()
	End Sub

	Private Sub OK_Click() Handles OK.Click
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: If the drive chosen is valid then write it to file, else prompt the user to choose another file
		'Called By: Many
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		Dim thedrive As String = Drivelist.List(Drivelist.ListIndex)

		If Isvaliddrive(thedrive) Then
			
			Call Drive_File(thedrive)
			
			'MsgBox "The Data Druid Drive has been set to " & thedrive & " Drive"
			isOK = True
			Me.Hide()
			
		Else
			
			MsgBox6("The Drive you have selected does not have the Corporate Data Set up on it, please choose another drive", MsgBoxStyle.Information, "Invalid Drive")
			
		End If
		
	End Sub

	' UPGRADE_INFO (#0701): The 'Direxists' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Public Function Direxists(ByRef pathname As String) As Object
		
		' EXCLUDED: '******************************************************************************************************************
		' EXCLUDED: 'Date: 28 / 03 / 04
		' EXCLUDED: 'Author: Nathan Eaton, CALM
		' EXCLUDED: 'Description
		' EXCLUDED: 'Called By: Many
		' EXCLUDED: 'Calls: None
		' EXCLUDED: 'Accepts: None
		' EXCLUDED: 'Returns: None
		' EXCLUDED: '******************************************************************************************************************
		' EXCLUDED: If Dir6(pathname, FileAttribute.Directory) <> "" Then
			
			' EXCLUDED: Return True
		' EXCLUDED: Else
			
			' EXCLUDED: Return False
		' EXCLUDED: End If
		
	' EXCLUDED: End Function

	' UPGRADE_INFO (#0561): The 'Isvaliddrive' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'thedrive' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Function Isvaliddrive(ByRef thedrive As String) As Object
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions evaluates whether the data druid drive specified is valid
		'Called By: Many
		'Calls: None
		'Accepts: The drive letter as a string, eg "E:"
		'Returns: True if the drive specified is valid, false if it is not
		'******************************************************************************************************************
		
		Dim subdirlist As New Collection

		Isvaliddrive = True
		
		'Insert strings that represent folders that should appear in the data druid directory structure
		
		subdirlist.Add(("Tenure"))
		subdirlist.Add(("Marine"))
		
		'Loop through the directory strings to see if they exist. If any of them do not, then the drive is not valid

		' UPGRADE_INFO (#0561): The 'idir' symbol was defined without an explicit "As" clause.
		Dim idir As Object = Nothing
		Dim Dirlocation As String = ""
		For Each idir In subdirlist
			Dirlocation = thedrive & ":" & GetDefaultMember6(modCommon_Functions.datalocation()) & idir
			
			If modCommon_Functions.Direxists(Dirlocation) = False Then
				Isvaliddrive = False
				'MsgBox Dirlocation
			End If
		Next
		
	End Function

	' UPGRADE_INFO (#0551): The 'thedrive' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub Drive_File(ByRef thedrive As String)
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This script writes the drive to the text file
		'Called By: OK_click
		'Calls: None
		'Accepts: The drive letter as a string, eg "E:"
		'Returns: Nothing
		'******************************************************************************************************************
		
		'Set Pathname to the text file
		Dim configpath As String = ""
		
		If GetDefaultMember6(OK.Tag) = "External" Then
			configpath = "c:\ExternalDataDruidPath.cfg" 'Change this to your path.
			If modCommon_Functions.Fileexists(configpath) Then
				modCommon_Functions.ChangeSingleFileAttributes(configpath, , Scripting.FileAttribute.[ReadOnly])
			End If
		Else
			configpath = "c:\DataDruidPath.cfg"
			If modCommon_Functions.Fileexists(configpath) Then
				modCommon_Functions.ChangeSingleFileAttributes(configpath, , Scripting.FileAttribute.[ReadOnly])
			End If
		End If
		FileOpen6(100, configpath, OpenMode.Output, OpenAccess.Default, OpenShare.Default, -1)
		
		'Print instead of Write (Write adds quote marks)
		FilePrintLine6(100, thedrive & ":")
		FileClose6(100)
		
	End Sub

End Class
