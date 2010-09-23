' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmRegion_Choice
' UPGRADE_INFO (#0511): The 'frmRegion_Choice' member is referenced only by members that haven't found to be used in the current application.

	Private Sub cancel_Click() Handles Cancel.Click
		
		Me.Hide()
		
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
		
		Dim theregion As String = ""
		
		If RegionList.List(RegionList.ListIndex) <> "" Then
			theregion = RegionList.List(RegionList.ListIndex)
			Call Region_write(theregion)
			'MsgBox "The Region has been set to " & theregion
			Me.Hide()
		Else
			MsgBox6("Either select a region or press cancel", MsgBoxStyle.Exclamation, "Select Region")
		End If
		
	End Sub

	Public Sub Region_write(ByRef theregion As String)
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This script writes the drive to the text file
		'Called By: Many
		'Calls: None
		'Accepts: The drive letter as a string, eg "E:"
		'Returns: Nothing
		'******************************************************************************************************************
		
		'Set Pathname to the text file
		
		If theregion = "South West" Then
			theregion = "South_West"
		ElseIf theregion = "South Coast" Then
			theregion = "South_Coast"
		End If
		
		Dim configpath As String = "c:\DataDruidRegion.cfg"
		'Change this to your path.
		FileOpen6(100, configpath, OpenMode.Output, OpenAccess.Default, OpenShare.Default, -1)
		
		'Print instead of Write (Write adds quote marks)
		FilePrintLine6(100, theregion)
		FileClose6(100)
		
	End Sub

	Private Sub RegionList_DblClick() Handles RegionList.DblClick
		
		OK_Click()
		
	End Sub

End Class
