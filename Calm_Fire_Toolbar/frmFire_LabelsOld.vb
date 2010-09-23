' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default
Imports Microsoft.VisualBasic.Compatibility
Friend Class frmFire_Labels

	'variables for form resizing
	Friend Structure ControlPositionType
	
		Public Left As Single
		Public Top As Single
		Public Width As Single
		Public Height As Single
		Public FontSize As Single
	End Structure

	Private m_ControlPositions() As frmFire_Labels.ControlPositionType
	Private m_FormWid As Single
	Private m_FormHgt As Single
	
	'Public pST As ICommandItem
	'Public pAP As IApplication
	
	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This form specifies text to be used to create Fire Labels with
	'Called By: clsFire_labels
	'Calls: None
	'Accepts: None
	'Returns: None
	'******************************************************************************************************************
	
   

    Private Sub btnEdit_Click() Handles btnEdit.Click

        'Dim origtext As String = Me.lblList.List(Me.lblList.ListIndex)
        Dim origtext As String = lblList.SelectedItem

        If origtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox6("Cannot edit this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If

        Dim edittext As String = ""
        If Me.lblList.SelectedItems.Count > 0 Then
            edittext = InputBox("Enter New Text", "Edit Text", VB6.GetItemString(Me.lblList, Me.lblList.SelectedIndex))
            '            edittext = InputBox6("Enter New Text", "Edit Text", Me.lblList.SelectedItem.ToString())
            If edittext = "" Then 'added 12/10/06 to account for cancel button
                Exit Sub
            Else
                VB6.SetItemString(Me.lblList, Me.lblList.SelectedIndex, edittext)

            End If
        End If

    End Sub

    Private Sub cb_CustomText_Click() Handles cb_CustomText.Click

        If cb_CustomText.Value = 1 Then
            Me.lblList.Enabled = False
            Me.btnAdd.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnEdit.Enabled = False
            Me.lblList.BackColor = SystemColors.ScrollBar
            Me.tbCustomText.BackColor = SystemColors.Window
        Else
            Me.btnAdd.Enabled = True
            Me.btnDelete.Enabled = True
            Me.btnEdit.Enabled = True
            Me.lblList.Enabled = True
            Me.lblList.BackColor = SystemColors.Window
            Me.tbCustomText.BackColor = SystemColors.ScrollBar
        End If

    End Sub

    Private Sub btnDelete_Click() Handles btnDelete.Click

        Dim origtext As String
        origtext = VB6.GetItemString(Me.lblList, Me.lblList.SelectedIndex)
        If origtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox6("Cannot delete this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If
        If Me.lblList.SelectedItems.Count > 0 Then
            Me.lblList.Items.RemoveAt((Me.lblList.SelectedIndex))
        End If

    End Sub

	Private Sub btnClose_Click() Handles btnClose.Click
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler
		
		Dim pSelectTool As ESRI.ArcGIS.Framework.ICommandItem = Nothing
		Dim pCommandBars As ESRI.ArcGIS.Framework.ICommandBars
		' The identifier for the Select Graphics Tool
		Dim pUID As New ESRI.ArcGIS.esriSystem.UIDClass
		pUID.Value = "esriArcMapUI.SelectTool"
		
		'Find the Select Graphics Tool
		pCommandBars = m_pApp.Document.CommandBars
		'GrantB 19-10-2007 - need to know if the tool was set
		pSelectTool = pCommandBars.Find(pUID)
		If pSelectTool Is Nothing Then
			MsgBox6("The Tool could not be set! Unable to perform this action.")
		End If
		m_pApp.CurrentTool = pSelectTool
		Me.Hide()
		modFire_Tools.Incident_StopEditing()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' Save the form's and controls' dimensions.
	Private Sub SaveSizes()
		'added 27/09/06
        Dim i As Short = 0
		Dim ctl As Object = Nothing
		
		' Save the controls' positions and sizes.
        ReDim m_ControlPositions(Controls6.Count)
		For Each ctl In Controls6
			With m_ControlPositions(i)
                If TypeOf ctl Is VB6Line Then
                    .Left = GetDefaultMember6(CObj(ctl).x1)
                    .Top = GetDefaultMember6(CObj(ctl).y1)
                    .Width = GetDefaultMember6(CObj(ctl).X2) - GetDefaultMember6(CObj(ctl).x1)
                    .Height = GetDefaultMember6(CObj(ctl).Y2) - GetDefaultMember6(CObj(ctl).y1)
                Else
                    .Left = GetDefaultMember6(ctl.Left)
                    .Top = GetDefaultMember6(ctl.Top)
                    .Width = GetDefaultMember6(ctl.Width)
                    .Height = GetDefaultMember6(ctl.Height)
                    On Error Resume Next
                    .FontSize = GetDefaultMember6(ctl.Font.Size)
                    On Error GoTo 0
                End If
			End With
			i += 1
		Next
		
		' Save the form's size.
		m_FormWid = ScaleWidth
		m_FormHgt = ScaleHeight
	End Sub

	Private Sub cmdAddLabel_Click() Handles cmdAddLabel.Click
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		'added 10/10/06 to make label editing easier
		On Error GoTo ErrorHandler
		
		Incident_StartEditing()
		
		Dim pSelectTool As ESRI.ArcGIS.Framework.ICommandItem
		Dim pCommandBars As ESRI.ArcGIS.Framework.ICommandBars = Nothing
		' The identifier for the Select Graphics Tool
		Dim pUID As New ESRI.ArcGIS.esriSystem.UIDClass
		pUID.Value = "CALM_FireTools.clsFire_Labels_Tool"
		
		'Find the Select Graphics Tool
		'GrantB 19-10-2007 - need to know if the toolBar was set
		pCommandBars = m_pApp.Document.CommandBars
		If pCommandBars Is Nothing Then
			MsgBox6("The Bar could not be set! Unable to perform this action.")
			Exit Sub
		End If
		
		'GrantB 19-10-2007 - need to know if the tool was set
		pSelectTool = Nothing
		pSelectTool = pCommandBars.Find(pUID)
		If pSelectTool Is Nothing Then
			MsgBox6("The Tool could not be set! Unable to perform this action.")
			Exit Sub
		End If
		
		m_pApp.CurrentTool = pSelectTool
		'    Set pAP.CurrentTool = pST
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick - " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub cmdMoveLabel_Click() Handles cmdMoveLabel.Click
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		'added 10/10/06 to make label editing easier
		On Error GoTo ErrorHandler
		
		modFire_Tools.makeOnlySelectableLayer(modFire_Tools.Get_Fire_Anno_Layer())
		Incident_StartEditing()
		
		Dim pSelectTool As ESRI.ArcGIS.Framework.ICommandItem = Nothing
		Dim pCommandBars As ESRI.ArcGIS.Framework.ICommandBars
		' The identifier for the Select Graphics Tool
		Dim pUID As New ESRI.ArcGIS.esriSystem.UIDClass
		pUID.Value = "esriEditor.EditTool"
		
		'Find the Select Graphics Tool
		pCommandBars = m_pApp.Document.CommandBars
		'GrantB 19-10-2007 - need to know if the tool was set
		pSelectTool = pCommandBars.Find(pUID)
		If pSelectTool Is Nothing Then
			MsgBox6("The Tool could not be set! Unable to perform this action")
		End If
		m_pApp.CurrentTool = pSelectTool
		
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
		
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
        SaveSizes()
	End Sub

	Private Sub Form_Resize() Handles MyBase.Resize
        ResizeControls()
	End Sub

	' Arrange the controls for the new size.
	Private Sub ResizeControls()
		'added 27/09/06
		Dim i As Short
		Dim ctl As Object = Nothing
		Dim xChange As Single
		Dim yChange As Single
		
		' Don't bother if we are minimized.
		If WindowState = FormWindowState.Minimized Then Exit Sub
		
		'Get the change in the form size
		xChange = ScaleWidth - m_FormWid
		yChange = ScaleHeight - m_FormHgt
		
		' Position the controls.
        i = 0
		For Each ctl In Controls6
			With m_ControlPositions(i)
				If TypeOf ctl Is VB6Line Then
					CObj(ctl).x1 = xChange + .Left
					CObj(ctl).y1 = yChange + .Top
					CObj(ctl).X2 = GetDefaultMember6(CObj(ctl).x1) + .Width
					CObj(ctl).Y2 = GetDefaultMember6(CObj(ctl).y1) + .Height
                ElseIf TypeOf ctl Is ListBox Then
                    ctl.Width = xChange + .Width
                    ctl.Height = yChange + .Height
				Else
					ctl.Left = xChange + .Left
					ctl.Top = yChange + .Top
					On Error Resume Next
				End If
			End With
			i += 1
		Next
	End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ' UPGRADE_INFO (#05B1): The 'origtext' variable wasn't declared explicitly.
        Dim origtext As Object = Nothing
        Dim newtext As String = InputBox6("Enter New Text", "Add Text to List", "Enter Text Here")

        If newtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox6("Cannot add this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If

        If newtext = "" Then 'added 12/10/06 to account for cancel button
            Exit Sub
        Else
            Me.lblList.Items.Add((newtext))
        End If
    End Sub
End Class
