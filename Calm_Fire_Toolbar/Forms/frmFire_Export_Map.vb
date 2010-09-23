' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports VB = Microsoft.VisualBasic

Friend Class frmFire_Export_Map

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: Nathan Eaton, CALM
	'Description: This form specifies parameters to export images
	'Called By: clsFire_Export_Map
	'Calls: None
	'Accepts: None
	'Returns: None
	'******************************************************************************************************************
	
	Public Did_Cancel As Boolean
	
	Private Sub btnNewFile_Click() Handles btnNewFile.Click
		Dim dest_string As String = modCommon_Functions.NewJPG(modCommon_Functions.m_pApp.hWnd)
		
		If dest_string = "empty" Then
			Exit Sub
		End If
		
		Dim filename As String = modCommon_Functions.Extract_filename(dest_string)
		If VB.Right(filename, 4) <> ".jpg" Then
			dest_string = Replace(dest_string, filename, filename & ".jpg")
			filename &= ".jpg"
		End If
		
		Me.txtfilename.Text = filename
		Me.txtfolder.Text = Replace(dest_string, Me.txtfilename.Text, "")
	End Sub

	Private Sub Cancel_btn_Click() Handles Cancel_btn.Click
		Did_Cancel = True
		Me.Hide()
	End Sub

	Private Sub OK_btn_Click() Handles OK_btn.Click
		Did_Cancel = False
		Me.Hide()
	End Sub

    Private Sub frmFire_Export_Map_Load() Handles MyBase.Load

    End Sub
End Class
