' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmFire_Zoom

	Public RS As New ADODB.RecordsetClass
	
	Private Sub coord_option_Click() Handles coord_option.Click
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Configure the form to utilise user input coords
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		On Error GoTo ErrorHandler
101:
		Me.coordtype_cb.ListIndex = 0
102:
		Me.coordtype_cb.BackColor = SystemColors.Window 
103:
		Me.coordtype_cb.Enabled = True
104:
		Me.fd_tb.BackColor = SystemColors.InactiveBorder 
105:
		Me.feature_tb.BackColor = SystemColors.InactiveBorder 
106:
		Me.xcoord_tb.BackColor = SystemColors.Window 
107:
		Me.ycoord_tb.BackColor = SystemColors.Window 
108:
		Me.xcoord_tb.Enabled = True
109:
		Me.ycoord_tb.Enabled = True
110:
		Me.fd_tb.Enabled = False
111:
		Me.feature_tb.Enabled = False
112:
		Me.feature_tb.Text = ""
113:
		Me.fd_tb.Text = ""
114:
		Me.Width = 4455
115:
		coordtype_cb_Click()
		Exit Sub
ErrorHandler:
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private Sub coordtype_cb_Click() Handles coordtype_cb.Click
        If coordtype_cb.Text Like "Geographical (Long/Lat)" Then
            lblX.Caption = "Long"
            lblX.Alignment = 0
            lblY.Caption = "Lat"
            lblY.Alignment = 0
            xcoord_tb.Text = ""
            ycoord_tb.Text = "-"
        End If
		If coordtype_cb.Text Like "Grid (Eastings/Northings)" Then
			lblX.Caption = "E"
			lblX.Alignment = 2
			lblY.Caption = "N"
			lblY.Alignment = 2
			xcoord_tb.Text = ""
			ycoord_tb.Text = ""
		End If
	End Sub

	Private Sub fd_option_Click() Handles fd_option.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Configure the form to utilise predefinde FD grid coords
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		' UPGRADE_INFO (#0701): The 'f' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim f As ADODB.Recordset
		Dim CN As New ADODB.ConnectionClass
		Dim foldername As String = ""
		Dim txtDRIVER As String = ""
		Dim DataFile As String = ""
		Dim strSQL As String = ""
		Dim strProvider As String = ""
		If fd_option.Value = True Then
			Me.fd_tb.BackColor = SystemColors.Window 
			Me.fd_tb.Enabled = True
			Me.feature_tb.BackColor = SystemColors.InactiveBorder 
			Me.feature_tb.Text = ""
			Me.xcoord_tb.BackColor = SystemColors.InactiveBorder 
			Me.xcoord_tb.Text = ""
			Me.ycoord_tb.BackColor = SystemColors.InactiveBorder 
			Me.ycoord_tb.Text = ""
			Me.coordtype_cb.BackColor = SystemColors.InactiveBorder 
			Me.coordtype_cb.Enabled = False
			Me.MSHFlexGrid1.Visible = False
			Me.feature_tb.Enabled = False
			Me.xcoord_tb.Enabled = False
			Me.ycoord_tb.Enabled = False
			
			Me.Width = 8325

			If Environ("ARCGISHOME") = "" Then
				MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
				Exit Sub
			End If
			
			foldername = Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\"
			DataFile = "fd_coords.csv"
			
			'Open the dbf file containing the list of fd grid locations and coords
			strProvider = "Provider=MSDASQL;Extended Properties="""
			txtDRIVER = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ="
			CN.ConnectionString = strProvider & txtDRIVER & foldername & """"
			CN.Open()
			
			strSQL = "Select * From " & DataFile
			
			If RS.State = ADODB.ObjectStateEnum.adStateOpen Then
                ' RS = Nothing
                RS.Close()
			End If
			RS.Open(strSQL, CN, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
			
			Me.MSHFlexGrid1.Clear()
			Me.MSHFlexGrid1.Rows = 1
			
			Me.MSHFlexGrid1.ColWidth(0) = 1500
			
		End If
	End Sub

	Private Sub fd_tb_Change() Handles fd_tb.Change
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Implement a filter for the FD grid points based on user specified text
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		RS.Filter = ""
		
		'Load the datagrid with all the matching FD Grid records, limiting the display to 500 items to optimise speed of refresh
		If Len6(fd_tb.Text) > 0 Then
			RS.Filter = "FD_Grid Like '" & fd_tb.Text & "*'"
			If RS.RecordCount > 500 Then
				Me.MSHFlexGrid1.Visible = False
			ElseIf RS.RecordCount > 0 Then
				Me.MSHFlexGrid1.Visible = True
				RS.MoveFirst()
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			Else
				Me.MSHFlexGrid1.Visible = True
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			End If
		Else
			If RS.RecordCount > 500 Then
				Me.MSHFlexGrid1.Visible = False
			Else
				Me.MSHFlexGrid1.Visible = True
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			End If
		End If
	End Sub

	Private Sub feature_tb_Change() Handles feature_tb.Change
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: implement a filter for the features based on user input text
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		RS.Filter = ""
		
		If Len6(feature_tb.Text) > 0 Then
			RS.Filter = "NAME Like '" & feature_tb.Text & "*'"
			If RS.RecordCount > 500 Then
				Me.MSHFlexGrid1.Visible = False
			ElseIf RS.RecordCount > 0 And RS.RecordCount <= 500 Then
				Me.MSHFlexGrid1.Visible = True
				RS.MoveFirst()
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			Else
				Me.MSHFlexGrid1.Visible = True
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			End If
		Else
			If RS.RecordCount > 500 Then
				Me.MSHFlexGrid1.Visible = False
			Else
				Me.MSHFlexGrid1.Visible = True
				LoadRecordSetIntoGrid(Me.MSHFlexGrid1, RS)
			End If
		End If
		
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Load default form parameters
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		' UPGRADE_INFO (#0701): The 'f' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim f As ADODB.Recordset
		Dim CN As New ADODB.ConnectionClass
		
		Dim foldername As String = ""
		Dim txtDRIVER As String = ""
		Dim DataFile As String = ""
		Dim strSQL As String = ""
		Dim strProvider As String = ""
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
			Exit Sub
		End If
		
		foldername = Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\" '"C:\Temp\"
		DataFile = "locality_coords.csv"
		
		'Connect to and load the locality text file
		strProvider = "Provider=MSDASQL;Extended Properties="""
		txtDRIVER = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ="
		CN.ConnectionString = strProvider & txtDRIVER & foldername & """"
		CN.Open()
		
		strSQL = "Select * From " & DataFile
		RS.Open(strSQL, CN, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		Me.MSHFlexGrid1.Clear()
		Me.MSHFlexGrid1.ColWidth(0) = 1500
		Me.coordtype_cb.ListIndex = 0
		Me.MSHFlexGrid1.Visible = False
		Me.scale_cb.ListIndex = 0
		
	End Sub

	' UPGRADE_INFO (#0561): The 'Number_format' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'number_string' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Function Number_format(ByRef number_string As String) As Object
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: format a number
		'Called By
		'Calls: None
		'Accepts: Number as string
		'Returns: Number as double
		'******************************************************************************************************************
		On Error GoTo errhandler
		
		Dim mynumber As Double = number_string
		
		Return Format6(mynumber, "0.000000")
		
errhandler:
		Return False
		
	End Function

	' UPGRADE_INFO (#0551): The 'ctlGrid' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'RS' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'filterstring' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'isstart' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function LoadRecordSetIntoGrid(ByRef ctlGrid As CodeArchitects.VB6LibraryOcx.VB6MSHFlexGrid, ByRef RS As ADODB.Recordset, Optional ByRef filterstring As String = "", Optional ByRef isstart As Boolean = False) As Boolean
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Load the given recordset into the datagrid
		'Called By
		'Calls: None
		'Accepts: datagrid ref, recordset, optional filter string
		'Returns: None
		'******************************************************************************************************************
		On Error GoTo ErrorHandler
		
		Dim sTmp As String = ""
		Dim vArr As Object = Nothing
		Dim i As Short
		
		ctlGrid.Clear()
		ctlGrid.Rows = 1
        ctlGrid.Cols = 4

        ctlGrid.TextMatrix(0, 0) = "Name"
        ctlGrid.TextMatrix(0, 1) = "Feature"
        ctlGrid.TextMatrix(0, 2) = "Latitude"
        ctlGrid.TextMatrix(0, 3) = "Longtitude"

        '        Me.MSHFlexGrid1.TextMatrix(0, 0) = "Name"
        '		Me.MSHFlexGrid1.TextMatrix(0, 1) = "Feature"
        '	Me.MSHFlexGrid1.TextMatrix(0, 2) = "Latitude"
        '	Me.MSHFlexGrid1.TextMatrix(0, 3) = "Longitude"
		
		ctlGrid.ColHeaderCaption(0, 1) = "Header"
		
		If Not RS.EOF And Not RS.BOF Then
			'get the recordset into a string delimiting the fields
			'with a tab character and the records with a semicolon
			'replace nulls with a space
			sTmp = RS.GetString(ADODB.StringFormatEnum.adClipString, , Chr6(9), ";", " ")
			
			'split the string into an array of individual records
			vArr = Split6(sTmp, ";")
			
			'now add the records to the grid
			For i = 0 To UBound6(vArr) - 1
				ctlGrid.AddItem(vArr(i))
			Next
			'set the return value
			LoadRecordSetIntoGrid = True
		Else
			Return False
		End If
		
		If ctlGrid.Rows > 2 Then
			On Error Resume Next 'get over removing a single record
			For i = 1 To 1
				If ctlGrid.TextMatrix(i, 1) = "" Then
				Else
					Exit For
				End If
			Next
		End If
		
		If ctlGrid.Rows > 1 Then
			ctlGrid.FixedRows = 1
		End If
		
		Exit Function
		
ErrorHandler:
		MsgBox6("error " & Err.Description)
		Return False
	End Function

	Private Sub Form_Resize() Handles MyBase.Resize
		'GrantB 27-09-2007 Just tidying up bits
		If Me.Width > 5000 Then
			With MSHFlexGrid1
				.Top = 120
				.Left = 4680
				' UPGRADE_INFO (#0181): Reference to default form instance 'frmFire_Zoom' was converted to Me keyword.
				.Width = Me.Width - Line1.X1 - 510
				' UPGRADE_INFO (#0181): Reference to default form instance 'frmFire_Zoom' was converted to Me keyword.
				.Height = Me.Height - 675
			End With
		End If
	End Sub

	Private Sub Form_Unload(ByRef Cancel As Short) Handles MyBase.Unload
		RS.Close()
		RS = Nothing 'added 26/09/06
	End Sub

	Private Sub locality_option_Click() Handles locality_option.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Load the locality recordset and populate the datagrid
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		' UPGRADE_INFO (#0701): The 'f' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim f As ADODB.Recordset
		Dim CN As New ADODB.ConnectionClass
		Dim foldername As String = ""
		Dim txtDRIVER As String = ""
		Dim DataFile As String = ""
		Dim strSQL As String = ""
		Dim strProvider As String = ""
		If locality_option.Value = True Then
			
			Me.fd_tb.Text = ""
			Me.fd_tb.Enabled = False
			Me.xcoord_tb.Enabled = False
			Me.ycoord_tb.Enabled = False
			Me.fd_tb.BackColor = SystemColors.InactiveBorder 
			Me.coordtype_cb.BackColor = SystemColors.InactiveBorder 
			Me.coordtype_cb.Enabled = False
			Me.feature_tb.Enabled = True
			Me.feature_tb.BackColor = SystemColors.Window 
			Me.xcoord_tb.BackColor = SystemColors.InactiveBorder 
			Me.ycoord_tb.BackColor = SystemColors.InactiveBorder 
			Me.Width = 8325
			Me.MSHFlexGrid1.Visible = False

			'Open and load the locality text file
			
			If Environ("ARCGISHOME") = "" Then
				MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
				Exit Sub
			End If
			
			foldername = Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\" '"C:\Temp\"
			DataFile = "locality_coords.csv"
			
			strProvider = "Provider=MSDASQL;Extended Properties="""
			txtDRIVER = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ="
			CN.ConnectionString = strProvider & txtDRIVER & foldername & """"
			CN.Open()
			
			strSQL = "Select * From " & DataFile
			
			If RS.State = ADODB.ObjectStateEnum.adStateOpen Then
                RS.Close()
                'RS = Nothing
			End If
			
			RS.Open(strSQL, CN, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
			
			Me.MSHFlexGrid1.Clear()
			Me.MSHFlexGrid1.Rows = 1
			Me.MSHFlexGrid1.ColWidth(0) = 1500
			
		End If
		
	End Sub

	Private Sub MSHFlexGrid1_DblClick() Handles MSHFlexGrid1.DblClick
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Call the zoom btn click when double clicking the datagrid
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		Zoom_btn_Click()
	End Sub

	Private Sub Zoom_btn_Click() Handles Zoom_btn.Click
		'******************************************************************************************************************
		'Date: 06 / 06 / 06
		'Author: Nathan Eaton, CALM
		'Description: Perform the zoom function
		'Called By
		'Calls: None
		'Accepts: None
		'Returns: None
		'******************************************************************************************************************
		
		On Error GoTo errhandler
		
		' UPGRADE_INFO (#0561): The 'x1' symbol was defined without an explicit "As" clause.
		Dim x1 As Object = Nothing
		' UPGRADE_INFO (#0561): The 'y1' symbol was defined without an explicit "As" clause.
		Dim y1 As Object = Nothing
		' UPGRADE_INFO (#0561): The 'X' symbol was defined without an explicit "As" clause.
		Dim X As Object = Nothing
		Dim y As Double
		Dim isproj_coord As Boolean = False
		
		'Reference the coordinate values from the form
		Dim myzone As String = ""
		' UPGRADE_INFO (#0701): The 'xcoord' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim xcoord As Object = Nothing
		' UPGRADE_INFO (#0701): The 'ycoord' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim ycoord As Object = Nothing
		If Me.coord_option.Value = True Then
			If Me.coordtype_cb.Text = "Grid (Eastings/Northings)" Then
				'******************
				isproj_coord = True
				frmCombo_Dialog.Caption = "Select MGA Zone"
				frmCombo_Dialog.cb_instruct.Caption = "Please select the MGA zone for your coordinates"
				frmCombo_Dialog.Values_cb.AddItem("49")
				frmCombo_Dialog.Values_cb.AddItem("50")
				frmCombo_Dialog.Values_cb.AddItem("51")
				frmCombo_Dialog.Values_cb.AddItem("52")
				frmCombo_Dialog.Values_cb.ListIndex = 1
				frmCombo_Dialog.Show(VBRUN.FormShowConstants.vbModal)
				'GrantB 15-10-2007 - If this form is canceled the process fals down - need to check for cancel
				If frmCombo_Dialog.frmCanceled Then
					Exit Sub
				Else
					myzone = frmCombo_Dialog.listvalue
				End If
				'*************************
			End If
			
			x1 = Number_format(ByVal6(Me.xcoord_tb.Text))
			y1 = Number_format(ByVal6(Me.ycoord_tb.Text))
			If x1 = False Or y1 = False Then
				MsgBox6("The coordinates need to be in decimal format", MsgBoxStyle.Exclamation, "Coordinate Format")
				Exit Sub
			End If
		Else
			y1 = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
			x1 = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 3)
		End If
		
		X = x1
		y = y1
		
		Dim pPoint As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point()
		
		pPoint.X = X
		pPoint.Y = y
		
		Dim projlocation As String = ""
		Dim geolocation As String = ""
		Dim spatialrefstring As String = ""
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Your ARCGISHOME Environment Variable has not been set. Contact the GIS Section on 9334 0158 for help", MsgBoxStyle.Exclamation, "Error in Setup")
			Unload6(Me)
			Exit Sub
		End If
		
		'Create the projection file string
		projlocation = Environ("ARCGISHOME") & "\Coordinate Systems\Projected Coordinate Systems\National Grids\Australia\"
		geolocation = Environ("ARCGISHOME") & "\Coordinate Systems\Geographic Coordinate Systems\Australia and New Zealand\"
		
		If isproj_coord = True Then
			spatialrefstring = projlocation & "GDA 1994 MGA Zone " & myzone & ".prj"
		Else
			spatialrefstring = geolocation & "Geocentric Datum of Australia 1994.prj"
		End If
		
		Dim pSpatReffact As ESRI.ArcGIS.Geometry.ISpatialReferenceFactory2 = New ESRI.ArcGIS.Geometry.SpatialReferenceEnvironment()
		
		'Create the Spatial reference from the Projection File
		Dim pspatref As ESRI.ArcGIS.Geometry.ISpatialReference = pSpatReffact.CreateESRISpatialReferenceFromPRJFile(spatialrefstring)
		
		pPoint.SpatialReference = pspatref
		Dim pEnv As ESRI.ArcGIS.Geometry.IEnvelope = New ESRI.ArcGIS.Geometry.Envelope()
		
		Dim mydoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		Dim myview As ESRI.ArcGIS.Carto.IActiveView = mydoc.FocusMap

		pEnv = myview.Extent
		
		Dim origspatref As ESRI.ArcGIS.Geometry.ISpatialReference = mydoc.FocusMap.SpatialReference
		pEnv.SpatialReference = origspatref
		
		'Centre view on X, Y coord
		pEnv.Project(pPoint.SpatialReference)
		pEnv.CenterAt(pPoint)
		pEnv.Project(origspatref)
		myview.Extent = pEnv
		
		Dim myscaletext As String = ""
		Dim myscale As Double
		
		'Set the scale
		If Me.scale_cb.Text <> "Current Scale" Then
			myscaletext = Me.scale_cb.Text
			If myscaletext Like "1:*" Then
				myscaletext = Replace(myscaletext, "1:", "")
			End If
			myscale = myscaletext
			mydoc.FocusMap.MapScale = myscale
		End If
		
		myview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography, Nothing, Nothing) 'call again to refresh selection
		
		Exit Sub
errhandler:
		MsgBox6("Unable to perform Zoom Function - " & Err.Description, MsgBoxStyle.Critical, "Error")
	End Sub

End Class
