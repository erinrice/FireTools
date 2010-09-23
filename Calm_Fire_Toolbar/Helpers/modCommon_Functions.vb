' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports VB = Microsoft.VisualBasic

Friend Module modCommon_Functions

    Public m_pApp As ESRI.ArcGIS.Framework.IApplication

	Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
    Private Const SW_SHOW As Short = 5
	
    Public Function Fileexists(ByRef pathname As String) As Object

        '******************************************************************************************************************
        'Date: 28 / 03 / 04
        'Author: Nathan Eaton, CALM
        'Description: This functions accepts a filename and evaluates whether or not the file exists
        'Called By: Many
        'Calls: None
        'Accepts: Full pathname of file, eg "C:\Temp\data.shp"
        'Returns: True if the file exists, false if the file does not exist
        '******************************************************************************************************************

        If Dir6(pathname) = "" Then
            Return False
        Else
            Return True
        End If

    End Function
	
	' UPGRADE_INFO (#0561): The 'Direxists' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'pathname' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function Direxists(ByRef pathname As String) As Object
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions accepts a directory name and evaluates whether or not the directory exists
		'Called By: Many
		'Calls: None
		'Accepts: Pathname of directory
		'Returns: True if the directory exists, false if the directory does not exist
		'******************************************************************************************************************
		If Dir6(pathname, FileAttribute.Directory) <> "" Then
			
			Return True
		Else
			
			Return False
		End If
		
	End Function
	

	' UPGRADE_INFO (#0561): The 'datalocation' symbol was defined without an explicit "As" clause.
	Public Function datalocation() As Object
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions contains the parameter that represents where the data is stored.
		'Called By: Many
		'Calls: None
		'Accepts: Nothing
		'Returns: A string the representing the base directory structure of Data Druid
		'******************************************************************************************************************
		
		'Dim Datalocation As String
		
		Return "\GIS1-Corporate\data\"
		
	End Function
	

    ' UPGRADE_INFO (#0551): The 'pdlayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    Public Function getPath(ByVal pDataLayer As ESRI.ArcGIS.Carto.IDataLayer2) As String
        '******************************************************************************************************************
        'Description: gets the full path of the specified layer
        'Accepts: a layer
        'Returns: full path to that layer
        '******************************************************************************************************************

        Dim pDSName As ESRI.ArcGIS.Geodatabase.IDatasetName
        Dim pFCName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
        Dim layerPath As String

        pDSName = pDataLayer.DataSourceName

        'Get the path of the workspace which holds the dataset
        layerPath = pDSName.WorkspaceName.PathName

        If TypeOf pDSName Is ESRI.ArcGIS.Geodatabase.IFeatureClassName Then
            pFCName = pDSName
            If Not pFCName.FeatureDatasetName Is Nothing Then
                layerPath = layerPath & "\" & pFCName.FeatureDatasetName.Name
            Else
                layerPath = layerPath & "\" & pDSName.Name
            End If
        End If


        'If TypeOf (pDSName.WorkspaceName.WorkspaceFactory) Is ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory Then
        'layerPath = layerPath & ".shp"
        'End If

        If pDSName.Category = "Shapefile Feature Class" Then
            layerPath = layerPath & ".shp"
        End If

        'If the Datalayer represents a shapefile, then .shp needs to be added onto the end of the string
        'If pDSName.WorkspaceName.WorkspaceFactory Is New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory Then
        '    layerPath = layerPath & ".shp"
        'End If

        Return layerPath
    End Function
	
    
	
	' UPGRADE_INFO (#0551): The 'hWnd' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ExecAss(ByRef hWnd As Integer, ByRef sCmdLine As String)
		'Executes scmdline via file association ie c:\bootlog.txt will run notepad!
		Dim Ret As Integer = ShellExecute(hWnd, "open", sCmdLine, Nothing, CurDir6(), SW_SHOW)
		If Ret < 32 Then
			'Error can't shell
			MsgBox6("Failed to execute: " & sCmdLine)
		End If
	End Sub
	

	' UPGRADE_INFO (#0561): The 'currenthwd' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'currenthwd' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function OpenGeodatabase(Optional ByRef currenthwd As Object = MissingValue6) As String
		' UPGRADE_INFO (#03C1): Next statement accounts for empty, null, and missing optional parameters.
		FixOptionalParameter6(currenthwd)
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions opens a dialog that allows the user to select a shapefile to open/use
		'Called By: Many
		'Calls: None
		'Accepts: Nothing
		'Returns: The shapefile as an IGxObject
		'******************************************************************************************************************
		
		Dim bObjectSelected As Boolean
		Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
		Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterPersonalGeodatabases()
		
		With pGxDialog
			.AllowMultiSelect = False
			.ButtonCaption = "Open"
			.Title = "Open CALM Fire Incident Geodatabase"
			.ObjectFilter = pFilter
			bObjectSelected = .DoModalOpen(currenthwd, pEnumGxObject) ' modCommon_Functions.m_pApp.hWnd
		End With
		
		If bObjectSelected = False Then
			Return "empty"
		End If
		
		Dim pgxobject As ESRI.ArcGIS.Catalog.IGxObject = pEnumGxObject.Next()
		
		Return pgxobject.FullName
		
	End Function
	
	' UPGRADE_INFO (#0561): The 'currenthwd' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'currenthwd' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function NewGeodatabase(Optional ByRef currenthwd As Object = MissingValue6) As String
		' UPGRADE_INFO (#03C1): Next statement accounts for empty, null, and missing optional parameters.
		FixOptionalParameter6(currenthwd)
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions opens a dialog that allows the user to select a shapefile to open/use
		'Called By: Many
		'Calls: None
		'Accepts: Nothing
		'Returns: The shapefile as an IGxObject
		'******************************************************************************************************************
		
		Dim bObjectSelected As Boolean
		Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
		' UPGRADE_INFO (#0701): The 'pEnumGxObject' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterPersonalGeodatabases()
		
		With pGxDialog
			'.AllowMultiSelect = False
			.ButtonCaption = "New"
			.Title = "Create CALM Fire Incident Geodatabase"
			.ObjectFilter = pFilter
			bObjectSelected = .DoModalSave(currenthwd) ' modCommon_Functions.m_pApp.hWnd
		End With
		
		If bObjectSelected = False Then
			Return "empty"
		End If
		
		Return pGxDialog.FinalLocation.FullName & "\" & Replace(pGxDialog.Name, ".mdb", "") & ".mdb"
		
	End Function
	
	' UPGRADE_INFO (#0561): The 'currenthwd' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'currenthwd' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function NewJPG(Optional ByRef currenthwd As Object = MissingValue6) As String
		' UPGRADE_INFO (#03C1): Next statement accounts for empty, null, and missing optional parameters.
		FixOptionalParameter6(currenthwd)
		
		'******************************************************************************************************************
		'Date: 28 / 03 / 04
		'Author: Nathan Eaton, CALM
		'Description: This functions opens a dialog that allows the user to select a shapefile to open/use
		'Called By: Many
		'Calls: None
		'Accepts: Nothing
		'Returns: The shapefile as an IGxObject
		'******************************************************************************************************************
		
		Dim bObjectSelected As Boolean
		Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
		' UPGRADE_INFO (#0701): The 'pEnumGxObject' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = modCommon_Functions.m_pApp.Document
		
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets()
		' GxFilterPersonalGeodatabases
		
		With pGxDialog
			'.AllowMultiSelect = False
			.ButtonCaption = "New"
			.Title = "Export Map to JPG Image"
			
			.ObjectFilter = pFilter
			
			bObjectSelected = .DoModalSave(currenthwd) ' modCommon_Functions.m_pApp.hWnd
		End With
		
		If bObjectSelected = False Then
			Return "empty"
		End If
		
		Return pGxDialog.FinalLocation.FullName & "\" & pGxDialog.Name
		
	End Function


    ' UPGRADE_INFO (#0561): The 'Extract_filename' symbol was defined without an explicit "As" clause.
    ' UPGRADE_INFO (#0551): The 'TokenList' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function Extract_filename(ByRef TokenList As String) As Object
		' extracts an element from a list of elements
		
		Dim strArray() As String
		Dim thecount As Short
		strArray = Split6(TokenList, "\")
		thecount = UBound6(strArray)
		Return strArray(thecount)

	End Function

    
	
    
    ' UPGRADE_INFO (#0561): The 'ChangeSingleFileAttributes' symbol was defined without an explicit "As" clause.
    ' UPGRADE_INFO (#0551): The 'strpath' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    ' UPGRADE_INFO (#0551): The 'lngSetAttr' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    ' UPGRADE_INFO (#0551): The 'lngRemoveAttr' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function ChangeSingleFileAttributes(ByRef strpath As String, Optional ByRef lngSetAttr As Scripting.FileAttribute = 0, Optional ByRef lngRemoveAttr As Scripting.FileAttribute = 0) As Object
		Dim fsoSysObj As VB6FileSystemObject = New VB6FileSystemObject()
		' UPGRADE_INFO (#0701): The 'fdrFolder' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim fdrFolder As VB6Folder
		' UPGRADE_INFO (#0701): The 'fdrSubFolder' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim fdrSubFolder As VB6Folder
		Dim filFile As VB6File = fsoSysObj.GetFile(strpath)

		If lngSetAttr Then
			
			If Not (filFile.Attributes And lngSetAttr) Then
				filFile.Attributes = filFile.Attributes Or lngSetAttr
			End If
			
		End If
		
		' If caller passed in attribute to remove, remove for all.
		If lngRemoveAttr Then
			
			If (filFile.Attributes And lngRemoveAttr) Then
				filFile.Attributes -= lngRemoveAttr
			End If
			
		End If
		
		Return True
		
	End Function
	
	
    
    ' UPGRADE_INFO (#0561): The 'expand_layer' symbol was defined without an explicit "As" clause.
    ' UPGRADE_INFO (#0551): The 'mylayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    ' UPGRADE_INFO (#0551): The 'expand' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function expand_layer(ByRef mylayer As ESRI.ArcGIS.Carto.ILayer, ByRef expand As Boolean) As Object
		
		Dim pLegendGroup As ESRI.ArcGIS.Carto.ILegendGroup = New ESRI.ArcGIS.Carto.LegendGroup()
		
		Dim pLegendInfo As ESRI.ArcGIS.Carto.ILegendInfo = mylayer
		
		pLegendGroup = pLegendInfo.LegendGroup(0)
		pLegendGroup.Visible = expand 'True for Expand, False for Collapse
	End Function

	' UPGRADE_INFO (#0561): The 'Apply_Layer_File_Symbology' symbol was defined without an explicit "As" clause.
	' UPGRADE_INFO (#0551): The 'mylayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'LayerFile' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function Apply_Layer_File_Symbology(ByRef mylayer As ESRI.ArcGIS.Carto.ILayer, ByRef LayerFile As String) As Object
		
		If Fileexists(LayerFile) = False Then
			Exit Function
		End If
		
		Dim pgxfile As ESRI.ArcGIS.Catalog.IGxFile
		Dim pGXLayer As ESRI.ArcGIS.Catalog.IGxLayer = New ESRI.ArcGIS.Catalog.GxLayer()
		
		pgxfile = pGXLayer
		
		' UPGRADE_INFO (#0701): The 'alternate_lyr_fullpath' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim alternate_lyr_fullpath As String = ""
		' UPGRADE_INFO (#0701): The 'end_layer_string' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim end_layer_string As String = ""
		
		pgxfile.Path = LayerFile
		
		Dim pSource As ESRI.ArcGIS.Carto.IGeoFeatureLayer = pGXLayer.Layer
		Dim pTarget As ESRI.ArcGIS.Carto.IGeoFeatureLayer = mylayer
		
		'??? 'get a reference the target layer however you want.

        If modCommon_Functions.Fileexists(ByVal6(pgxfile.Path)) Then
            pTarget.Renderer = pSource.Renderer
        End If
	End Function


	' UPGRADE_INFO (#0551): The 'App' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'ExtensionName' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function FindMXExtByName(ByRef App As ESRI.ArcGIS.Framework.IApplication, ByRef ExtensionName As String) As ESRI.ArcGIS.esriSystem.IExtension
		' UPGRADE_INFO (#05B1): The 'c_ModuleFileName' variable wasn't declared explicitly.
		Dim c_ModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler
		
		Dim pExMgr As ESRI.ArcGIS.esriSystem.IExtensionManager
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
12:
		pExtension = Nothing
		
14:
		pExMgr = App
15:
		pExtension = pExMgr.FindExtension(ExtensionName)
16:
		Return pExtension

ErrorHandler:
		HandleError(False, "FindMXExtByName " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
22:
		Return Nothing
	End Function

	' UPGRADE_INFO (#0551): The 'pWkSpace' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'sProperty' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function GetFireInfoProp(ByRef pWkSpace As ESRI.ArcGIS.Geodatabase.IWorkspace, ByRef sProperty As String) As String
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler

		'PURPOSE: Get the Value for an Incident Information Property
		'RETURN: a string of the value or 'ERROR' with msg on second line
		
		Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
18:
		pCursor = QueryInfoTable(pWkSpace, sProperty, False)
		
		'Get the Row and check if it OK
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
22:
		pRow = pCursor.NextRow()
		
24:
		If pRow Is Nothing Then
			'
			'Return Error message for no row
			'
28:
			Err.Raise(10000, "GetFireInfoProp", "NO row returned in query. WhereClaus:" & g_FIELD_PROPNAME & "='" & sProperty & "'")
			
30:
		End If
		
		'Find the value field index
		Dim lFld As Integer
34:
		lFld = pRow.Fields.FindField(g_FIELD_PROPVALUE)
		
36:
		If lFld = -1 Then
			'
			'Return Error message for no row
			'
40:
			Err.Raise(10000, "GetCSInfoProp", "Can't find value field:" & g_FIELD_PROPVALUE)
			
42:
		End If
		
		'Get the Value
45:
		If IsNull6(GetDefaultMember6(pRow.Value(lFld))) Then
			
47:
			GetFireInfoProp = ""
			
49:
		Else
			
51:
			GetFireInfoProp = GetDefaultMember6(pRow.Value(lFld))
			
53:
		End If
		
		'Clean up
56:
		pCursor = Nothing
57:
		pRow = Nothing

		Exit Function
ErrorHandler:
		HandleError(False, "GetFireInfoProp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	' UPGRADE_INFO (#0551): The 'pWkSpace' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'sProperty' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'sValue' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function SetFireInfoProp(ByRef pWkSpace As ESRI.ArcGIS.Geodatabase.IWorkspace, ByRef sProperty As String, ByRef sValue As String) As Boolean
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing

        Try

            '           On Error GoTo ErrorHandler

            'PURPOSE: Set the Value for an Incident Information Property
            'RETURN: a boolean if success

            Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
74:
            pCursor = QueryInfoTable(pWkSpace, sProperty, True)

            'Get the Row and check if it OK
            Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
78:
            pRow = pCursor.NextRow()

80:
            If pRow Is Nothing Then
                '
                'Return Error message for no row
                '
84:
                Err.Raise(10000, "SetFireInfoProp", "NO row returned in query. WhereClaus:" & g_FIELD_PROPNAME & "='" & sProperty & "'")

86:
            End If

            'Find the value field index
            Dim lFld As Integer
90:
            lFld = pRow.Fields.FindField(g_FIELD_PROPVALUE)

92:
            If lFld = -1 Then
                '
                'Return Error message for no row
                '
96:
                Err.Raise(10000, "SetCSInfoProp", "Can't find value field:" & g_FIELD_PROPVALUE)

98:
            End If

100:
            pRow.Value(lFld) = sValue
101:
            pRow.Store()

103:
            Return True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'ErrorHandler:
        '        HandleError(False, "SetFireInfoProp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function
	
	' UPGRADE_INFO (#0551): The 'pRow' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'pField' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'pValue' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function SetTableValue(ByRef pRow As ESRI.ArcGIS.Geodatabase.IRow, ByRef pField As String, ByRef pValue As String) As Boolean
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler

		'Find the value field index
		Dim lFld As Integer
90:
		lFld = pRow.Fields.FindField(pField)
		
92:
		If lFld = -1 Then
			'
			'Return Error message for no row
			'
96:
			Err.Raise(10000, "SetCSInfoProp", "Can't find value field:" & pField)
			
98:
		End If
		
100:
		pRow.Value(lFld) = pValue
101:
		pRow.Store()
		
103:
		Return True

ErrorHandler:
		HandleError(False, "SetFireInfoProp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	' UPGRADE_INFO (#0551): The 'pFWKS' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'sProperty' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'bWrite' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Function QueryInfoTable(ByRef pFWKS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace, ByRef sProperty As String, ByRef bWrite As Boolean) As ESRI.ArcGIS.Geodatabase.ICursor
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler

		'Purpose: Query the Incident Information Table for a give property
		'Return: a row
		
		'
		'Locate the incident info table from the workspace
		'
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
123:
		pTable = pFWKS.OpenTable(g_TABLE_INCIDENT_INFO)
124:
		If pTable Is Nothing Then
			
			'
			'Raise Error message for no table
			'
129:
			Err.Raise(10000, "QueryInfoTable", "Unable to find Incident Information table (" & g_TABLE_INCIDENT_INFO & ") in workspace.")
			
131:
		End If
		
		'
		'Create a Query Filter
		'
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter
137:
		pQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter()
		
		' create the where statement
		Dim sWhere As String = ""
141:
		sWhere = g_FIELD_PROPNAME & "='" & sProperty & "'"
142:
		pQueryFilter.WhereClause = sWhere
143:
		pQueryFilter.SubFields = g_FIELD_PROPVALUE
		
		' Query the Table for the property value row
		Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
147:
		If bWrite Then
			
149:
			pCursor = pTable.Update(pQueryFilter, False)
			
151:
		Else
			
153:
			pCursor = pTable.Search(pQueryFilter, True)
			
155:
		End If
156:
		pTable = Nothing
157:
		pQueryFilter = Nothing
		
159:
		If pCursor Is Nothing Then
			
			'
			'Return Error message for no cursor
			'
164:
			Err.Raise(10000, "QueryInfoTable", "Unable to create Cursor. WhereClaus:" & sWhere)
			
166:
		End If
		
168:
		Return pCursor

ErrorHandler:
		HandleError(False, "QueryInfoTable " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	' UPGRADE_INFO (#0551): The 'App' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'ExtensionName' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function ExtensionEnabled(ByRef App As ESRI.ArcGIS.Framework.IApplication, ByRef ExtensionName As String) As Boolean
		' UPGRADE_INFO (#05B1): The 'c_ModuleFileName' variable wasn't declared explicitly.
		Dim c_ModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler
		
		'  Determine whether the extension specified is enabled or not.
		
		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension
		Dim pExtConfig As ESRI.ArcGIS.esriSystem.IExtensionConfig
		
34:
		pExt = FindMXExtByName(App, ExtensionName)
35:
		If pExt Is Nothing Then
36:
			ExtensionEnabled = False
37:
		Else
38:
			pExtConfig = pExt
39:
			If pExtConfig Is Nothing Then
40:
				ExtensionEnabled = False
41:
			Else
42:
				ExtensionEnabled = pExtConfig.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled
43:
			End If
44:
		End If
        'saeid
        ' ExtensionEnabled = True
		Exit Function
		
ErrorHandler:
		HandleError(False, "ExtensionEnabled " & c_ModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
50:
		Return False
	End Function

	Public Function GetOutputFolder() As String
		' UPGRADE_INFO (#05B1): The 'c_sModuleFileName' variable wasn't declared explicitly.
		Dim c_sModuleFileName As Object = Nothing
		On Error GoTo ErrorHandler

		Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
		Dim pObjectFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		Dim pEnumSelect As ESRI.ArcGIS.Catalog.IEnumGxObject
		Dim pgxobject As ESRI.ArcGIS.Catalog.IGxObject
		
136:
		pGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog()
137:
		pObjectFilter = New FolderFilter()
		
139:
		With pGxDialog
140:
			.RememberLocation = True
141:
			.ObjectFilter = pObjectFilter
142:
			.AllowMultiSelect = False
143:
			.Title = "Destination Folder"
144:
			.ButtonCaption = "Open"
145:
		End With
		
147:
		If pGxDialog.DoModalOpen(0, pEnumSelect) Then
148:
			pgxobject = pEnumSelect.Next()
149:
			GetOutputFolder = pgxobject.FullName
150:
		Else
151:
			GetOutputFolder = ""
152:
		End If

		Exit Function
ErrorHandler:
		HandleError(False, "GetOutputFolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

    
    ' UPGRADE_INFO (#0551): The 'requiredfolder' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Function Get_ProgramFiles_Folder(Optional ByRef requiredfolder As String = "") As String
		'********************************************************************''
		'Date: 10 / 05 / 06
		'Author: Nathan Eaton, CALM
		'Description: This functions returns the program files folder
		'Accepts: Optional required folder to exist under the program files folder
		'Returns: True if the directory exists, false if the directory does not exist
		'********************************************************************
		
		On Error GoTo errh
		
		Dim Drivelist As New Collection
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
		Dim d As Object = Nothing
		
		Get_ProgramFiles_Folder = "C:\Program Files"
		For Each d In Drivelist
			
			If modCommon_Functions.Direxists(d & ":\Program Files") Then
				If requiredfolder <> "" Then
					If modCommon_Functions.Direxists(d & ":\Program Files\" & requiredfolder) Then
						Return d & ":\Program Files"
						
					End If
				Else
					
					Return d & ":\Program Files"
					
				End If
			End If
			
		Next
		
		Exit Function
		
errh:
		Return "C:\Program Files"
	End Function

End Module
