' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Tools_Menu")> _ 
<VB6Object("CALM_FireTools.clsFire_Tools_Menu")> _
Public Class clsFire_Tools_Menu
	Implements SystemUI.IMenuDef
' UPGRADE_INFO (#0501): The 'clsFire_Tools_Menu' member isn't used anywhere in current application.

	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = "C:\eats\Scripts\CALM\VB\Development\CALM_Extension\clsFire_Incident_Menu.cls"
	' Constant reflect file module name
	
	Private ReadOnly Property IMenuDef_ItemCount() As Integer Implements SystemUI.IMenuDef.ItemCount
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
            Return 10 '13 - changed 22-08-2007 by FMS request - GrantB
			
ErrorHandler:
			HandleError(True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private Sub IMenuDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements SystemUI.IMenuDef.GetItemInfo
		On Error GoTo ErrorHandler
		
		Select Case pos
            Case 0
                itemDef.ID = "CALM_FireTools.clsFire_Zoomcls"
            Case 1
                itemDef.ID = "CALM_Extension.clsMap_Pro_Load_Fire"
                itemDef.Group = True
            Case 2
                itemDef.ID = "CALM_FireTools.clsFire_Export_Map"
            Case 3
                itemDef.ID = "CALM_FireTools.clsFire_UpdateMeasure"
                itemDef.Group = True
            Case 4
                itemDef.ID = "CALM_FireTools.clsFire_Labels"
            Case 5
                itemDef.ID = "CALM_FireTools.clsFire_Symbols"
            Case 6 'changed 30-08-2007 by FMS request - GrantB
                itemDef.ID = "CALM_FireTools.clsFire_Label_Mult_Flds"
            Case 7 'changed 30-08-2007 by FMS request - GrantB
                itemDef.ID = "CALM_FireTools.clsFire_Export_GDB"
            Case 8
                itemDef.ID = "CALM_FireTools.clsBackup_Incident"
            Case 9 'changed 30-08-2007 by FMS request - GrantB
                itemDef.ID = "CALM_FireTools.clsBackup_HeadOffice"
                'changed 22-08-2007 by FMS request - GrantB
                '    Case 10
                '    itemDef.Id = "CALM_FireTools.clsShow_FireLog"
                '    Case 11
                '    itemDef.Id = "CALM_FireTools.clsTrim_Unique"
                '    Case 12
                '    itemDef.Id = "CALM_FireTools.clsFire_Legend_Restorer"
        End Select
		
		Exit Sub
		
ErrorHandler:
		HandleError(True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private ReadOnly Property IMenuDef_Name() As String Implements SystemUI.IMenuDef.Name
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Tools_Menu"
			
ErrorHandler:
			HandleError(True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property IMenuDef_Caption() As String Implements SystemUI.IMenuDef.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Tools "
			
ErrorHandler:
			HandleError(True, "IMenuDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

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
