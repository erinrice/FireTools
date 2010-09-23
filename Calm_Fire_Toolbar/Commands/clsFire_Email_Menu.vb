' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Email_Menu")> _ 
<VB6Object("CALM_FireTools.clsFire_Email_Menu")> _
Public Class clsFire_Email_Menu
	Implements SystemUI.IMenuDef
' UPGRADE_INFO (#0501): The 'clsFire_Email_Menu' member isn't used anywhere in current application.

	'changed 28-08-2007 by FMS request - GrantB
	'commented out to test impact on app
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = "C:\eats\Scripts\CALM\VB\Development\CALM_Extension\clsFire_Incident_Menu.cls"
	' Constant reflect file module name
	
	Private ReadOnly Property IMenuDef_ItemCount() As Integer Implements SystemUI.IMenuDef.ItemCount
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return 4
			
ErrorHandler:
			HandleError(True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private Sub IMenuDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements SystemUI.IMenuDef.GetItemInfo
		On Error GoTo ErrorHandler
		Select Case pos
			Case 0
				itemDef.ID = "CALM_FireTools.clsFire_Email_View"
			Case 1
				itemDef.ID = "CALM_FireTools.clsFire_Email_Layout"
			Case 2
				itemDef.ID = "CALM_FireTools.clsFire_Email_Boundary"
			Case 3
				itemDef.ID = "CALM_FireTools.clsFire_Email_GDB"
		End Select
		
		Exit Sub
ErrorHandler:
		HandleError(True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	Private ReadOnly Property IMenuDef_Name() As String Implements SystemUI.IMenuDef.Name
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email_Menu"
			
ErrorHandler:
			HandleError(True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	 	End Get
	End Property

	Private ReadOnly Property IMenuDef_Caption() As String Implements SystemUI.IMenuDef.Caption
		Get
			On Error GoTo ErrorHandler
			
			' TODO: Add your implementation here
			Return "Email "
			
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
