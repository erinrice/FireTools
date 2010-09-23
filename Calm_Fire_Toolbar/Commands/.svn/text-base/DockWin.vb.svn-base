' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.DockWin")> _ 
<VB6Object("CALM_FireTools.DockWin")> _
Public Class DockWin
	Implements Framework.IDockableWindowDef
' UPGRADE_INFO (#0501): The 'DockWin' member isn't used anywhere in current application.

	Private ReadOnly Property IDockableWindowDef_Caption() As String Implements Framework.IDockableWindowDef.Caption
		Get
			Return "Fire Log Details"
	 	End Get
	End Property

	Private ReadOnly Property IDockableWindowDef_ChildHWND() As Integer Implements Framework.IDockableWindowDef.ChildHWND
		Get
			' The dockable window will consist of a list box, so pass back the hWnd of
			' the listbox on a form
			Return frmFire_Log_Details.TreeView1.hWnd
	 	End Get
	End Property

	Private ReadOnly Property IDockableWindowDef_Name() As String Implements Framework.IDockableWindowDef.Name
		Get
			Return "Fire_Log_Window"
	 	End Get
	End Property

	Private Sub IDockableWindowDef_OnCreate(ByVal hook As Object) Implements Framework.IDockableWindowDef.OnCreate
		
	End Sub

	Private Sub IDockableWindowDef_OnDestroy() Implements Framework.IDockableWindowDef.OnDestroy
		
	End Sub

	Private ReadOnly Property IDockableWindowDef_UserData() As Object Implements Framework.IDockableWindowDef.UserData
		Get
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
