' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.MySystem")> _ 
<VB6Object("CALM_FireTools.MySystem")> _
Public Class MySystem
	Implements IDisposable

    ' UPGRADE_INFO (#0561): The 'REGISTRY_PATH' symbol was defined without an explicit "As" clause.
	Private Const REGISTRY_PATH As String = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
	Private oRegistry As Registry = New Registry()
	
	
	Private mValuePairs As cValuePairs
	
    Private Sub Class_Initialize_VB6()
        oRegistry.ClassKey = ERegistryClassConstants.HKEY_LOCAL_MACHINE
        oRegistry.SectionKey = REGISTRY_PATH
    End Sub

	' UPGRADE_INFO (#0541): This 'Class_Terminate' method appears to be useless. Please delete it in the original project and restart migration.
	Private Sub Class_Terminate_VB6()
		oRegistry = Nothing
		mValuePairs = Nothing
	End Sub

	#Region "Constructor"
	
	'A public default constructor
	Public Sub New()
		Class_Initialize_VB6()
		' Add initialization code here
	End Sub
	
	#End Region

	#Region "Static constructor"
	
	' This static constructor ensures that the VB6 support library be initialized before using current class.
	Shared Sub New()
		EnsureVB6LibraryInitialization()
	End Sub
	
	#End Region
	
	Protected Overrides Sub Finalize()
		Dispose(False)
	End Sub
	
	Public Sub Dispose() Implements System.IDisposable.Dispose
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
	
	Protected Overridable Sub Dispose(ByVal disposing As Boolean)
				Class_Terminate_VB6()
	End Sub

End Class
