' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.cValuePairs")> _ 
<VB6Object("CALM_FireTools.cValuePairs")> _
Public Class cValuePairs
	Implements System.Collections.IEnumerable

	' The private collection used to hold the real data
	Private m_cValuePairs As Collection = New Collection()
	
	
	' Implement support for enumeration (For Each)
	
	<System.Runtime.InteropServices.DispId(-4)> _
	<System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)> _
	Public Function NewEnum() As IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		' delegate to the private collection
		Return m_cValuePairs.GetEnumerator()
	End Function

	#Region "Constructor"
	
	'A default constructor with Friend scope to prevent instantiation from outside the assembly
	Protected Friend Sub New()
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
