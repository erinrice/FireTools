' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.FolderFilter")> _ 
<VB6Object("CALM_FireTools.FolderFilter")> _
Public Class FolderFilter
	Implements Catalog.IGxObjectFilter

	'  GXFilter to limit user to browsing and selecting file system
	'  folders.  Used by the shape file export and import functions
	'  to let the user select the source or destination folder.

	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = "FolderFilter.cls"

	Private Function IGxObjectFilter_CanChooseObject(ByVal [Object] As ESRI.ArcGIS.Catalog.IGxObject, ByRef result As ESRI.ArcGIS.Catalog.esriDoubleClickResult) As Boolean Implements Catalog.IGxObjectFilter.CanChooseObject
		On Error GoTo ErrorHandler
		
		Dim bCanDisplay As Boolean
		If TypeOf [Object] Is ESRI.ArcGIS.Catalog.IGxFolder Then
			bCanDisplay = True
		Else
			bCanDisplay = False
		End If
		
		Return bCanDisplay

ErrorHandler:
		HandleError(False, "IGxObjectFilter_CanChooseObject " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	Private Function IGxObjectFilter_CanDisplayObject(ByVal [Object] As ESRI.ArcGIS.Catalog.IGxObject) As Boolean Implements Catalog.IGxObjectFilter.CanDisplayObject
		On Error GoTo ErrorHandler

		Dim bCanDisplay As Boolean
		If TypeOf [Object] Is ESRI.ArcGIS.Catalog.IGxFolder Then
			bCanDisplay = True
		Else
			bCanDisplay = False
		End If
		
		Return bCanDisplay

ErrorHandler:
		HandleError(False, "IGxObjectFilter_CanDisplayObject " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	Private Function IGxObjectFilter_CanSaveObject(ByVal Location As ESRI.ArcGIS.Catalog.IGxObject, ByVal newObjectName As String, ByRef objectAlreadyExists As Boolean) As Boolean Implements Catalog.IGxObjectFilter.CanSaveObject
		On Error GoTo ErrorHandler
		
		' Not implemented

		Exit Function
ErrorHandler:
		HandleError(False, "IGxObjectFilter_CanSaveObject " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	Private ReadOnly Property IGxObjectFilter_Description() As String Implements Catalog.IGxObjectFilter.Description
		Get
			On Error GoTo ErrorHandler
			
			Return "Folders"
	
ErrorHandler:
			HandleError(True, "IGxObjectFilter_Description " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	 	End Get
	End Property

	Private ReadOnly Property IGxObjectFilter_Name() As String Implements Catalog.IGxObjectFilter.Name
		Get
			On Error GoTo ErrorHandler
			
			Return "Folders"
	
ErrorHandler:
			HandleError(True, "IGxObjectFilter_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
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
