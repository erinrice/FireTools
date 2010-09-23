' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

#Region "Global migration warnings"

' UPGRADE_INFO (#06E1): Current project references the 'ErrorHandlerUI' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'ADODB' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'dao' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'SHDocVw' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Outlook' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'OutlookViewCtl' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'RegObj' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Catalog' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Framework' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'System' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Carto' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Geodatabase' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Editor' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'ArcMapUI' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Display' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Geometry' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'SystemUI' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'CatalogUI' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'GeoDatabaseUI' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'DataSourcesGDB' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'DataSourcesFile' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'Output' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'DataSourcesRaster' COM type library.
' UPGRADE_INFO (#06E1): Current project references the 'MSHierarchicalFlexGridLib' COM type library.
' UPGRADE_INFO (#07F1): The 'dao' type library is never used in current project. Consider deleting it from VB6 project references.
' UPGRADE_INFO (#07F1): The 'SHDocVw' type library is never used in current project. Consider deleting it from VB6 project references.
' UPGRADE_INFO (#07F1): The 'OutlookViewCtl' type library is never used in current project. Consider deleting it from VB6 project references.
' UPGRADE_INFO (#07F1): The 'RegObj' type library is never used in current project. Consider deleting it from VB6 project references.

#End Region

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports VB = Microsoft.VisualBasic

Public Enum ERegistryClassConstants
	HKEY_CLASSES_ROOT = &H80000000
	HKEY_CURRENT_USER = &H80000001
	HKEY_LOCAL_MACHINE = &H80000002
	HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
	'Predefined Value Types
	REG_NONE = (0) 'No value type
	REG_SZ = (1) 'Unicode nul terminated string
	REG_EXPAND_SZ = (2) 'Unicode nul terminated string w/enviornment var
	REG_BINARY = (3) 'Free form binary
	REG_DWORD = (4) '32-bit number
	REG_DWORD_LITTLE_ENDIAN = (4) '32-bit number (same as REG_DWORD)
	REG_DWORD_BIG_ENDIAN = (5) '32-bit number
	REG_LINK = (6) 'Symbolic Link (unicode)
	REG_MULTI_SZ = (7) 'Multiple Unicode strings
	REG_RESOURCE_LIST = (8) 'Resource list in the resource map
	REG_FULL_RESOURCE_DESCRIPTOR = (9) 'Resource list in the hardware description
	REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Friend Class Registry

	' =========================================================
	' Class:    cRegistry
	' Author:   Steve McMahon
	' Date  :   21 Feb 1997
	'
	' A nice class wrapper around the registry functions
	' Allows searching,deletion,modification and addition
	' of Keys or Values.
	'
	' Updated 29 April 1998 for VB5.
	'   * Fixed GPF in EnumerateValues
	'   * Added support for all registry types, not just strings
	'   * Put all declares in local class
	'   * Added VB5 Enums
	'   * Added CreateKey and DeleteKey methods
	'
	' Updated 2 January 1999
	'   * The CreateExeAssociation method failed to set up the
	'     association correctly if the optional document icon
	'     was not provided.
	'   * Added new parameters to CreateExeAssociation to set up
	'     other standard handlers: Print, Add, New
	'   * Provided the CreateAdditionalEXEAssociations method
	'     to allow non-standard menu items to be added (for example,
	'     right click on a .VBP file.  VB installs Run and Make
	'     menu items).
	'
	' Updated 8 February 2000
	'   * Ensure CreateExeAssociation and related items sets up the
	'     registry keys in the
	'           HKEY_LOCAL_MACHINE\SOFTWARE\Classes
	'     branch as well as the HKEY_CLASSES_ROOT branch.
	'
	' ---------------------------------------------------------------------------
	' vbAccelerator - free, advanced source code for VB programmers.
	'     http://vbaccelerator.com
	' =========================================================
	

	Private m_hClassKey As Integer
	Private m_sSectionKey As String = ""

	Public Property ClassKey() As ERegistryClassConstants
		Get
			Return m_hClassKey
	 	End Get
		Set(ByVal eKey As ERegistryClassConstants)
			m_hClassKey = eKey
	 	End Set
	End Property

	Public Property SectionKey() As String
		Get
			Return m_sSectionKey
	 	End Get
		Set(ByVal sSectionKey As String)
			m_sSectionKey = sSectionKey
	 	End Set
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
