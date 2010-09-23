' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

'  Interface through which the FIRE extension shares information
'  about the current incident, datasets, etc. with tools
'  created for the extension.

' Enumeration for Layers

<VB6Object("CALM_FireTools.Fire_Layers")> _
Public Enum Fire_Layers
	FIRELYR_PERIMETER
	FIRELYR_ZONE
	FIRELYR_BRANCH
	FIRELYR_DIVISION
	FIRELYR_SECTOR
	FIRELYR_ASSIGNMENTBREAK
	FIRELYR_FIREPOINT
	FIRELYR_FIRELINE
	FIRELYR_PERIMETER_HIS
	FIRELYR_ZONE_HIS
	FIRELYR_BRANCH_HIS
	FIRELYR_DIVISION_HIS
	FIRELYR_SECTOR_HIS
	FIRELYR_ASSIGNMENTBREAK_HIS
	FIRELYR_FIREPOINT_HIS
	FIRELYR_FIRELINE_HIS
End Enum

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Incident_Info")> _ 
<VB6Object("CALM_FireTools.clsFire_Incident_InfoClass")> _
Public Class clsFire_Incident_InfoClass
	Implements clsFire_Incident_Info

	'   Name of this fire or incident.

	Public Property incidentName() As String Implements clsFire_Incident_Info.incidentName
		Get
	 	End Get
		Set(ByVal Name As String)
	 	End Set
	End Property

	'  Incident number.

	Public Property IncidentNumber() As String Implements clsFire_Incident_Info.IncidentNumber
		Get
	 	End Get
		Set(ByVal Number As String)
	 	End Set
	End Property

	'  Property to get/set any of the Fire Tool dependent layers.

	Public Property IncidentLayer(ByVal elayer As Fire_Layers) As ESRI.ArcGIS.Carto.IFeatureLayer Implements clsFire_Incident_Info.IncidentLayer
		Get
	 	End Get
		Set(ByVal player As ESRI.ArcGIS.Carto.IFeatureLayer)
	 	End Set
	End Property

	'  Workspace (Geodatabase) containing the incident layers.
	'  Additional background data can be anywhere.

	Public Property IncidentWorkspace() As ESRI.ArcGIS.Geodatabase.IWorkspace Implements clsFire_Incident_Info.IncidentWorkspace
		Get
	 	End Get
		Set(ByVal Workspace As ESRI.ArcGIS.Geodatabase.IWorkspace)
	 	End Set
	End Property

	'  Incident Region

	Public Property incidentRegion() As String Implements clsFire_Incident_Info.incidentRegion
		Get
	 	End Get
		Set(ByVal [Region] As String)
	 	End Set
	End Property

	'  Incident Region Ex

	Public Property IncidentRegionEx() As String Implements clsFire_Incident_Info.IncidentRegionEx
		Get
	 	End Get
		Set(ByVal Reg_ex As String)
	 	End Set
	End Property

	'  Incident WorkCentre

	Public Property IncidentWorkCentre() As String Implements clsFire_Incident_Info.IncidentWorkCentre
		Get
	 	End Get
		Set(ByVal WorkCentre As String)
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

<VB6Object("CALM_FireTools.clsFire_Incident_Info")> _
<VB6Annotation(UseClassSuffix:=True)> _ 
Public Interface clsFire_Incident_Info
	Property incidentName() As String
	Property IncidentNumber() As String
	Property IncidentLayer(ByVal elayer As Fire_Layers) As ESRI.ArcGIS.Carto.IFeatureLayer
	Property IncidentWorkspace() As ESRI.ArcGIS.Geodatabase.IWorkspace
	Property incidentRegion() As String
	Property IncidentRegionEx() As String
	Property IncidentWorkCentre() As String
End Interface

