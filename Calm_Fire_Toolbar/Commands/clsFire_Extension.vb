' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs

Imports ESRI.ArcGIS.esriSystem

<ComClass(clsFire_Extension.ClassId, clsFire_Extension.InterfaceId, clsFire_Extension.EventsId), _
System.Runtime.InteropServices.ProgId("CALM_FireTools.clsFire_Extension")> _
<VB6Object("CALM_FireTools.clsFire_Extension")> _
Public Class clsFire_Extension
    Implements IExtension
    Implements clsFire_Incident_Info
    Implements IExtensionConfig

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Unregister(regKey)

    End Sub

#End Region
#End Region
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "161f2cdf-6ec9-4dc1-94eb-02da402fb4b3"
    Public Const InterfaceId As String = "048f8fd9-941a-4160-b52c-3872039fb3c9"
    Public Const EventsId As String = "9ccc8eec-fd27-4d24-b8bc-b64936b2432c"
#End Region

    ' UPGRADE_INFO (#0501): The 'clsFire_Extension' member isn't used anywhere in current application.

    ' UPGRADE_INFO (#0561): The 'c_sModuleFileName' symbol was defined without an explicit "As" clause.
    Private Const c_sModuleFileName As String = "C:\eats\Scripts\CALM\VB\Development\CALM_Extension\clsFire_Extension.cls"

    Private g_IncidentName As String = ""
    Private g_IncidentLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private g_IncidentNumber As String = ""
    Private g_IncidentWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
    Private g_IncidentRegion As String = ""
    Private g_IncidentRegion_ext As String = ""
    Private g_IncidentWorkCentre As String = ""

    Private mConfigState As ESRI.ArcGIS.esriSystem.esriExtensionState

    Private Property clsFire_Incident_Info_IncidentLayer(ByVal elayer As Fire_Layers) As ESRI.ArcGIS.Carto.IFeatureLayer Implements clsFire_Incident_Info.IncidentLayer
        Get
            Return g_IncidentLayer
        End Get
        Set(ByVal RHS As ESRI.ArcGIS.Carto.IFeatureLayer)
            g_IncidentLayer = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentName() As String Implements clsFire_Incident_Info.incidentName
        Get
            Return g_IncidentName
        End Get
        Set(ByVal RHS As String)
            g_IncidentName = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentNumber() As String Implements clsFire_Incident_Info.IncidentNumber
        Get
            Return g_IncidentNumber
        End Get
        Set(ByVal RHS As String)
            g_IncidentNumber = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentRegion() As String Implements clsFire_Incident_Info.incidentRegion
        Get
            Return g_IncidentRegion
        End Get
        Set(ByVal RHS As String)
            g_IncidentRegion = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentWorkspace() As ESRI.ArcGIS.Geodatabase.IWorkspace Implements clsFire_Incident_Info.IncidentWorkspace
        Get
            Return g_IncidentWorkspace
        End Get
        Set(ByVal RHS As ESRI.ArcGIS.Geodatabase.IWorkspace)
            g_IncidentWorkspace = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentRegionEx() As String Implements clsFire_Incident_Info.IncidentRegionEx
        Get
            Return g_IncidentRegion_ext
        End Get
        Set(ByVal RHS As String)
            g_IncidentRegion_ext = RHS
        End Set
    End Property

    Private Property clsFire_Incident_Info_IncidentWorkCentre() As String Implements clsFire_Incident_Info.IncidentWorkCentre
        Get
            Return g_IncidentWorkCentre
        End Get
        Set(ByVal RHS As String)
            g_IncidentWorkCentre = RHS
        End Set
    End Property

    Public ReadOnly Property IExtension_Name() As String Implements IExtension.Name
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return g_FIRE_EXTENSION

ErrorHandler:
            HandleError(True, "IExtension_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Public Sub IExtension_Startup(ByRef initializationData As Object) Implements IExtension.Startup
        On Error GoTo ErrorHandler

        ' TODO: Add your implementation here

        Exit Sub
ErrorHandler:
        HandleError(True, "IExtension_Startup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
    End Sub

    Public Sub IExtension_Shutdown() Implements IExtension.Shutdown
        On Error GoTo ErrorHandler

        ' TODO: Add your implementation here

        Exit Sub
ErrorHandler:
        HandleError(True, "IExtension_Shutdown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
    End Sub

    Public ReadOnly Property IExtensionConfig_ProductName() As String Implements IExtensionConfig.ProductName
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return g_FIRE_EXTENSION

ErrorHandler:
            HandleError(True, "IExtensionConfig_ProductName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Public ReadOnly Property IExtensionConfig_Description() As String Implements IExtensionConfig.Description
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return "DEC Fire Extension Tools for Fire Incident Mapping" & ControlChars.NewLine & "Version 2.0, 2010 fire season(ArcGis 9.3.1 version)"

ErrorHandler:
            HandleError(True, "IExtensionConfig_Description " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Public Property IExtensionConfig_State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements IExtensionConfig.State
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return mConfigState

ErrorHandler:
            HandleError(True, "IExtensionConfig_State " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
        Set(ByVal ExtensionState As ESRI.ArcGIS.esriSystem.esriExtensionState)
            On Error GoTo ErrorHandler

            If mConfigState <> ESRI.ArcGIS.esriSystem.esriExtensionState.esriESUnavailable Then
                mConfigState = ExtensionState
            End If

            Exit Property
ErrorHandler:
            HandleError(True, "IExtensionConfig_State " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
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
