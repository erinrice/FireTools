Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

<ComClass(BoundaryFill.ClassId, BoundaryFill.InterfaceId, BoundaryFill.EventsId), _
 ProgId("CALM_FireTools.BoundaryFill")> _
Public NotInheritable Class BoundaryFill
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "88fc0d2b-04be-42b7-a7cd-701f9938897f"
    Public Const InterfaceId As String = "49ad197e-4be1-4514-98a2-a99a592906ab"
    Public Const EventsId As String = "78bd045f-f379-452b-8f01-d0b3c40b55b0"
#End Region

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
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private m_application As IApplication
    Public m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = "Change Boundary Fill Status"   'localizable text 
        MyBase.m_message = "Change Boundary Fill Status"   'localizable text 
        MyBase.m_toolTip = "Change Boundary Fill Status" 'localizable text 
        MyBase.m_name = "BoundaryFill"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            m_application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
        m_pDoc = m_application.Document
        modCommon_Functions.m_pApp = m_application

        ' TODO:  Add other initialization code
    End Sub

    'Public Overrides ReadOnly Property Enabled() As Boolean
    '    Get
    '        Dim bEnabled As Boolean = False
    '        bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
    '        bEnabled = Incident_Workspace_There()
    '        If g_Username = "" Then
    '            bEnabled = False
    '        End If
    '        Return bEnabled
    '    End Get
    'End Property

    Public Overrides Sub OnClick()
        Dim bEnabled As Boolean = False
        bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
        bEnabled = Incident_Workspace_There()
        If g_Username = "" Then
            bEnabled = False
        End If
        If bEnabled = False Then
            MessageBox.Show(IncidentMessage)
            Exit Sub
        End If

        Dim frm As New frmBoundaryFill()
        frm.ShowDialog()

        'TODO: Add ChangeBoundaryFillStatus.OnClick implementation
    End Sub
End Class



