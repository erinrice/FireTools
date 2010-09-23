' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs

<ComClass(clsFire_Log_Dockwin.ClassId, clsFire_Log_Dockwin.InterfaceId, clsFire_Log_Dockwin.EventsId), _
ProgId("CALM_FireTools.clsFire_Log_Dockwin")> _
<VB6Object("CALM_FireTools.clsFire_Log_Dockwin")> _
Public Class clsFire_Log_Dockwin
    Implements Framework.IDockableWindowDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "04eff954-80d7-4f73-a09f-890fc2eeca8d"
    Public Const InterfaceId As String = "e14f681b-23a8-4ddc-b13a-27a3c60e649d"
    Public Const EventsId As String = "9c235175-c6be-438f-a366-6af461e6c222"
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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxDockableWindows.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxDockableWindows.Unregister(regKey)

    End Sub

#End Region
#End Region




    ' UPGRADE_INFO (#0501): The 'clsFire_Log_Dockwin' member isn't used anywhere in current application.

    Public ReadOnly Property IDockableWindowDef_Caption() As String Implements Framework.IDockableWindowDef.Caption
        Get
            Return "Fire Log Details"
        End Get
    End Property

    Public ReadOnly Property IDockableWindowDef_ChildHWND() As Integer Implements Framework.IDockableWindowDef.ChildHWND
        Get
            ' The dockable window will consist of a list box, so pass back the hWnd of
            ' the listbox on a form
            Return frmFire_Log_Details.TreeView1.hWnd
        End Get
    End Property

    Public ReadOnly Property IDockableWindowDef_Name() As String Implements Framework.IDockableWindowDef.Name
        Get
            Return "Fire_Log_Window"
        End Get
    End Property

    Public Sub IDockableWindowDef_OnCreate(ByVal hook As Object) Implements Framework.IDockableWindowDef.OnCreate

    End Sub

    Public Sub IDockableWindowDef_OnDestroy() Implements Framework.IDockableWindowDef.OnDestroy

    End Sub

    Public ReadOnly Property IDockableWindowDef_UserData() As Object Implements Framework.IDockableWindowDef.UserData
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
