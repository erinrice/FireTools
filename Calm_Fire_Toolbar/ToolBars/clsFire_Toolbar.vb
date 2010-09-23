' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Explicit Off
Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports ESRI.ArcGIS.ADF.CATIDs
Imports System.Runtime.InteropServices

<ComClass(clsFire_Toolbar.ClassId, clsFire_Toolbar.InterfaceId, clsFire_Toolbar.EventsId), _
ProgId("CALM_FireTools.clsFire_Toolbar")> _
<VB6Object("CALM_FireTools.clsFire_Toolbar")> _
Public Class clsFire_Toolbar
    Implements SystemUI.IToolBarDef
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
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "cb5c9fbc-df16-4198-9730-dfa5dd415877"
    Public Const InterfaceId As String = "ee74f5dd-25bd-439c-aa7e-837aec7277b5"
    Public Const EventsId As String = "479fb503-b637-4bd7-bc17-92cdef8c4660"
#End Region



    ' UPGRADE_INFO (#0501): The 'clsFire_Toolbar' member isn't used anywhere in current application.

    Public ReadOnly Property IToolBarDef_ItemCount() As Integer Implements SystemUI.IToolBarDef.ItemCount
        Get
            ' Two items on the toolbar
            Return 12 '13
        End Get
    End Property

    Public Sub IToolBarDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements SystemUI.IToolBarDef.GetItemInfo
        ' Add the sample command and the built-in Add Data command to the toolbar

        Select Case pos
            Case 0
                itemDef.ID = "CALM_FireTools.clsFire_Logon"
            Case 1
                itemDef.ID = "CALM_FireTools.clsFire_Incident_Menu"
            Case 2
                itemDef.ID = "CALM_FireTools.clsFire_Capture_Poly"
                itemDef.Group = True
            Case 3
                itemDef.ID = "CALM_FireTools.clsFire_ModifyPly"
            Case 4
                itemDef.ID = "CALM_FireTools.clsFire_Copy_2_Ply"
            Case 5
                itemDef.ID = "CALM_FireTools.clsFire_Capture_Line"
                itemDef.Group = True
            Case 6
                itemDef.ID = "CALM_FireTools.clsFire_Copy_2_Line"
            Case 7
                itemDef.ID = "CALM_FireTools.clsFire_Capture_Point"
                itemDef.Group = True
            Case 8
                itemDef.ID = "CALM_FireTools.clsFire_Copy_2_Point"
            Case 9
                itemDef.ID = "CALM_FireTools.clsFire_Capture_Div"
                itemDef.Group = True
            Case 10
                itemDef.ID = "CALM_FireTools.BoundaryFill"
                itemDef.Group = True

                '           Case 10
                '                itemDef.ID = "CALM_FireTools.clsFire_Labels_Tool"
                'changed 22-08-2007 by FMS request - GrantB
                '    Case 11
                '    itemDef.Id = "CALM_FireTools.clsFire_Email_Menu"
            Case 11 '12
                itemDef.ID = "CALM_FireTools.clsFire_Tools_Menu"
                itemDef.Group = True
        End Select

    End Sub

    Public ReadOnly Property IToolBarDef_Name() As String Implements SystemUI.IToolBarDef.Name
        Get
            Return "CALM_Fire_Toolbar"
        End Get
    End Property

    Public ReadOnly Property IToolBarDef_Caption() As String Implements SystemUI.IToolBarDef.Caption
        Get
            Return "DEC Fire Toolbar"
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
