' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs

<ComClass(clsFire_Labels_Tool.ClassId, clsFire_Labels_Tool.InterfaceId, clsFire_Labels_Tool.EventsId), _
System.Runtime.InteropServices.ProgId("CALM_FireTools.clsFire_Labels_Tool")> _
<VB6Object("CALM_FireTools.clsFire_Labels_Tool")> _
Public Class clsFire_Labels_Tool
    Implements SystemUI.ICommand
    Implements SystemUI.ITool
    ' UPGRADE_INFO (#0501): The 'clsFire_Labels_Tool' member isn't used anywhere in current application.

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "31235c19-9859-41fe-90e2-d5095c8846aa"
    Public Const InterfaceId As String = "5815153d-0606-4df9-8188-ebac92e7b15c"
    Public Const EventsId As String = "1b7f20e8-cd6b-4cf9-bd7a-2f21df9e8e06"
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
    Private m_pApp As ESRI.ArcGIS.Framework.IApplication
    ' UPGRADE_INFO (#0701): The 'm_bInUse' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_bInUse As Boolean
    ' UPGRADE_INFO (#0701): The 'm_pLineSymbol' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_pLineSymbol As ESRI.ArcGIS.Display.ILineSymbol
    ' UPGRADE_INFO (#0701): The 'm_pLinePolyline' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_pLinePolyline As ESRI.ArcGIS.Geometry.IPolyline
    ' UPGRADE_INFO (#0701): The 'm_pTextSymbol' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_pTextSymbol As ESRI.ArcGIS.Display.ITextSymbol
    ' UPGRADE_INFO (#0701): The 'm_pStartPoint' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_pStartPoint As ESRI.ArcGIS.Geometry.IPoint
    ' UPGRADE_INFO (#0701): The 'm_pTextPoint' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private m_pTextPoint As ESRI.ArcGIS.Geometry.IPoint

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
        Get
            'ICommand_Bitmap = frmResources.imlBitmaps.ListImages(1).Picture
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
        Get
            Return " "
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
        Get
            'ICommand_Category = "Developer Samples"
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
        Get
        End Get
    End Property

    Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
        Get
            'Dim bEnabled As Boolean = False

            'bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            'If modFire_Tools.Fire_Anno_Exists() = True Then
            '	bEnabled = True
            'Else
            '	bEnabled = False
            'End If
            'If g_Username = "" Then
            '	bEnabled = False
            'End If

            'Return bEnabled
            Return True
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
        Get
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
        Get
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
        Get
            '  ICommand_Message = "Measure Distance Tool"
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
        Get
            '  ICommand_Name = "Developer Samples_Measure Tool"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
    End Sub

    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
        m_pApp = hook
    End Sub

    Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
        Get
            '  ICommand_Tooltip = "Measure Tool"
        End Get
    End Property

    Private ReadOnly Property ITool_Cursor() As Integer Implements SystemUI.ITool.Cursor
        Get
            'ITool_Cursor = frmResources.imlBitmaps.ListImages(2).Picture
        End Get
    End Property

    Private Function ITool_Deactivate() As Boolean Implements SystemUI.ITool.Deactivate
        ' stop doing operation
        'Set m_pTextSymbol = Nothing
        'Set m_pTextPoint = Nothing
        'Set m_pLinePolyline = Nothing
        'Set m_pLineSymbol = Nothing
        'm_bInUse = False
        '  If frmFire_Labels.Visible = True Then
        '    frmFire_Labels.Hide
        '  End If

        Return True

    End Function

    Private Function ITool_OnContextMenu(ByVal X As Integer, ByVal y As Integer) As Boolean Implements SystemUI.ITool.OnContextMenu
    End Function

    Private Sub ITool_OnDblClick() Implements SystemUI.ITool.OnDblClick
    End Sub

    Private Sub ITool_OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer) Implements SystemUI.ITool.OnKeyDown
    End Sub

    Private Sub ITool_OnKeyUp(ByVal keyCode As Integer, ByVal Shift As Integer) Implements SystemUI.ITool.OnKeyUp
    End Sub

    Private Sub ITool_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer) Implements SystemUI.ITool.OnMouseDown
        Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
        Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView
        Dim pMap As ESRI.ArcGIS.Carto.IMap = pMxDoc.FocusMap
        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint
        Dim pGraCont As ESRI.ArcGIS.Carto.IGraphicsContainer = pMxDoc.ActiveView
        ' UPGRADE_INFO (#0701): The 'pGraContSel' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim pGraContSel As ESRI.ArcGIS.Carto.IGraphicsContainerSelect
        Dim pRubberPoint As ESRI.ArcGIS.Display.IRubberBand = New ESRI.ArcGIS.Display.RubberPoint()
        Dim pElem As ESRI.ArcGIS.Carto.IElement

        ' QI for the MXDocument interface
        pActiveView = pMap

        ' QI for the IGraphicsContainerSelect interface on the document's activeview

        ' Create a new RubberPolygon

        ' Return a new Polygon from the tracker object using TrackNew with the new symbol
        pPoint = pRubberPoint.TrackNew(pMxDoc.ActiveView.ScreenDisplay, Nothing)
        pPoint.SpatialReference = pMap.SpatialReference
        Dim labeltext As String = ""
        Dim mydate As Date
        ' UPGRADE_INFO (#0701): The 'myhourint' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim myhourint As Short
        ' UPGRADE_INFO (#0701): The 'myminuteint' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim myminuteint As Short
        Dim mydatestring As String = ""
        Dim mytimestring As String = ""
        ' UPGRADE_INFO (#0701): The 'MyMonthtext' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim MyMonthtext As Object = Nothing
        ' UPGRADE_INFO (#0701): The 'MyDatetext' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim MyDatetext As Object = Nothing
        ' UPGRADE_INFO (#0561): The 'myhour' symbol was defined without an explicit "As" clause.
        Dim myhour As Object = Nothing
        ' UPGRADE_INFO (#0561): The 'myminute' symbol was defined without an explicit "As" clause.
        Dim myminute As Object = Nothing
        ' UPGRADE_INFO (#0561): The 'MyDay' symbol was defined without an explicit "As" clause.
        Dim MyDay As Object = Nothing
        ' UPGRADE_INFO (#0561): The 'MyMonth' symbol was defined without an explicit "As" clause.
        Dim MyMonth As Object = Nothing
        Dim MyYear As String = ""
        Dim pExt As ESRI.ArcGIS.esriSystem.IExtension
        Dim myIncidentInfo As clsFire_Incident_Info
        Dim pAnnolayer As ESRI.ArcGIS.Carto.IFDOGraphicsLayer
        Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
        Dim pTransactions As ESRI.ArcGIS.Geodatabase.ITransactions
        Dim pannofeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pannofeatclass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pannoclass As ESRI.ArcGIS.Carto.IAnnoClass
        Dim pannoclassadmin As ESRI.ArcGIS.Carto.IAnnoClassAdmin
        Dim refscalenum As Double
        Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
        Dim elementcol As ESRI.ArcGIS.Carto.IElementCollection
        Dim z As Integer
        If pPoint IsNot Nothing Then
            Dim myForm As frmFire_Labels = clsFire_Labels.loadedForm
            ' Create a new PolygonElement and set its Geometry
            If myForm.cb_CustomText.CheckState = 0 Then
                labeltext = myForm.lblList.Text
                mydate = Now
                myhour = DatePart("h", mydate)
                If Len6(myhour) = 1 Then
                    myhour = "0" & myhour
                End If
                myminute = DatePart("n", mydate)
                If Len6(myminute) = 1 Then
                    myminute = "0" & myminute
                End If
                MyDay = DatePart("d", mydate)
                MyMonth = DatePart("m", mydate)
                MyYear = DatePart("yyyy", mydate)
                mydatestring = MyDay & "/" & MyMonth & "/" & MyYear
                mytimestring = myhour & myminute & "hrs"
                If labeltext = "Date and Time" Then
                    labeltext = mydatestring & " - " & mytimestring
                ElseIf labeltext = "Date" Then
                    labeltext = mydatestring
                ElseIf labeltext = "Time" Then
                    labeltext = mytimestring
                End If
            Else
                labeltext = myForm.tbCustomText.Text
            End If

            pExt = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)

            If TypeOf pExt Is clsFire_Incident_Info Then
                myIncidentInfo = pExt
            End If

            pAnnolayer = modFire_Tools.Get_Fire_Anno_Layer()

            pAnnolayer.BeginAddElements()

            pannofeatlayer = pAnnolayer

            pannofeatclass = pannofeatlayer.FeatureClass

            'set the referencescale to the current scale of the map
            pannoclass = pannofeatclass.Extension

            pannoclassadmin = pannoclass

            refscalenum = pMap.MapScale
            pannoclassadmin.ReferenceScale = refscalenum 'pMap.ReferenceScale

            pDataset = pannofeatclass

            ' Inline QI to ITransactions
            pTransactions = pDataset.Workspace
            pTransactions.StartTransaction()

            pGeoDataset = pannofeatlayer

            pPoint.Project(pGeoDataset.SpatialReference)

            pElem = MakeATextElement(pPoint, labeltext)

            elementcol = New ESRI.ArcGIS.Carto.ElementCollection()
            elementcol.Add(pElem)

            pAnnolayer.DoAddElements(elementcol, z)

            pAnnolayer.EndAddElements()

            pTransactions.CommitTransaction()

            pActiveView = pMap
            pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography, Nothing, Nothing) 'refresh the graphics part of the map
        End If

    End Sub

    ' UPGRADE_INFO (#0551): The 'pPoint' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    ' UPGRADE_INFO (#0551): The 'strText' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
    Public Function MakeATextElement(ByRef pPoint As ESRI.ArcGIS.Geometry.IPoint, ByRef strText As String) As ESRI.ArcGIS.Carto.IElement

        'Pointers needed to make text element
        Dim pRGBcolor As ESRI.ArcGIS.Display.IRgbColor = New ESRI.ArcGIS.Display.RgbColor()
        Dim pTextElement As ESRI.ArcGIS.Carto.ITextElement = New ESRI.ArcGIS.Carto.TextElement()
        Dim pTextSymbol As ESRI.ArcGIS.Display.ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol()
        ' UPGRADE_INFO (#0701): The 'fnt' member has been removed because it isn't used anywhere in current application.
        ' EXCLUDED: Dim fnt As Font
        Dim pElement As ESRI.ArcGIS.Carto.IElement = pTextElement

        'First setup a color.  We'll use RGB red
        pRGBcolor.Blue = 0
        pRGBcolor.Red = 0
        pRGBcolor.Green = 0

        'Next, cocreate a new TextElement
        'Query Interface (QI) to an IElement pointer and set
        'the geometry that was passed in
        pElement.Geometry = pPoint

        'Next, setup a font
        Dim pFontDisp As Font = New Font("Arial", 8)
        FontChangeBold6(pFontDisp, True)
        FontChangeName6(pFontDisp, "Tahoma")

        'Next, setup a TextSymbol that the TextElement will draw with
        FontChangeName6(pTextSymbol.Font, pFontDisp.Name)
        pTextSymbol.Color = pRGBcolor
        pTextSymbol.Size = 11 'set the size of the text symbol here, rather than on the font

        Dim pMask As ESRI.ArcGIS.Display.IMask = pTextSymbol
        pMask.MaskStyle = ESRI.ArcGIS.Display.esriMaskStyle.esriMSHalo
        pMask.MaskSize = 2

        'Next, Give the TextSymbol and text string to the TextElement
        pTextElement.Symbol = pTextSymbol
        pTextElement.Text = strText

        'Finally, hand back the new element as the output of this function
        Return pTextElement

    End Function

    Private Sub ITool_OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer) Implements SystemUI.ITool.OnMouseMove
    End Sub

    Private Sub ITool_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal y As Integer) Implements SystemUI.ITool.OnMouseUp
    End Sub

    Private Sub ITool_Refresh(ByVal hDC As Integer) Implements SystemUI.ITool.Refresh
    End Sub

    ' UPGRADE_INFO (#0701): The 'GetSmashedLine' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Private Function GetSmashedLine(ByRef pDisplay As ESRI.ArcGIS.Display.IScreenDisplay, ByRef pTextSymbol As ESRI.ArcGIS.Display.ISymbol, ByRef pPoint As ESRI.ArcGIS.Geometry.IPoint, ByRef pPolyline As ESRI.ArcGIS.Geometry.IPolyline) As ESRI.ArcGIS.Geometry.IPolyline

    ' EXCLUDED: 'Returns a Polyline with a blank space for the text to go in
    ' UPGRADE_INFO (#0701): The 'pSmashed' member has been removed because it isn't used anywhere in current application.
    ' EXCLUDED: Dim pSmashed As ESRI.ArcGIS.Geometry.IPolyline
    ' EXCLUDED: Dim pBoundary As ESRI.ArcGIS.Geometry.IPolygon = New ESRI.ArcGIS.Geometry.Polygon()
    ' EXCLUDED: pTextSymbol.QueryBoundary(pDisplay.hDC, pDisplay.DisplayTransformation, pPoint, pBoundary)
    ' EXCLUDED: Dim pTopo As ESRI.ArcGIS.Geometry.ITopologicalOperator = pBoundary

    ' EXCLUDED: Dim pIntersect As ESRI.ArcGIS.Geometry.IPolyline = pTopo.Intersect(pPolyline, ESRI.ArcGIS.Geometry.esriGeometryDimension.esriGeometry1Dimension)
    ' EXCLUDED: pTopo = pPolyline
    ' EXCLUDED: Return pTopo.Difference(pIntersect)

    ' EXCLUDED: End Function

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
