Option Strict Off
Option Explicit On
Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports Microsoft.VisualBasic.Compatibility
Friend Class frmFire_Labels
    Inherits System.Windows.Forms.Form
 

#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmFire_Labels
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmFire_Labels
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmFire_Labels()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmFire_Labels)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    'variables for form resizing
    Private Structure ControlPositionType
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
        Dim Left_Renamed As Single
        'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
        Dim Top_Renamed As Single
        'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
        Dim Width_Renamed As Single
        'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
        Dim Height_Renamed As Single
        Dim FontSize As Single
    End Structure

    Private m_ControlPositions() As ControlPositionType
    Private m_FormWid As Single
    Private m_FormHgt As Single

    '******************************************************************************************************************
    'Date: 06 / 06 / 06
    'Author: Nathan Eaton, CALM
    'Description: This form specifies text to be used to create Fire Labels with
    'Called By: clsFire_labels
    'Calls: None
    'Accepts: None
    'Returns: None
    '******************************************************************************************************************

    Private Sub btnAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAdd.Click
        Dim origtext As Object
        Dim newtext As String

        newtext = InputBox("Enter New Text", "Add Text to List", "Enter Text Here")
        'UPGRADE_WARNING: Couldn't resolve default property of object origtext. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        If newtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox("Cannot add this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If

        If newtext = "" Then 'added 12/10/06 to account for cancel button
            Exit Sub
        Else
            Me.lblList.Items.Add((newtext))
        End If

    End Sub



    Private Sub btnEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEdit.Click

        Dim origtext As String
        origtext = VB6.GetItemString(Me.lblList, Me.lblList.SelectedIndex)
        If origtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox("Cannot edit this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If

        Dim edittext As String
        If Me.lblList.SelectedItems.Count > 0 Then
            edittext = InputBox("Enter New Text", "Edit Text", VB6.GetItemString(Me.lblList, Me.lblList.SelectedIndex))
            If edittext = "" Then 'added 12/10/06 to account for cancel button
                Exit Sub
            Else
                VB6.SetItemString(Me.lblList, Me.lblList.SelectedIndex, edittext)
            End If
        End If

    End Sub

    'UPGRADE_WARNING: Event cb_CustomText.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub cb_CustomText_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cb_CustomText.CheckStateChanged

        If cb_CustomText.CheckState = 1 Then
            Me.lblList.Enabled = False
            Me.btnAdd.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnEdit.Enabled = False
            Me.lblList.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000000)
            Me.tbCustomText.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        Else
            Me.btnAdd.Enabled = True
            Me.btnDelete.Enabled = True
            Me.btnEdit.Enabled = True
            Me.lblList.Enabled = True
            Me.lblList.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            Me.tbCustomText.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000000)
        End If

    End Sub

    Private Sub btnDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDelete.Click

        Dim origtext As String
        origtext = VB6.GetItemString(Me.lblList, Me.lblList.SelectedIndex)
        If origtext = "Date and Time" Or origtext = "Date" Or origtext = "Time" Then
            MsgBox("Cannot delete this text", MsgBoxStyle.Critical, "Reserved Text")
            Exit Sub
        End If
        If Me.lblList.SelectedItems.Count > 0 Then
            Me.lblList.Items.RemoveAt((Me.lblList.SelectedIndex))
        End If

    End Sub

    Private Sub btnClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnClose.Click

        Dim pSelectTool As ICommandItem
        Dim pCommandBars As ICommandBars
        ' The identifier for the Select Graphics Tool
        Dim pUID As New UID
        pUID.Value = "esriArcMapUI.SelectTool"

        'Find the Select Graphics Tool
        pCommandBars = m_pApp.Document.CommandBars
        pSelectTool = pCommandBars.Find(pUID)

        m_pApp.CurrentTool = pSelectTool

        Me.Hide()
        modFire_Tools.Incident_StopEditing()

    End Sub

    ' Save the form's and controls' dimensions.
    Private Sub SaveSizes() 'added 27/09/06
        Dim i As Short
        Dim ctl As System.Windows.Forms.Control

        ' Save the controls' positions and sizes.
        'UPGRADE_WARNING: Lower bound of array m_ControlPositions was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
        'UPGRADE_WARNING: Controls method Controls.count has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
        ReDim m_ControlPositions(Controls.Count())
        i = 1
        For Each ctl In Controls
            With m_ControlPositions(i)
                'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
                If TypeOf ctl Is System.Windows.Forms.Label Then
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.x1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '.Left_Renamed = ctl.x1
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.y1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '.Top_Renamed = ctl.y1
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.x1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.X2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '.Width_Renamed = ctl.X2 - ctl.x1
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.y1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    ''UPGRADE_WARNING: Couldn't resolve default property of object ctl.Y2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '.Height_Renamed = ctl.Y2 - ctl.y1
                Else
                    .Left_Renamed = VB6.PixelsToTwipsX(ctl.Left)
                    .Top_Renamed = VB6.PixelsToTwipsY(ctl.Top)
                    .Width_Renamed = VB6.PixelsToTwipsX(ctl.Width)
                    .Height_Renamed = VB6.PixelsToTwipsY(ctl.Height)
                    On Error Resume Next
                    .FontSize = ctl.Font.SizeInPoints
                    On Error GoTo 0
                End If
            End With
            i = i + 1
        Next ctl

        ' Save the form's size.
        m_FormWid = VB6.PixelsToTwipsX(ClientRectangle.Width)
        m_FormHgt = VB6.PixelsToTwipsY(ClientRectangle.Height)
    End Sub

    Private Sub cmdAddLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddLabel.Click
        Dim c_sModuleFileName As Object 'added 10/10/06 to make label editing easier
        On Error GoTo ErrorHandler
        Incident_StartEditing()

        Dim pSelectTool As ICommandItem
        Dim pCommandBars As ICommandBars
        ' The identifier for the Select Graphics Tool
        Dim pUID As New UID
        pUID.Value = "CALM_FireTools.clsFire_Labels_tool"

        'Find the Select Graphics Tool
        pCommandBars = m_pApp.Document.CommandBars
        pSelectTool = pCommandBars.Find(pUID)

        m_pApp.CurrentTool = pSelectTool

        Exit Sub
ErrorHandler:
        'UPGRADE_WARNING: Couldn't resolve default property of object c_sModuleFileName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)

    End Sub

    Private Sub cmdMoveLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMoveLabel.Click
        Dim c_sModuleFileName As Object 'added 10/10/06 to make label editing easier
        On Error GoTo ErrorHandler
        'UPGRADE_WARNING: Couldn't resolve default property of object modFire_Tools.Get_Fire_Anno_Layer(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        modFire_Tools.makeOnlySelectableLayer(modFire_Tools.Get_Fire_Anno_Layer())
        Incident_StartEditing()

        Dim pSelectTool As ICommandItem
        Dim pCommandBars As ICommandBars
        ' The identifier for the Select Graphics Tool
        Dim pUID As New UID
        pUID.Value = "esriEditor.EditTool"

        'Find the Select Graphics Tool
        pCommandBars = m_pApp.Document.CommandBars
        pSelectTool = pCommandBars.Find(pUID)

        m_pApp.CurrentTool = pSelectTool

        Exit Sub
ErrorHandler:
        'UPGRADE_WARNING: Couldn't resolve default property of object c_sModuleFileName. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)

    End Sub

    Private Sub frmFire_Labels_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'SaveSizes()
    End Sub

    'UPGRADE_WARNING: Event frmFire_Labels.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
    Private Sub frmFire_Labels_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        ' ResizeControls()
    End Sub

    ' Arrange the controls for the new size.
    Private Sub ResizeControls() 'added 27/09/06
        Dim i As Short
        Dim ctl As System.Windows.Forms.Control
        Dim xChange As Single
        Dim yChange As Single

        ' Don't bother if we are minimized.
        If WindowState = System.Windows.Forms.FormWindowState.Minimized Then Exit Sub

        'Get the change in the form size
        xChange = VB6.PixelsToTwipsX(ClientRectangle.Width) - m_FormWid
        yChange = VB6.PixelsToTwipsY(ClientRectangle.Height) - m_FormHgt

        ' Position the controls.
        i = 1
        For Each ctl In Controls
            With m_ControlPositions(i)
                'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
                If TypeOf ctl Is System.Windows.Forms.Label Then
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.x1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    ctl.x1 = xChange + .Left_Renamed
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.y1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    ctl.y1 = yChange + .Top_Renamed
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.X2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.x1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    ctl.X2 = ctl.x1 + .Width_Renamed
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Y2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    'UPGRADE_WARNING: Couldn't resolve default property of object ctl.y1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    '    ctl.Y2 = ctl.y1 + .Height_Renamed
                    '    'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
                    'ElseIf TypeOf ctl Is System.Windows.Forms.ListBox Then
                    '    ctl.Width = VB6.TwipsToPixelsX(xChange + .Width_Renamed)
                    '    ctl.Height = VB6.TwipsToPixelsY(yChange + .Height_Renamed)
                Else
                    ctl.Left = VB6.TwipsToPixelsX(xChange + .Left_Renamed)
                    ctl.Top = VB6.TwipsToPixelsY(yChange + .Top_Renamed)
                    On Error Resume Next
                End If
            End With
            i = i + 1
        Next ctl
    End Sub
End Class