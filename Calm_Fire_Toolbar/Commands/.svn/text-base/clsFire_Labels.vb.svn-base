' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsFire_Labels")> _ 
<VB6Object("CALM_FireTools.clsFire_Labels")> _
Public Class clsFire_Labels
	Implements SystemUI.ICommand
' UPGRADE_INFO (#0501): The 'clsFire_Labels' member isn't used anywhere in current application.

	'This is called from the Fire Toolbar > Tools Menu
	
	' Variables used by the Error handler function - DO NOT REMOVE
    Public papplication As ESRI.ArcGIS.Framework.IApplication
	Public m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
	Private WithEvents m_wpDoc As ESRI.ArcGIS.ArcMapUI.MxDocument
	Private Const c_sModuleFileName As String = ""
	' Constant reflect file module name
	' Publicly declared API function and constant.
	' UPGRADE_INFO (#0531): You can replace calls to the 'SetWindowLong' unmanaged method with the following .NET member(s): System.Windows.Forms.Control.Parent, System.Windows.Forms.Form.CreateParams.ExStyle, System.Windows.Forms.Form.CreateParams.Style, System.Windows.Forms.Form, System.Windows.Forms.Form.
	Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	' UPGRADE_INFO (#0561): The 'GWL_HWNDPARENT' symbol was defined without an explicit "As" clause.
	Private Const GWL_HWNDPARENT As Short = (-8)

    'saeid: added this to save the current instance of the form
    Private Shared _loadedForm As System.Windows.Forms.Form

    Public Shared Property loadedForm() As System.Windows.Forms.Form
        Get
            Return _loadedForm
        End Get
        Set(ByVal value As System.Windows.Forms.Form)
            _loadedForm = value
        End Set
    End Property


    Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
        Get
            On Error GoTo ErrorHandler

            Dim bEnabled As Boolean = False

            bEnabled = ExtensionEnabled(ByVal6(modCommon_Functions.m_pApp), g_FIRE_EXTENSION)
            If modFire_Tools.Fire_Anno_Exists() = True Then
                bEnabled = True
            Else
                bEnabled = False
            End If
            If g_Username = "" Then
                bEnabled = False
            End If

            Return bEnabled

ErrorHandler:
            HandleError(True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
            Return True
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            ' ICommand_Checked =

            Exit Property
ErrorHandler:
            HandleError(True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return "FireLabels"

ErrorHandler:
            HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return "Fire Labels"

ErrorHandler:
            HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return "Create Standardised Fire Labels"

ErrorHandler:
            HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return "Create Standardised Fire Labels"

ErrorHandler:
            HandleError(True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            ' ICommand_HelpFile =

            Exit Property
ErrorHandler:
            HandleError(True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            ' ICommand_HelpContextID =

            Exit Property
ErrorHandler:
            HandleError(True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            Return GetPictureHandle6(frmImages.label.Picture)

ErrorHandler:
            HandleError(True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
        Get
            On Error GoTo ErrorHandler

            ' TODO: Add your implementation here
            ' ICommand_Category =

            Exit Property
ErrorHandler:
            HandleError(True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
        End Get
    End Property

    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
        On Error GoTo ErrorHandler
        papplication = hook
        m_pApp = hook
        m_pDoc = papplication.Document
        m_wpDoc = papplication.Document
        ' TODO: Add your implementation here
        Exit Sub
ErrorHandler:
        HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
    End Sub

    Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick

        Try
            modFire_Tools.Incident_StopEditing() 'added 23/10/06

            Dim pSelectTool As ESRI.ArcGIS.Framework.ICommandItem
            Dim pCommandBars As ESRI.ArcGIS.Framework.ICommandBars = Nothing
            ' The identifier for the Select Graphics Tool
            Dim pUID As New ESRI.ArcGIS.esriSystem.UIDClass
            pUID.Value = "CALM_FireTools.clsFire_Labels_Tool"

            'Find the Select Graphics Tool
            pCommandBars = m_pApp.Document.CommandBars
            'GrantB 19-10-2007 - need to know if the toolBar was set
            If pCommandBars Is Nothing Then
                MsgBox6("The FireToolBar could not be set! This tool might not function correctly.")
            End If

            'GrantB 19-10-2007 - need to know if the tool was set
            pSelectTool = Nothing
            pSelectTool = pCommandBars.Find(pUID)
            If pSelectTool Is Nothing Then
                MsgBox6("The Tool could not be set! This tool might not function correctly.")
                Exit Sub
            End If

            m_pApp.CurrentTool = pSelectTool
            '    Set frmFire_Labels.pST = pSelectTool
            '    Set frmFire_Labels.pAP = m_lApp
            Dim myForm As New frmFire_Labels
            ' Saeid: added this part to make it possible to access the form from clsFire_labels_tool
            loadedForm = myForm
            myForm.Show()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try


        'frmFire_Labels.Show(VBRUN.FormShowConstants.vbModeless) ', m_pApp
        'SetWindowLong(frmFire_Labels.hWnd, GWL_HWNDPARENT, m_pApp.hWnd)
        'SetWindowLong(frmFire_Labels.hWnd, GWL_HWNDPARENT, m_pApp.hWnd)

        'ErrorHandler:
        '        HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
    End Sub

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
