'' --------------------------------------------------------------------------------
'' Code generated automatically by Code Architects' VB Migration Partner
'' --------------------------------------------------------------------------------

'Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

'<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsTrim_Unique")> _ 
'<VB6Object("CALM_FireTools.clsTrim_Unique")> _
'Public Class clsTrim_Unique
'	Implements SystemUI.ICommand
'' UPGRADE_INFO (#0501): The 'clsTrim_Unique' member isn't used anywhere in current application.

'	Private m_pApp As ESRI.ArcGIS.Framework.IApplication
'	Private m_pMXDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
'	Private WithEvents m_pDocEvent As ESRI.ArcGIS.ArcMapUI.MxDocument ' Listen for the MxDocument events
'	Private WithEvents m_pMapEvent As ESRI.ArcGIS.Carto.Map ' Listen for the Map events
'	Private WithEvents m_pLayoutEvent As ESRI.ArcGIS.Carto.PageLayout ' Listen for the PageLayout events
'	' UPGRADE_ISSUE (#0338): Unable to translate next statement. Please fix this code manually.
'    Private WithEvents m_pEditor As Editor.Editor
'	Private m_env As ESRI.ArcGIS.Geometry.IEnvelope
'	' UPGRADE_INFO (#0701): The 'm_layercount' member has been removed because it isn't used anywhere in current application.
'	' EXCLUDED: Private m_layercount As Short
'	Public m_checked As Boolean
'	Private m_mapchanged As Boolean
'	' UPGRADE_INFO (#0701): The 'justdone' member has been removed because it isn't used anywhere in current application.
'	' EXCLUDED: Private justdone As Boolean

'	' UPGRADE_INFO (#0721): The 'Class_Initialize' method is empty because all its statements were rendered as class field initializers.
'	' EXCLUDED: Private Sub Class_Initialize_VB6()
'		' EXCLUDED: 'Set m_pBitmap = LoadResPicture(101, vbResBitmap)
'	' EXCLUDED: End Sub

'	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
'		Get
'			On Error GoTo ErrorHandler
'			If TypeOf m_pApp.Document Is ESRI.ArcGIS.ArcMapUI.IMxDocument Then
'				Return True
'			Else
'				Return False
'			End If

'			Exit Property
'ErrorHandler:
'			MsgBox6("ICommand_Enabled - " & Err.Description)
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
'		Get
'			Return m_checked
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
'		Get
'			Return "Fauna"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
'		Get
'			Return "Legend Limiter"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
'		Get
'			Return "Limit Layer Legends to only show features in current extent"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
'		Get
'			Return "Limit Layer Legends to only show features in current extent"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
'		Get
'			Return ""
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
'		Get
'			Return 0
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
'		Get
'			Return GetPictureHandle6(frmImages.legend_limiter.Picture)
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
'		Get
'			Return "Developer Samples"
'	 	End Get
'	End Property

'	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
'		' The hook argument is a pointer to Application object.
'		' Establish a hook to the application
'		m_pApp = hook
'		'Start listening to the MxDocument events
'		If m_pDocEvent IsNot m_pApp.Document Then
'			m_pDocEvent = m_pApp.Document
'		End If

'		m_pMXDoc = m_pApp.Document

'		If m_pMapEvent IsNot m_pMXDoc.FocusMap Then
'			m_pMapEvent = m_pMXDoc.FocusMap
'		End If

'		If m_pLayoutEvent IsNot m_pMXDoc.PageLayout Then
'			m_pLayoutEvent = m_pMXDoc.PageLayout
'		End If

'		Dim pUID As New ESRI.ArcGIS.esriSystem.UIDClass

'		pUID.Value = "esriEditor.Editor"
'		m_pEditor = m_pApp.FindExtensionByCLSID(pUID)

'		m_checked = False

'	End Sub

'	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
'		On Error GoTo ErrorHandler
'		Dim i As Short
'		Dim f As Short
'		Dim activemap As ESRI.ArcGIS.Carto.IMap
'		If m_checked = False Then
'			modFire_Tools.StartTrimUniqueValues()
'			m_checked = True
'		Else
'			m_checked = False
'			activemap = m_pMXDoc.ActiveView.FocusMap
'			For i = 0 To m_pMXDoc.Maps.Count - 1
'				m_pMXDoc.ActiveView = m_pMXDoc.Maps.Item(i)
'				modFire_Tools.FinishTrimUniqueValues()
'			Next

'			m_pMXDoc.ActiveView = activemap

'			For f = 1 To layers_added.Count()
'				layers_added.Remove((1))
'			Next
'			m_pMXDoc.UpdateContents()
'		End If

'		Exit Sub
'ErrorHandler:
'		MsgBox6("ICommand_OnClick - " & Err.Description)
'	End Sub

'	' UPGRADE_INFO (#0701): The 'm_pDocEvent_ActiveViewChanged' member has been removed because it isn't used anywhere in current application.
'	' EXCLUDED: Private Function m_pDocEvent_ActiveViewChanged() As Boolean
'		' EXCLUDED: m_pMapEvent = m_pMXDoc.FocusMap
'	' EXCLUDED: End Function

'	Private Sub m_pEditor_OnSketchFinished()
'		m_mapchanged = True
'		m_pMXDoc.ActiveView.Refresh()
'	End Sub

'	Private Sub m_pLayoutEvent_ContentsChanged()
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pLayoutEvent_ItemAdded(ByVal Item As Object) Handles m_pLayoutEvent.ItemAdded
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pLayoutEvent_ItemDeleted(ByVal Item As Object) Handles m_pLayoutEvent.ItemDeleted
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pMapEvent_ContentsChanged()
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pMapEvent_ItemAdded(ByVal Item As Object) Handles m_pMapEvent.ItemAdded
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pMapEvent_ItemDeleted(ByVal Item As Object) Handles m_pMapEvent.ItemDeleted
'		m_mapchanged = True
'	End Sub

'	Private Sub m_pMapEvent_ViewRefreshed(ByVal view As ESRI.ArcGIS.Carto.IActiveView, ByVal phase As ESRI.ArcGIS.Carto.esriViewDrawPhase, ByVal data As Object, ByVal envelope As ESRI.ArcGIS.Geometry.IEnvelope) Handles m_pMapEvent.ViewRefreshed
'		'need to make sure the trimunique function is only executed once
'		'determine by extents and number of layers
'		If m_env Is Nothing Then
'			m_env = envelope
'		End If

'		Dim myrelation As ESRI.ArcGIS.Geometry.IRelationalOperator
'		If m_checked = True Then
'			myrelation = m_env

'			If envelope IsNot Nothing And m_env IsNot Nothing Then
'				If myrelation.Equals(envelope) = False Then
'					m_mapchanged = True
'				End If
'				If m_mapchanged = True Then
'					m_mapchanged = False
'					m_env = envelope
'					modFire_Tools.StartTrimUniqueValues()
'				End If
'			End If
'		End If
'	End Sub

'	#Region "Constructor"

'	'A public default constructor
'	Public Sub New()
'		' Add initialization code here
'	End Sub

'	#End Region

'	#Region "Static constructor"

'	' This static constructor ensures that the VB6 support library be initialized before using current class.
'	Shared Sub New()
'		EnsureVB6LibraryInitialization()
'	End Sub

'	#End Region

'End Class
