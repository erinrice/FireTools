'' --------------------------------------------------------------------------------
'' Code generated automatically by Code Architects' VB Migration Partner
'' --------------------------------------------------------------------------------

'Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

'<System.Runtime.InteropServices.ProgID("CALM_FireTools.clsShow_FireLog")> _ 
'<VB6Object("CALM_FireTools.clsShow_FireLog")> _
'Public Class clsShow_FireLog
'	Implements SystemUI.ICommand
'' UPGRADE_INFO (#0501): The 'clsShow_FireLog' member isn't used anywhere in current application.

'	' Copyright 1995-2005 ESRI

'	' All rights reserved under the copyright laws of the United States.

'	' You may freely redistribute and use this sample code, with or without modification.

'	' Disclaimer: THE SAMPLE CODE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED
'	' WARRANTIES, INCLUDING THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
'	' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ESRI OR
'	' CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
'	' OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
'	' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
'	' INTERRUPTION) SUSTAINED BY YOU OR A THIRD PARTY, HOWEVER CAUSED AND ON ANY
'	' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT ARISING IN ANY
'	' WAY OUT OF THE USE OF THIS SAMPLE CODE, EVEN IF ADVISED OF THE POSSIBILITY OF
'	' SUCH DAMAGE.

'	' For additional information contact: Environmental Systems Research Institute, Inc.

'	' Attn: Contracts Dept.

'	' 380 New York Street

'	' Redlands, California, U.S.A. 92373

'	' Email: contracts@esri.com

'	' Command that toggles the visibility of the dockable window

'	Private m_pApp As ESRI.ArcGIS.Framework.IApplication
'	Private m_pDockWinMgr As ESRI.ArcGIS.Framework.IDockableWindowManager
'	Private m_pDockWin As ESRI.ArcGIS.Framework.IDockableWindow

'	Private ReadOnly Property ICommand_Bitmap() As Integer Implements SystemUI.ICommand.Bitmap
'		Get
'			Return GetPictureHandle6(frmImages.log.Picture)
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Caption() As String Implements SystemUI.ICommand.Caption
'		Get
'			Return "Fire Log"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Category() As String Implements SystemUI.ICommand.Category
'		Get
'			Return "Developer Samples"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Checked() As Boolean Implements SystemUI.ICommand.Checked
'		Get
'			' If the dockable window is visible, make the button of this command look
'			' pushed in
'			Return m_pDockWin.IsVisible()
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Enabled() As Boolean Implements SystemUI.ICommand.Enabled
'		Get
'			Return True
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements SystemUI.ICommand.HelpContextID
'		Get
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_HelpFile() As String Implements SystemUI.ICommand.HelpFile
'		Get
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Message() As String Implements SystemUI.ICommand.Message
'		Get
'			Return "Display the selection count dockable window"
'	 	End Get
'	End Property

'	Private ReadOnly Property ICommand_Name() As String Implements SystemUI.ICommand.Name
'		Get
'			Return "DeveloperSamples_SelCountDockableWindow"
'	 	End Get
'	End Property

'	Private Sub ICommand_OnClick() Implements SystemUI.ICommand.OnClick
'		' Toggle the visibility of the dockable window
'		On Error GoTo errhandler

'		m_pDockWin.Show(Not m_pDockWin.IsVisible())

'		Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)
'		Dim myIncidentInfo As clsFire_Incident_Info

'		If TypeOf pExt Is clsFire_Incident_Info Then
'			myIncidentInfo = pExt
'		End If

'		If myIncidentInfo.IncidentWorkspace IsNot Nothing Then
'			modFire_Tools.Update_Log_Window()
'		End If

'		Exit Sub
'errhandler:
'		MsgBox6(Err.Description)
'	End Sub

'	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements SystemUI.ICommand.OnCreate
'		m_pApp = hook
'		m_pDockWinMgr = hook 'QI for IDockableWindowManager

'		' Get a reference to the selection count dockable window
'		Dim u As New ESRI.ArcGIS.esriSystem.UIDClass
'		u.Value = "CALM_FireTools.clsFire_Log_Dockwin"
'		m_pDockWin = m_pDockWinMgr.GetDockableWindow(u)
'	End Sub

'	Private ReadOnly Property ICommand_Tooltip() As String Implements SystemUI.ICommand.Tooltip
'		Get
'			Return "Selection Count"
'	 	End Get
'	End Property

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
