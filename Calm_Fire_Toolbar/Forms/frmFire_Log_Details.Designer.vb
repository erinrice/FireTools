﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Log_Details
	Inherits CodeArchitects.VB6Library.VB6Form


#Region "Static constructor"

' This static constructor ensures that the VB6 support library be initialized before using current class.
Shared Sub New()
	EnsureVB6LibraryInitialization()
End Sub

#End Region

#Region "Windows Form Designer generated code "

	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'Create all controls and control arrays.
		InitializeComponents()
	End Sub

	' This method wraps the call to InitializeComponent, but can be called from base class.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overrides Sub InitializeComponents()
		Me.ObjectIsInitializing = True
		InitializeComponent()
		Me.ObjectIsInitializing = False
	End Sub

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing Then
				If components IsNot Nothing Then components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public WithEvents TreeView1 As CodeArchitects.VB6Library.VB6TreeView

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.TreeView1 = New CodeArchitects.VB6Library.VB6TreeView()
		Me.SuspendLayout()
		'
		' TreeView1
		'
		Me.TreeView1.Name = "TreeView1"
		Me.TreeView1.Size = New System.Drawing.Size(297, 193)
		Me.TreeView1.Location = New System.Drawing.Point(8, 8)
		Me.TreeView1.TabIndex = 0
		Me.TreeView1.LineStyle = CodeArchitects.VB6Library.MSComctlLib.TreeLineStyleConstants.tvwRootLines
		Me.TreeView1.Style = CodeArchitects.VB6Library.MSComctlLib.TreeStyleConstants.tvwTreelinesPlusMinusText
		Me.TreeView1.Indentation = 37.7929
		Me.TreeView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TreeView1.LabelEdit = CodeArchitects.VB6Library.MSComctlLib.LabelEditConstants.tvwAutomatic
		'
		' frmFire_Log_Details
		'
		Me.Name = "frmFire_Log_Details"
		Me.Caption = "Fire Log Details"
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(4, 30)
		Me.ClientSize = New System.Drawing.Size(312, 206)

		Me.Controls.Add(TreeView1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
