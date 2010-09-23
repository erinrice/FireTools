<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDriveChoice
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
	Public WithEvents CANCEL As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents OK As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Drivelist As CodeArchitects.VB6Library.VB6ListBox

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDriveChoice))
		Me.CANCEL = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.OK = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.Drivelist = New CodeArchitects.VB6Library.VB6ListBox()
		Me.SuspendLayout()
		'
		' CANCEL
		'
		Me.CANCEL.Name = "CANCEL"
		Me.CANCEL.Size = New System.Drawing.Size(73, 25)
		Me.CANCEL.Location = New System.Drawing.Point(88, 144)
		Me.CancelButton = Me.CANCEL
		Me.CANCEL.Caption = "CANCEL"
		Me.CANCEL.TabIndex = 2
		Me.CANCEL.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CANCEL.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' OK
		'
		Me.OK.Name = "OK"
		Me.OK.Size = New System.Drawing.Size(73, 25)
		Me.OK.Location = New System.Drawing.Point(8, 144)
		Me.AcceptButton = Me.OK
		Me.OK.Caption = "OK"
		Me.OK.TabIndex = 1
		Me.OK.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OK.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Drivelist
		'
		Me.Drivelist.Name = "Drivelist"
		Me.Drivelist.Size = New System.Drawing.Size(153, 119)
		Me.Drivelist.Location = New System.Drawing.Point(8, 8)
		Me.Drivelist.TabIndex = 0
		Me.Drivelist.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Drivelist.ItemDataValues = "C\r0\rD\r0\rE\r0\rF\r0\rG\r0\rH\r0\rI\r0\rJ\r0\rK\r0\rL\r0\rM\r0\rN\r0\rO\r0\rP\r0\rQ\r0\rR\r0\rS\r0\rT\r0\rU\r0\rV\r0\rW\r0\rX\r0\rY\r0\rZ\r0"
		'
		' frmDriveChoice
		'
		Me.Name = "frmDriveChoice"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Corporate Data Drive"
		Me.ControlBox = False
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(3, 29)
		Me.ClientSize = New System.Drawing.Size(169, 175)

		Me.Controls.Add(CANCEL)
		Me.Controls.Add(OK)
		Me.Controls.Add(Drivelist)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
