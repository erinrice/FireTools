<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Effective_From
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
	Public WithEvents btnOK As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents btnCancel As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents txtDate As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents ckbAppend As CodeArchitects.VB6Library.VB6CheckBox
	Public WithEvents Line1 As CodeArchitects.VB6Library.VB6Line

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.btnOK = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.btnCancel = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.txtDate = New CodeArchitects.VB6Library.VB6TextBox()
		Me.ckbAppend = New CodeArchitects.VB6Library.VB6CheckBox()
		Me.Line1 = New CodeArchitects.VB6Library.VB6Line()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		' btnOK
		'
		Me.btnOK.Name = "btnOK"
		Me.btnOK.Size = New System.Drawing.Size(97, 25)
		Me.btnOK.Location = New System.Drawing.Point(24, 96)
		Me.AcceptButton = Me.btnOK
		Me.btnOK.Caption = "OK"
		Me.btnOK.TabIndex = 3
		Me.btnOK.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnOK.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' btnCancel
		'
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(97, 25)
		Me.btnCancel.Location = New System.Drawing.Point(152, 96)
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Caption = "CANCEL"
		Me.btnCancel.TabIndex = 2
		Me.btnCancel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCancel.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' txtDate
		'
		Me.txtDate.Name = "txtDate"
		Me.txtDate.Size = New System.Drawing.Size(241, 19)
		Me.txtDate.Location = New System.Drawing.Point(16, 64)
		Me.txtDate.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtDate.ReadOnly = True
		Me.txtDate.TabIndex = 1
		Me.txtDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' ckbAppend
		'
		Me.ckbAppend.Name = "ckbAppend"
		Me.ckbAppend.Size = New System.Drawing.Size(265, 17)
		Me.ckbAppend.Location = New System.Drawing.Point(8, 16)
		Me.ckbAppend.Caption = "Add new shapes to existing features"
		Me.ckbAppend.TabIndex = 0
		Me.ckbAppend.Value = CodeArchitects.VB6Library.VBRUN.CheckBoxConstants.vbChecked
		Me.ckbAppend.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ckbAppend.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Line1
		'
		Me.Line1.Name = "Line1"
		Me.Line1.ParentForm = Me
		Me.Line1.Name6 = "Line1"
		Me.Line1.X1 = 240
		Me.Line1.X2 = 3840
		Me.Line1.Y1 = 720
		Me.Line1.Y2 = 720
		'
		' frmFire_Effective_From
		'
		Me.Name = "frmFire_Effective_From"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Effective From"
		Me.ControlBox = False
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(3, 29)
		Me.ClientSize = New System.Drawing.Size(273, 132)

		Me.Controls.Add(btnOK)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(txtDate)
		Me.Controls.Add(ckbAppend)
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
