<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRegion_Choice
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
	Public WithEvents Cancel As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents OK As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents RegionList As CodeArchitects.VB6Library.VB6ListBox

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRegion_Choice))
		Me.Cancel = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.OK = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.RegionList = New CodeArchitects.VB6Library.VB6ListBox()
		Me.SuspendLayout()
		'
		' Cancel
		'
		Me.Cancel.Name = "Cancel"
		Me.Cancel.Size = New System.Drawing.Size(73, 25)
		Me.Cancel.Location = New System.Drawing.Point(88, 160)
		Me.CancelButton = Me.Cancel
		Me.Cancel.Caption = "CANCEL"
		Me.Cancel.TabIndex = 2
		Me.Cancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Cancel.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' OK
		'
		Me.OK.Name = "OK"
		Me.OK.Size = New System.Drawing.Size(73, 25)
		Me.OK.Location = New System.Drawing.Point(8, 160)
		Me.AcceptButton = Me.OK
		Me.OK.Caption = "OK"
		Me.OK.TabIndex = 1
		Me.OK.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OK.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' RegionList
		'
		Me.RegionList.Name = "RegionList"
		Me.RegionList.Size = New System.Drawing.Size(153, 148)
		Me.RegionList.Location = New System.Drawing.Point(8, 8)
		Me.RegionList.TabIndex = 0
		Me.RegionList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.RegionList.ItemDataValues = "Goldfields\r0\rKimberley\r0\rMidwest\r0\rPilbara\r0\rSouth\ Coast\r0\rSouth\ West\r0\rSwan\r0\rWarren\r0\rWheatbelt\r0"
		'
		' frmRegion_Choice
		'
		Me.Name = "frmRegion_Choice"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Select Region"
		Me.ControlBox = False
		Me.Icon = Nothing
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(3, 29)
		Me.ClientSize = New System.Drawing.Size(171, 192)

		Me.Controls.Add(Cancel)
		Me.Controls.Add(OK)
		Me.Controls.Add(RegionList)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
