<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Create_Incident
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
	Public WithEvents cbWorkCentre As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents btnCancel As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents btnOK As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cbInc_Region_ext As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents btnNewGDB As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents txtNewGDB As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtInc_Number As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents cbInc_Region As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents Label4 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line3 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Line1 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label3 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label2 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label1 As CodeArchitects.VB6Library.VB6Label

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFire_Create_Incident))
		Me.cbWorkCentre = New CodeArchitects.VB6Library.VB6ComboBox()
		Me.btnCancel = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.btnOK = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cbInc_Region_ext = New CodeArchitects.VB6Library.VB6ComboBox()
		Me.btnNewGDB = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.txtNewGDB = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtInc_Number = New CodeArchitects.VB6Library.VB6TextBox()
		Me.cbInc_Region = New CodeArchitects.VB6Library.VB6ComboBox()
		Me.Label4 = New CodeArchitects.VB6Library.VB6Label()
		Me.Line3 = New CodeArchitects.VB6Library.VB6Line()
		Me.Line1 = New CodeArchitects.VB6Library.VB6Line()
		Me.Label3 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label2 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label1 = New CodeArchitects.VB6Library.VB6Label()
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		' cbWorkCentre
		'
		Me.cbWorkCentre.Name = "cbWorkCentre"
		Me.cbWorkCentre.Size = New System.Drawing.Size(81, 21)
		Me.cbWorkCentre.Location = New System.Drawing.Point(120, 48)
		Me.cbWorkCentre.Style = CodeArchitects.VB6Library.VBRUN.ComboBoxConstants.vbComboDropdownList
		Me.cbWorkCentre.TabIndex = 10
		Me.cbWorkCentre.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbWorkCentre.ItemDataValues = "Kalgoorlie\r0"
		'
		' btnCancel
		'
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(89, 25)
		Me.btnCancel.Location = New System.Drawing.Point(376, 48)
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Caption = "CANCEL"
		Me.btnCancel.TabIndex = 9
		Me.btnCancel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCancel.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' btnOK
		'
		Me.btnOK.Name = "btnOK"
		Me.btnOK.Size = New System.Drawing.Size(89, 25)
		Me.btnOK.Location = New System.Drawing.Point(376, 16)
		Me.AcceptButton = Me.btnOK
		Me.btnOK.Caption = "OK"
		Me.btnOK.TabIndex = 8
		Me.btnOK.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnOK.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cbInc_Region_ext
		'
		Me.cbInc_Region_ext.Name = "cbInc_Region_ext"
		Me.cbInc_Region_ext.Size = New System.Drawing.Size(81, 21)
		Me.cbInc_Region_ext.Location = New System.Drawing.Point(264, 16)
		Me.cbInc_Region_ext.Style = CodeArchitects.VB6Library.VBRUN.ComboBoxConstants.vbComboDropdownList
		Me.cbInc_Region_ext.TabIndex = 4
		Me.cbInc_Region_ext.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbInc_Region_ext.ItemDataValues = "GFD\r0"
		'
		' btnNewGDB
		'
		Me.btnNewGDB.Name = "btnNewGDB"
		Me.btnNewGDB.Size = New System.Drawing.Size(25, 25)
		Me.btnNewGDB.Location = New System.Drawing.Point(320, 144)
		Me.btnNewGDB.BackColor = FromOleColor6(CInt(&H80000009))
		Me.btnNewGDB.Picture = CType(resources.GetObject("btnNewGDB.Picture"), System.Drawing.Image)
		Me.btnNewGDB.Style = CodeArchitects.VB6Library.VBRUN.ButtonConstants.vbButtonGraphical
		Me.btnNewGDB.TabIndex = 3
		Me.btnNewGDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtNewGDB
		'
		Me.txtNewGDB.Name = "txtNewGDB"
		Me.txtNewGDB.Size = New System.Drawing.Size(297, 21)
		Me.txtNewGDB.Location = New System.Drawing.Point(16, 144)
		Me.txtNewGDB.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtNewGDB.ReadOnly = True
		Me.txtNewGDB.TabIndex = 2
		Me.txtNewGDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtInc_Number
		'
		Me.txtInc_Number.Name = "txtInc_Number"
		Me.txtInc_Number.Size = New System.Drawing.Size(81, 19)
		Me.txtInc_Number.Location = New System.Drawing.Point(120, 80)
		Me.txtInc_Number.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtInc_Number.MaxLength = 4
		Me.txtInc_Number.TabIndex = 1
		Me.txtInc_Number.DataFormatValues = "1°@°0°@°True°@°False°@°°@°0°@°0"
		Me.txtInc_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' cbInc_Region
		'
		Me.cbInc_Region.Name = "cbInc_Region"
		Me.cbInc_Region.Size = New System.Drawing.Size(129, 21)
		Me.cbInc_Region.Location = New System.Drawing.Point(120, 16)
		Me.cbInc_Region.Sorted = True
		Me.cbInc_Region.Style = CodeArchitects.VB6Library.VBRUN.ComboBoxConstants.vbComboDropdownList
		Me.cbInc_Region.TabIndex = 0
		Me.cbInc_Region.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbInc_Region.ItemDataValues = "Goldfields\r0\rKimberley\r0\rMidwest\r0\rPilbara\r0\rSth\ Coast\r0\rSth\ West\r0\rSwan\r0\rWarren\r0\rWheatbelt\r0"
		'
		' Label4
		'
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(137, 17)
		Me.Label4.Location = New System.Drawing.Point(16, 48)
		Me.Label4.Caption = "Work Centre"
		Me.Label4.TabIndex = 11
		Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.AutoSize = False
		Me.Label4.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Line3
		'
		Me.Line3.Name = "Line3"
		Me.Line3.ParentForm = Me
		Me.Line3.Name6 = "Line3"
		Me.Line3.X1 = 5400
		Me.Line3.X2 = 5400
		Me.Line3.Y1 = 240
		Me.Line3.Y2 = 2520
		'
		' Line1
		'
		Me.Line1.Name = "Line1"
		Me.Line1.ParentForm = Me
		Me.Line1.Name6 = "Line1"
		Me.Line1.X1 = 240
		Me.Line1.X2 = 5160
		Me.Line1.Y1 = 1680
		Me.Line1.Y2 = 1680
		'
		' Label3
		'
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(209, 17)
		Me.Label3.Location = New System.Drawing.Point(16, 120)
		Me.Label3.Caption = "New Fire Incident Geodatabase"
		Me.Label3.TabIndex = 7
		Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.AutoSize = False
		Me.Label3.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label2
		'
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(137, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 80)
		Me.Label2.Caption = "Fire Number"
		Me.Label2.TabIndex = 6
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.AutoSize = False
		Me.Label2.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label1
		'
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(129, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 16)
		Me.Label1.Caption = "Fire Region"
		Me.Label1.TabIndex = 5
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.AutoSize = False
		Me.Label1.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' frmFire_Create_Incident
		'
		Me.Name = "frmFire_Create_Incident"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Create Incident"
		Me.ControlBox = False
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(3, 29)
		Me.ClientSize = New System.Drawing.Size(476, 182)

		Me.Controls.Add(cbWorkCentre)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.Controls.Add(cbInc_Region_ext)
		Me.Controls.Add(btnNewGDB)
		Me.Controls.Add(txtNewGDB)
		Me.Controls.Add(txtInc_Number)
		Me.Controls.Add(cbInc_Region)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
