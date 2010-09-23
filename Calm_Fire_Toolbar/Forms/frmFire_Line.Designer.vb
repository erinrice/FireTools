<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Line
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
	Public WithEvents cbCapt_Method As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents txtTime As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtDate As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtDept As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtUserName As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtFireNumber As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtRegion_Ex As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtRegion As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents Feature_type_cb As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents OK_btn As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Cancel_btn As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Label5 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line4 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label6 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line3 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Line2 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Line1 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label7 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label4 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label3 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label2 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label1 As CodeArchitects.VB6Library.VB6Label

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFire_Line))
		Me.cbCapt_Method = New CodeArchitects.VB6Library.VB6ComboBox()
		Me.txtTime = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtDate = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtDept = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtUserName = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtFireNumber = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtRegion_Ex = New CodeArchitects.VB6Library.VB6TextBox()
		Me.txtRegion = New CodeArchitects.VB6Library.VB6TextBox()
		Me.Feature_type_cb = New CodeArchitects.VB6Library.VB6ComboBox()
		Me.OK_btn = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.Cancel_btn = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.Label5 = New CodeArchitects.VB6Library.VB6Label()
		Me.Line4 = New CodeArchitects.VB6Library.VB6Line()
		Me.Label6 = New CodeArchitects.VB6Library.VB6Label()
		Me.Line3 = New CodeArchitects.VB6Library.VB6Line()
		Me.Line2 = New CodeArchitects.VB6Library.VB6Line()
		Me.Line1 = New CodeArchitects.VB6Library.VB6Line()
		Me.Label7 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label4 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label3 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label2 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label1 = New CodeArchitects.VB6Library.VB6Label()
		CType(Me.Line4, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		' cbCapt_Method
		'
		Me.cbCapt_Method.Name = "cbCapt_Method"
		Me.cbCapt_Method.Size = New System.Drawing.Size(169, 21)
		Me.cbCapt_Method.Location = New System.Drawing.Point(152, 176)
		Me.cbCapt_Method.TabIndex = 15
		Me.cbCapt_Method.Text = "Unknown"
		Me.cbCapt_Method.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCapt_Method.ItemDataValues = "Unknown\r0\rGPS\ -\ Differential\ \(\+/-2-5m\)\r0\rGPS\ -\ Non\ Differential\ \(\+/-5-10m\)\r0\rReference\ Map\ /\ Photo\ \(50k\)\ \(\+/-250m\)\r0\rMudmap\ \(\+/-500m\)\r0\rAnecdotal\ \(\+/-1km\)\r0"
		'
		' txtTime
		'
		Me.txtTime.Name = "txtTime"
		Me.txtTime.Size = New System.Drawing.Size(73, 19)
		Me.txtTime.Location = New System.Drawing.Point(248, 120)
		Me.txtTime.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtTime.BackColor = FromOleColor6(CInt(&H80000004))
		Me.txtTime.TabIndex = 9
		Me.txtTime.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtDate
		'
		Me.txtDate.Name = "txtDate"
		Me.txtDate.Size = New System.Drawing.Size(81, 19)
		Me.txtDate.Location = New System.Drawing.Point(152, 120)
		Me.txtDate.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtDate.BackColor = FromOleColor6(CInt(&H80000004))
		Me.txtDate.TabIndex = 8
		Me.txtDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtDept
		'
		Me.txtDept.Name = "txtDept"
		Me.txtDept.Size = New System.Drawing.Size(169, 19)
		Me.txtDept.Location = New System.Drawing.Point(152, 88)
		Me.txtDept.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtDept.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtDept.ReadOnly = True
		Me.txtDept.TabIndex = 7
		Me.txtDept.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtUserName
		'
		Me.txtUserName.Name = "txtUserName"
		Me.txtUserName.Size = New System.Drawing.Size(169, 19)
		Me.txtUserName.Location = New System.Drawing.Point(152, 64)
		Me.txtUserName.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtUserName.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtUserName.ReadOnly = True
		Me.txtUserName.TabIndex = 6
		Me.txtUserName.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtFireNumber
		'
		Me.txtFireNumber.Name = "txtFireNumber"
		Me.txtFireNumber.Size = New System.Drawing.Size(57, 19)
		Me.txtFireNumber.Location = New System.Drawing.Point(200, 8)
		Me.txtFireNumber.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtFireNumber.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtFireNumber.ReadOnly = True
		Me.txtFireNumber.TabIndex = 5
		Me.txtFireNumber.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtRegion_Ex
		'
		Me.txtRegion_Ex.Name = "txtRegion_Ex"
		Me.txtRegion_Ex.Size = New System.Drawing.Size(41, 19)
		Me.txtRegion_Ex.Location = New System.Drawing.Point(152, 8)
		Me.txtRegion_Ex.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtRegion_Ex.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtRegion_Ex.ReadOnly = True
		Me.txtRegion_Ex.TabIndex = 4
		Me.txtRegion_Ex.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' txtRegion
		'
		Me.txtRegion.Name = "txtRegion"
		Me.txtRegion.Size = New System.Drawing.Size(137, 19)
		Me.txtRegion.Location = New System.Drawing.Point(152, 32)
		Me.txtRegion.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
		Me.txtRegion.BackColor = FromOleColor6(CInt(&H80000000))
		Me.txtRegion.ReadOnly = True
		Me.txtRegion.TabIndex = 3
		Me.txtRegion.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' Feature_type_cb
		'
		Me.Feature_type_cb.Name = "Feature_type_cb"
		Me.Feature_type_cb.Size = New System.Drawing.Size(169, 21)
		Me.Feature_type_cb.Location = New System.Drawing.Point(152, 152)
		Me.Feature_type_cb.TabIndex = 2
		Me.Feature_type_cb.Text = "Estimate Fire Edge"
		Me.Feature_type_cb.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Feature_type_cb.ItemDataValues = "Estimate\ Fire\ Edge\r0\rPredicted\ Fire\ Edge\r0\rFire\ Stopped\r0\rFire\ Line\ Constructed\r0\rPatrol\ Only\r0\rControl\ Line\ Proposed\r0\rControl\ Line\ Constructed\r0\rBackburn\ Line\r0rRoad\ Closed\r0"
		'
		' OK_btn
		'
		Me.OK_btn.Name = "OK_btn"
		Me.OK_btn.Size = New System.Drawing.Size(105, 25)
		Me.OK_btn.Location = New System.Drawing.Point(40, 216)
		Me.AcceptButton = Me.OK_btn
		Me.OK_btn.Caption = "OK"
		Me.OK_btn.TabIndex = 1
		Me.OK_btn.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OK_btn.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Cancel_btn
		'
		Me.Cancel_btn.Name = "Cancel_btn"
		Me.Cancel_btn.Size = New System.Drawing.Size(105, 25)
		Me.Cancel_btn.Location = New System.Drawing.Point(176, 216)
		Me.CancelButton = Me.Cancel_btn
		Me.Cancel_btn.Caption = "CANCEL"
		Me.Cancel_btn.TabIndex = 0
		Me.Cancel_btn.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Cancel_btn.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label5
		'
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(129, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 152)
		Me.Label5.Caption = "Fire Line Feature"
		Me.Label5.TabIndex = 17
		Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.AutoSize = False
		Me.Label5.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Line4
		'
		Me.Line4.Name = "Line4"
		Me.Line4.ParentForm = Me
		Me.Line4.Name6 = "Line4"
		Me.Line4.X1 = 120
		Me.Line4.X2 = 4800
		Me.Line4.Y1 = 3000
		Me.Line4.Y2 = 3000
		'
		' Label6
		'
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(97, 17)
		Me.Label6.Location = New System.Drawing.Point(8, 176)
		Me.Label6.Caption = "Capture Method"
		Me.Label6.TabIndex = 16
		Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.AutoSize = False
		Me.Label6.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Line3
		'
		Me.Line3.Name = "Line3"
		Me.Line3.ParentForm = Me
		Me.Line3.Name6 = "Line3"
		Me.Line3.X1 = 120
		Me.Line3.X2 = 4800
		Me.Line3.Y1 = 2160
		Me.Line3.Y2 = 2160
		'
		' Line2
		'
		Me.Line2.Name = "Line2"
		Me.Line2.ParentForm = Me
		Me.Line2.Name6 = "Line2"
		Me.Line2.X1 = 120
		Me.Line2.X2 = 4800
		Me.Line2.Y1 = 1680
		Me.Line2.Y2 = 1680
		'
		' Line1
		'
		Me.Line1.Name = "Line1"
		Me.Line1.ParentForm = Me
		Me.Line1.Name6 = "Line1"
		Me.Line1.X1 = 120
		Me.Line1.X2 = 4800
		Me.Line1.Y1 = 840
		Me.Line1.Y2 = 840
		'
		' Label7
		'
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(89, 17)
		Me.Label7.Location = New System.Drawing.Point(8, 120)
		Me.Label7.Caption = "Date / Time"
		Me.Label7.TabIndex = 14
		Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.AutoSize = False
		Me.Label7.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label4
		'
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(121, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 88)
		Me.Label4.Caption = "User's Department"
		Me.Label4.TabIndex = 13
		Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.AutoSize = False
		Me.Label4.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label3
		'
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(89, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 64)
		Me.Label3.Caption = "User's Name"
		Me.Label3.TabIndex = 12
		Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.AutoSize = False
		Me.Label3.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label2
		'
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.Caption = "Fire ID"
		Me.Label2.TabIndex = 11
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.AutoSize = False
		Me.Label2.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label1
		'
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(73, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 32)
		Me.Label1.Caption = "CALM Region"
		Me.Label1.TabIndex = 10
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.AutoSize = False
		Me.Label1.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' frmFire_Line
		'
		Me.Name = "frmFire_Line"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Fire Line"
		Me.ControlBox = False
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(3, 29)
		Me.ClientSize = New System.Drawing.Size(329, 250)

		Me.Controls.Add(cbCapt_Method)
		Me.Controls.Add(txtTime)
		Me.Controls.Add(txtDate)
		Me.Controls.Add(txtDept)
		Me.Controls.Add(txtUserName)
		Me.Controls.Add(txtFireNumber)
		Me.Controls.Add(txtRegion_Ex)
		Me.Controls.Add(txtRegion)
		Me.Controls.Add(Feature_type_cb)
		Me.Controls.Add(OK_btn)
		Me.Controls.Add(Cancel_btn)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		CType(Me.Line4, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
