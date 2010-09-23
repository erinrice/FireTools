<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Export_Map
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
		' Initialize control arrays.
		Me.Label10 = New CodeArchitects.VB6Library.VB6ControlArray(Of CodeArchitects.VB6Library.VB6Label)(Label10_000)
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

	Public Label10 As CodeArchitects.VB6Library.VB6ControlArray(Of CodeArchitects.VB6Library.VB6Label)
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public WithEvents cbQuality As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents btnNewFile As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents txtfilename As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtfolder As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtMapScale As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtUsername As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtOrientation As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtPage_Size As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents cbMap_Type As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents cbResolution As CodeArchitects.VB6Library.VB6ComboBox
	Public WithEvents OK_btn As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Cancel_btn As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents txtRegion_Ex As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtFireNumber As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtDept As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtDate As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents txtTime As CodeArchitects.VB6Library.VB6TextBox
	Public WithEvents Label13 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line6 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label12 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label11 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line4 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label10_000 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label9 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label6 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label1 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line5 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Label2 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label3 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label4 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label7 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Line1 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Line2 As CodeArchitects.VB6Library.VB6Line
	Public WithEvents Line3 As CodeArchitects.VB6Library.VB6Line

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFire_Export_Map))
        Me.cbQuality = New CodeArchitects.VB6Library.VB6ComboBox
        Me.btnNewFile = New CodeArchitects.VB6Library.VB6CommandButton
        Me.txtfilename = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtfolder = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtMapScale = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtUsername = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtOrientation = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtPage_Size = New CodeArchitects.VB6Library.VB6TextBox
        Me.cbMap_Type = New CodeArchitects.VB6Library.VB6ComboBox
        Me.cbResolution = New CodeArchitects.VB6Library.VB6ComboBox
        Me.OK_btn = New CodeArchitects.VB6Library.VB6CommandButton
        Me.Cancel_btn = New CodeArchitects.VB6Library.VB6CommandButton
        Me.txtRegion_Ex = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtFireNumber = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtDept = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtDate = New CodeArchitects.VB6Library.VB6TextBox
        Me.txtTime = New CodeArchitects.VB6Library.VB6TextBox
        Me.Label13 = New CodeArchitects.VB6Library.VB6Label
        Me.Line6 = New CodeArchitects.VB6Library.VB6Line
        Me.Label12 = New CodeArchitects.VB6Library.VB6Label
        Me.Label11 = New CodeArchitects.VB6Library.VB6Label
        Me.Line4 = New CodeArchitects.VB6Library.VB6Line
        Me.Label10_000 = New CodeArchitects.VB6Library.VB6Label
        Me.Label9 = New CodeArchitects.VB6Library.VB6Label
        Me.Label6 = New CodeArchitects.VB6Library.VB6Label
        Me.Label1 = New CodeArchitects.VB6Library.VB6Label
        Me.Line5 = New CodeArchitects.VB6Library.VB6Line
        Me.Label2 = New CodeArchitects.VB6Library.VB6Label
        Me.Label3 = New CodeArchitects.VB6Library.VB6Label
        Me.Label4 = New CodeArchitects.VB6Library.VB6Label
        Me.Label7 = New CodeArchitects.VB6Library.VB6Label
        Me.Line1 = New CodeArchitects.VB6Library.VB6Line
        Me.Line2 = New CodeArchitects.VB6Library.VB6Line
        Me.Line3 = New CodeArchitects.VB6Library.VB6Line
        CType(Me.Line6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbQuality
        '
        Me.cbQuality.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbQuality.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbQuality.ItemDataValues = "80\r0\r70\r0\r60\r0\r50\r0"
        Me.cbQuality.Location = New System.Drawing.Point(232, 8)
        Me.cbQuality.Name = "cbQuality"
        Me.cbQuality.Size = New System.Drawing.Size(41, 21)
        Me.cbQuality.TabIndex = 26
        Me.cbQuality.Text = "80"
        '
        'btnNewFile
        '
        Me.btnNewFile.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnNewFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewFile.Image = CType(resources.GetObject("btnNewFile.Image"), System.Drawing.Image)
        Me.btnNewFile.Location = New System.Drawing.Point(296, 88)
        Me.btnNewFile.Name = "btnNewFile"
        Me.btnNewFile.Picture = CType(resources.GetObject("btnNewFile.Picture"), System.Drawing.Image)
        Me.btnNewFile.Size = New System.Drawing.Size(25, 25)
        Me.btnNewFile.Style = CodeArchitects.VB6Library.VBRUN.ButtonConstants.vbButtonGraphical
        Me.btnNewFile.TabIndex = 25
        Me.btnNewFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnNewFile.UseVisualStyleBackColor = False
        '
        'txtfilename
        '
        Me.txtfilename.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtfilename.BackColor = System.Drawing.SystemColors.Menu
        Me.txtfilename.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfilename.LinkItem = Nothing
        Me.txtfilename.LinkMode = CType(0, Short)
        Me.txtfilename.LinkTimeout = CType(50, Short)
        Me.txtfilename.Location = New System.Drawing.Point(120, 104)
        Me.txtfilename.Name = "txtfilename"
        Me.txtfilename.ReadOnly = True
        Me.txtfilename.Size = New System.Drawing.Size(169, 21)
        Me.txtfilename.TabIndex = 23
        '
        'txtfolder
        '
        Me.txtfolder.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtfolder.BackColor = System.Drawing.SystemColors.Menu
        Me.txtfolder.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfolder.LinkItem = Nothing
        Me.txtfolder.LinkMode = CType(0, Short)
        Me.txtfolder.LinkTimeout = CType(50, Short)
        Me.txtfolder.Location = New System.Drawing.Point(120, 80)
        Me.txtfolder.Name = "txtfolder"
        Me.txtfolder.ReadOnly = True
        Me.txtfolder.Size = New System.Drawing.Size(169, 21)
        Me.txtfolder.TabIndex = 21
        '
        'txtMapScale
        '
        Me.txtMapScale.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtMapScale.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtMapScale.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMapScale.LinkItem = Nothing
        Me.txtMapScale.LinkMode = CType(0, Short)
        Me.txtMapScale.LinkTimeout = CType(50, Short)
        Me.txtMapScale.Location = New System.Drawing.Point(96, 136)
        Me.txtMapScale.Name = "txtMapScale"
        Me.txtMapScale.ReadOnly = True
        Me.txtMapScale.Size = New System.Drawing.Size(113, 21)
        Me.txtMapScale.TabIndex = 20
        '
        'txtUsername
        '
        Me.txtUsername.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtUsername.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtUsername.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.LinkItem = Nothing
        Me.txtUsername.LinkMode = CType(0, Short)
        Me.txtUsername.LinkTimeout = CType(50, Short)
        Me.txtUsername.Location = New System.Drawing.Point(152, 224)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.ReadOnly = True
        Me.txtUsername.Size = New System.Drawing.Size(169, 21)
        Me.txtUsername.TabIndex = 19
        '
        'txtOrientation
        '
        Me.txtOrientation.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtOrientation.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtOrientation.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOrientation.LinkItem = Nothing
        Me.txtOrientation.LinkMode = CType(0, Short)
        Me.txtOrientation.LinkTimeout = CType(50, Short)
        Me.txtOrientation.Location = New System.Drawing.Point(144, 160)
        Me.txtOrientation.Name = "txtOrientation"
        Me.txtOrientation.ReadOnly = True
        Me.txtOrientation.Size = New System.Drawing.Size(113, 21)
        Me.txtOrientation.TabIndex = 18
        '
        'txtPage_Size
        '
        Me.txtPage_Size.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtPage_Size.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtPage_Size.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPage_Size.LinkItem = Nothing
        Me.txtPage_Size.LinkMode = CType(0, Short)
        Me.txtPage_Size.LinkTimeout = CType(50, Short)
        Me.txtPage_Size.Location = New System.Drawing.Point(96, 160)
        Me.txtPage_Size.Name = "txtPage_Size"
        Me.txtPage_Size.ReadOnly = True
        Me.txtPage_Size.Size = New System.Drawing.Size(41, 21)
        Me.txtPage_Size.TabIndex = 15
        '
        'cbMap_Type
        '
        Me.cbMap_Type.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbMap_Type.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbMap_Type.ItemDataValues = "Locality\ Map\r0\rFuel\ Age\ Map\r0\rSector\ Map\r0\rInformation\ Map\r0\rFire\ G" & _
            "rowth\ Map\r0\rOperations\ Overview\ Map\r0"
        Me.cbMap_Type.Location = New System.Drawing.Point(96, 40)
        Me.cbMap_Type.Name = "cbMap_Type"
        Me.cbMap_Type.Size = New System.Drawing.Size(145, 21)
        Me.cbMap_Type.TabIndex = 14
        Me.cbMap_Type.Text = "Sector Map"
        '
        'cbResolution
        '
        Me.cbResolution.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbResolution.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbResolution.ItemDataValues = "75\r0\r100\r0\r125\r0\r150\r0\r175\r0\r200\r0\r225\r0\r250\r0\r275\r0\r300\r0\r32" & _
            "5\r0\r350\r0\r375\r0\r400\r0\r425\r0\r450\r0\r475\r0\r500\r0\r525\r0\r550\r0\r57" & _
            "5\r0\r600\r0"
        Me.cbResolution.Location = New System.Drawing.Point(96, 8)
        Me.cbResolution.Name = "cbResolution"
        Me.cbResolution.Size = New System.Drawing.Size(49, 21)
        Me.cbResolution.TabIndex = 12
        Me.cbResolution.Text = "300"
        '
        'OK_btn
        '
        Me.OK_btn.BackColor = System.Drawing.SystemColors.Control
        Me.OK_btn.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK_btn.Location = New System.Drawing.Point(32, 320)
        Me.OK_btn.Name = "OK_btn"
        Me.OK_btn.Size = New System.Drawing.Size(113, 25)
        Me.OK_btn.TabIndex = 6
        Me.OK_btn.Text = "OK"
        Me.OK_btn.UseVisualStyleBackColor = False
        '
        'Cancel_btn
        '
        Me.Cancel_btn.BackColor = System.Drawing.SystemColors.Control
        Me.Cancel_btn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_btn.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel_btn.Location = New System.Drawing.Point(176, 320)
        Me.Cancel_btn.Name = "Cancel_btn"
        Me.Cancel_btn.Size = New System.Drawing.Size(113, 25)
        Me.Cancel_btn.TabIndex = 5
        Me.Cancel_btn.Text = "CANCEL"
        Me.Cancel_btn.UseVisualStyleBackColor = False
        '
        'txtRegion_Ex
        '
        Me.txtRegion_Ex.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtRegion_Ex.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtRegion_Ex.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegion_Ex.LinkItem = Nothing
        Me.txtRegion_Ex.LinkMode = CType(0, Short)
        Me.txtRegion_Ex.LinkTimeout = CType(50, Short)
        Me.txtRegion_Ex.Location = New System.Drawing.Point(152, 192)
        Me.txtRegion_Ex.Name = "txtRegion_Ex"
        Me.txtRegion_Ex.ReadOnly = True
        Me.txtRegion_Ex.Size = New System.Drawing.Size(41, 21)
        Me.txtRegion_Ex.TabIndex = 4
        '
        'txtFireNumber
        '
        Me.txtFireNumber.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtFireNumber.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtFireNumber.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFireNumber.LinkItem = Nothing
        Me.txtFireNumber.LinkMode = CType(0, Short)
        Me.txtFireNumber.LinkTimeout = CType(50, Short)
        Me.txtFireNumber.Location = New System.Drawing.Point(200, 192)
        Me.txtFireNumber.Name = "txtFireNumber"
        Me.txtFireNumber.ReadOnly = True
        Me.txtFireNumber.Size = New System.Drawing.Size(57, 21)
        Me.txtFireNumber.TabIndex = 3
        '
        'txtDept
        '
        Me.txtDept.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtDept.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtDept.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDept.LinkItem = Nothing
        Me.txtDept.LinkMode = CType(0, Short)
        Me.txtDept.LinkTimeout = CType(50, Short)
        Me.txtDept.Location = New System.Drawing.Point(152, 248)
        Me.txtDept.Name = "txtDept"
        Me.txtDept.ReadOnly = True
        Me.txtDept.Size = New System.Drawing.Size(169, 21)
        Me.txtDept.TabIndex = 2
        '
        'txtDate
        '
        Me.txtDate.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtDate.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.LinkItem = Nothing
        Me.txtDate.LinkMode = CType(0, Short)
        Me.txtDate.LinkTimeout = CType(50, Short)
        Me.txtDate.Location = New System.Drawing.Point(152, 280)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.ReadOnly = True
        Me.txtDate.Size = New System.Drawing.Size(81, 21)
        Me.txtDate.TabIndex = 1
        '
        'txtTime
        '
        Me.txtTime.Appearance = CodeArchitects.VB6Library.VB6AppearanceConstants.Flat
        Me.txtTime.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.txtTime.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTime.LinkItem = Nothing
        Me.txtTime.LinkMode = CType(0, Short)
        Me.txtTime.LinkTimeout = CType(50, Short)
        Me.txtTime.Location = New System.Drawing.Point(248, 280)
        Me.txtTime.Name = "txtTime"
        Me.txtTime.ReadOnly = True
        Me.txtTime.Size = New System.Drawing.Size(73, 21)
        Me.txtTime.TabIndex = 0
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.LinkItem = Nothing
        Me.Label13.LinkMode = CType(0, Short)
        Me.Label13.LinkTimeout = CType(50, Short)
        Me.Label13.Location = New System.Drawing.Point(176, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 17)
        Me.Label13.TabIndex = 27
        Me.Label13.Text = "Quality"
        '
        'Line6
        '
        Me.Line6.Container = Me
        Me.Line6.Name = "Line6"
        Me.Line6.Name6 = "Line6"
        Me.Line6.ParentForm = Me
        Me.Line6.X1 = 120.0!
        Me.Line6.X2 = 4800.0!
        Me.Line6.Y1 = 1080.0!
        Me.Line6.Y2 = 1080.0!
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.LinkItem = Nothing
        Me.Label12.LinkMode = CType(0, Short)
        Me.Label12.LinkTimeout = CType(50, Short)
        Me.Label12.Location = New System.Drawing.Point(8, 104)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(105, 17)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Output Filename"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.LinkItem = Nothing
        Me.Label11.LinkMode = CType(0, Short)
        Me.Label11.LinkTimeout = CType(50, Short)
        Me.Label11.Location = New System.Drawing.Point(8, 80)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(89, 17)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Output Folder"
        '
        'Line4
        '
        Me.Line4.Container = Me
        Me.Line4.Name = "Line4"
        Me.Line4.Name6 = "Line4"
        Me.Line4.ParentForm = Me
        Me.Line4.X1 = 120.0!
        Me.Line4.X2 = 4800.0!
        Me.Line4.Y1 = 1920.0!
        Me.Line4.Y2 = 1920.0!
        '
        'Label10_000
        '
        Me.Label10_000.BackColor = System.Drawing.SystemColors.Control
        Me.Label10_000.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10_000.LinkItem = Nothing
        Me.Label10_000.LinkMode = CType(0, Short)
        Me.Label10_000.LinkTimeout = CType(50, Short)
        Me.Label10_000.Location = New System.Drawing.Point(8, 136)
        Me.Label10_000.Name = "Label10_000"
        Me.Label10_000.Size = New System.Drawing.Size(89, 17)
        Me.Label10_000.TabIndex = 17
        Me.Label10_000.Text = "Map Scale"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.LinkItem = Nothing
        Me.Label9.LinkMode = CType(0, Short)
        Me.Label9.LinkTimeout = CType(50, Short)
        Me.Label9.Location = New System.Drawing.Point(8, 160)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(65, 17)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Page Size"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.LinkItem = Nothing
        Me.Label6.LinkMode = CType(0, Short)
        Me.Label6.LinkTimeout = CType(50, Short)
        Me.Label6.Location = New System.Drawing.Point(8, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 17)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Map Type"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.LinkItem = Nothing
        Me.Label1.LinkMode = CType(0, Short)
        Me.Label1.LinkTimeout = CType(50, Short)
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Resolution"
        '
        'Line5
        '
        Me.Line5.Container = Me
        Me.Line5.Name = "Line5"
        Me.Line5.Name6 = "Line5"
        Me.Line5.ParentForm = Me
        Me.Line5.X1 = 120.0!
        Me.Line5.X2 = 4800.0!
        Me.Line5.Y1 = 2760.0!
        Me.Line5.Y2 = 2760.0!
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.LinkItem = Nothing
        Me.Label2.LinkMode = CType(0, Short)
        Me.Label2.LinkTimeout = CType(50, Short)
        Me.Label2.Location = New System.Drawing.Point(8, 192)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Fire ID"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.LinkItem = Nothing
        Me.Label3.LinkMode = CType(0, Short)
        Me.Label3.LinkTimeout = CType(50, Short)
        Me.Label3.Location = New System.Drawing.Point(8, 224)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "User's Name"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.LinkItem = Nothing
        Me.Label4.LinkMode = CType(0, Short)
        Me.Label4.LinkTimeout = CType(50, Short)
        Me.Label4.Location = New System.Drawing.Point(8, 248)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(121, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "User's Department"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.LinkItem = Nothing
        Me.Label7.LinkMode = CType(0, Short)
        Me.Label7.LinkTimeout = CType(50, Short)
        Me.Label7.Location = New System.Drawing.Point(8, 280)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(89, 17)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Date / Time"
        '
        'Line1
        '
        Me.Line1.Container = Me
        Me.Line1.Name = "Line1"
        Me.Line1.Name6 = "Line1"
        Me.Line1.ParentForm = Me
        Me.Line1.X1 = 120.0!
        Me.Line1.X2 = 4800.0!
        Me.Line1.Y1 = 3240.0!
        Me.Line1.Y2 = 3240.0!
        '
        'Line2
        '
        Me.Line2.Container = Me
        Me.Line2.Name = "Line2"
        Me.Line2.Name6 = "Line2"
        Me.Line2.ParentForm = Me
        Me.Line2.X1 = 120.0!
        Me.Line2.X2 = 4800.0!
        Me.Line2.Y1 = 4080.0!
        Me.Line2.Y2 = 4080.0!
        '
        'Line3
        '
        Me.Line3.Container = Me
        Me.Line3.Name = "Line3"
        Me.Line3.Name6 = "Line3"
        Me.Line3.ParentForm = Me
        Me.Line3.X1 = 120.0!
        Me.Line3.X2 = 4800.0!
        Me.Line3.Y1 = 4560.0!
        Me.Line3.Y2 = 4560.0!
        '
        'frmFire_Export_Map
        '
        Me.AcceptButton = Me.OK_btn
        Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
        Me.CancelButton = Me.Cancel_btn
        Me.ClientSize = New System.Drawing.Size(329, 353)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbQuality)
        Me.Controls.Add(Me.btnNewFile)
        Me.Controls.Add(Me.txtfilename)
        Me.Controls.Add(Me.txtfolder)
        Me.Controls.Add(Me.txtMapScale)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.txtOrientation)
        Me.Controls.Add(Me.txtPage_Size)
        Me.Controls.Add(Me.cbMap_Type)
        Me.Controls.Add(Me.cbResolution)
        Me.Controls.Add(Me.OK_btn)
        Me.Controls.Add(Me.Cancel_btn)
        Me.Controls.Add(Me.txtRegion_Ex)
        Me.Controls.Add(Me.txtFireNumber)
        Me.Controls.Add(Me.txtDept)
        Me.Controls.Add(Me.txtDate)
        Me.Controls.Add(Me.txtTime)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10_000)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label7)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFire_Export_Map"
        Me.ShowInTaskbar = False
        Me.Text = "Export Map"
        CType(Me.Line6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


End Class
