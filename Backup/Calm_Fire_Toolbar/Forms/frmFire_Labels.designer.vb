<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFire_Labels
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdMoveLabel As System.Windows.Forms.Button
	Public WithEvents cmdAddLabel As System.Windows.Forms.Button
	Public WithEvents btnClose As System.Windows.Forms.Button
	Public WithEvents btnDelete As System.Windows.Forms.Button
	Public WithEvents tbCustomText As System.Windows.Forms.TextBox
	Public WithEvents cb_CustomText As System.Windows.Forms.CheckBox
	Public WithEvents btnEdit As System.Windows.Forms.Button
	Public WithEvents btnAdd As System.Windows.Forms.Button
	Public WithEvents lblList As System.Windows.Forms.ListBox
	Public WithEvents Line4 As System.Windows.Forms.Label
	Public WithEvents Line2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Line1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdMoveLabel = New System.Windows.Forms.Button
        Me.cmdAddLabel = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.tbCustomText = New System.Windows.Forms.TextBox
        Me.cb_CustomText = New System.Windows.Forms.CheckBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.lblList = New System.Windows.Forms.ListBox
        Me.Line4 = New System.Windows.Forms.Label
        Me.Line2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Line1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdMoveLabel
        '
        Me.cmdMoveLabel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMoveLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMoveLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMoveLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMoveLabel.Location = New System.Drawing.Point(256, 16)
        Me.cmdMoveLabel.Name = "cmdMoveLabel"
        Me.cmdMoveLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMoveLabel.Size = New System.Drawing.Size(41, 25)
        Me.cmdMoveLabel.TabIndex = 9
        Me.cmdMoveLabel.Text = "Move"
        Me.cmdMoveLabel.UseVisualStyleBackColor = False
        '
        'cmdAddLabel
        '
        Me.cmdAddLabel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddLabel.Location = New System.Drawing.Point(216, 16)
        Me.cmdAddLabel.Name = "cmdAddLabel"
        Me.cmdAddLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddLabel.Size = New System.Drawing.Size(41, 25)
        Me.cmdAddLabel.TabIndex = 8
        Me.cmdAddLabel.Text = "Add"
        Me.cmdAddLabel.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnClose.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnClose.Location = New System.Drawing.Point(224, 160)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnClose.Size = New System.Drawing.Size(73, 41)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "CLOSE"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDelete.Location = New System.Drawing.Point(216, 120)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDelete.Size = New System.Drawing.Size(81, 25)
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.Text = "DELETE ITEM"
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'tbCustomText
        '
        Me.tbCustomText.AcceptsReturn = True
        Me.tbCustomText.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.tbCustomText.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.tbCustomText.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCustomText.ForeColor = System.Drawing.SystemColors.WindowText
        Me.tbCustomText.Location = New System.Drawing.Point(16, 184)
        Me.tbCustomText.MaxLength = 0
        Me.tbCustomText.Name = "tbCustomText"
        Me.tbCustomText.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCustomText.Size = New System.Drawing.Size(185, 21)
        Me.tbCustomText.TabIndex = 4
        Me.tbCustomText.Text = "Enter Custom Text Here"
        '
        'cb_CustomText
        '
        Me.cb_CustomText.BackColor = System.Drawing.SystemColors.Control
        Me.cb_CustomText.Cursor = System.Windows.Forms.Cursors.Default
        Me.cb_CustomText.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cb_CustomText.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cb_CustomText.Location = New System.Drawing.Point(24, 160)
        Me.cb_CustomText.Name = "cb_CustomText"
        Me.cb_CustomText.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cb_CustomText.Size = New System.Drawing.Size(17, 17)
        Me.cb_CustomText.TabIndex = 3
        Me.cb_CustomText.Text = "Check1"
        Me.cb_CustomText.UseVisualStyleBackColor = False
        '
        'btnEdit
        '
        Me.btnEdit.BackColor = System.Drawing.SystemColors.Control
        Me.btnEdit.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEdit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEdit.Location = New System.Drawing.Point(216, 88)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEdit.Size = New System.Drawing.Size(81, 25)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "EDIT ITEM"
        Me.btnEdit.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.SystemColors.Control
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAdd.Location = New System.Drawing.Point(216, 56)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAdd.Size = New System.Drawing.Size(81, 25)
        Me.btnAdd.TabIndex = 1
        Me.btnAdd.Text = "ADD NEW"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'lblList
        '
        Me.lblList.BackColor = System.Drawing.SystemColors.Window
        Me.lblList.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblList.ItemHeight = 16
        Me.lblList.Items.AddRange(New Object() {"Date and Time", "Date", "Time", "Sector 1", "Sector 2", "Sector 3", "Division 1", "Division 2", "Division 3", "Active", "Inactive", "MU", "PO", "HT", "MT", "Landing Ground"})
        Me.lblList.Location = New System.Drawing.Point(16, 16)
        Me.lblList.Name = "lblList"
        Me.lblList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblList.Size = New System.Drawing.Size(185, 116)
        Me.lblList.TabIndex = 0
        '
        'Line4
        '
        Me.Line4.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line4.Location = New System.Drawing.Point(216, 48)
        Me.Line4.Name = "Line4"
        Me.Line4.Size = New System.Drawing.Size(80, 1)
        Me.Line4.TabIndex = 10
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(208, 160)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(1, 40)
        Me.Line2.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(48, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(129, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Use Custom Text"
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(16, 152)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(280, 1)
        Me.Line1.TabIndex = 12
        '
        'frmFire_Labels
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(303, 212)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdMoveLabel)
        Me.Controls.Add(Me.cmdAddLabel)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.tbCustomText)
        Me.Controls.Add(Me.cb_CustomText)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.lblList)
        Me.Controls.Add(Me.Line4)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Line1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(192, 186)
        Me.Name = "frmFire_Labels"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Standard Fire Labels"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class