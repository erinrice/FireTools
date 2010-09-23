<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSymLbl
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
	Public WithEvents cmdClose As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents fraColor As CodeArchitects.VB6Library.VB6Frame
	Public WithEvents cmdBW As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdColor As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Label6 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents fraSize As CodeArchitects.VB6Library.VB6Frame
	Public WithEvents cmdClearAll As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdSelectAll As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents lstLayerNames As CodeArchitects.VB6Library.VB6CheckedListBox
	Public WithEvents CmdIncreaseLine As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdDecreaseLine As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdIncreaseLabel As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdDecreaseLabel As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdBoldText As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdNormalText As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdDecreaseMarker As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents cmdIncreaseMarker As CodeArchitects.VB6Library.VB6CommandButton
	Public WithEvents Label7 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label4 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label2 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label1 As CodeArchitects.VB6Library.VB6Label
	Public WithEvents Label5 As CodeArchitects.VB6Library.VB6Label

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSymLbl))
		Me.cmdClose = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.fraColor = New CodeArchitects.VB6Library.VB6Frame()
		Me.cmdBW = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdColor = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.Label6 = New CodeArchitects.VB6Library.VB6Label()
		Me.fraSize = New CodeArchitects.VB6Library.VB6Frame()
		Me.cmdClearAll = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdSelectAll = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.lstLayerNames = New CodeArchitects.VB6Library.VB6CheckedListBox()
		Me.CmdIncreaseLine = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdDecreaseLine = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdIncreaseLabel = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdDecreaseLabel = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdBoldText = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdNormalText = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdDecreaseMarker = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.cmdIncreaseMarker = New CodeArchitects.VB6Library.VB6CommandButton()
		Me.Label7 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label4 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label2 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label1 = New CodeArchitects.VB6Library.VB6Label()
		Me.Label5 = New CodeArchitects.VB6Library.VB6Label()
		Me.fraColor.SuspendLayout()
		Me.fraSize.SuspendLayout()
		Me.SuspendLayout()
		'
		' cmdClose
		'
		Me.cmdClose.Name = "cmdClose"
		Me.cmdClose.Size = New System.Drawing.Size(69, 27)
		Me.cmdClose.Location = New System.Drawing.Point(398, 338)
		Me.AcceptButton = Me.cmdClose
		Me.CancelButton = Me.cmdClose
		Me.cmdClose.Caption = "Close"
		Me.cmdClose.TabIndex = 13
		Me.ToolTip1.SetToolTip(Me.cmdClose, "Apply changes and close form")
		Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClose.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' fraColor
		'
		Me.fraColor.Name = "fraColor"
		Me.fraColor.Size = New System.Drawing.Size(461, 82)
		Me.fraColor.Location = New System.Drawing.Point(7, 46)
		Me.fraColor.Caption = "Reset fire incident symbols to:"
		Me.fraColor.TabIndex = 14
		Me.fraColor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraColor.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdBW
		'
		Me.cmdBW.Name = "cmdBW"
		Me.cmdBW.Size = New System.Drawing.Size(177, 27)
		Me.cmdBW.Location = New System.Drawing.Point(259, 22)
		Me.cmdBW.Caption = "Standard Black && White Symbols"
		Me.cmdBW.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.cmdBW, "Use standard black and white symbols")
		Me.cmdBW.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBW.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdColor
		'
		Me.cmdColor.Name = "cmdColor"
		Me.cmdColor.Size = New System.Drawing.Size(173, 27)
		Me.cmdColor.Location = New System.Drawing.Point(26, 21)
		Me.cmdColor.Caption = "Standard Colour Symbols"
		Me.cmdColor.TabIndex = 0
		Me.ToolTip1.SetToolTip(Me.cmdColor, "Revert to standard symbology")
		Me.cmdColor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdColor.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label6
		'
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(433, 18)
		Me.Label6.Location = New System.Drawing.Point(16, 55)
		Me.Label6.Alignment = CodeArchitects.VB6Library.VBRUN.AlignmentConstants.vbCenter
		Me.Label6.BackColor = FromOleColor6(CInt(&H80000018))
		Me.Label6.Caption = "Warning:  This will restore the symbol sizes, colors, and classifications to their defaults."
		Me.Label6.ForeColor = FromOleColor6(CInt(&H80000017))
		Me.Label6.TabIndex = 17
		Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.AutoSize = False
		'
		' fraSize
		'
		Me.fraSize.Name = "fraSize"
		Me.fraSize.Size = New System.Drawing.Size(462, 197)
		Me.fraSize.Location = New System.Drawing.Point(7, 136)
		Me.fraSize.Caption = "Change these layers:  "
		Me.fraSize.TabIndex = 15
		Me.fraSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraSize.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdClearAll
		'
		Me.cmdClearAll.Name = "cmdClearAll"
		Me.cmdClearAll.Size = New System.Drawing.Size(69, 27)
		Me.cmdClearAll.Location = New System.Drawing.Point(95, 163)
		Me.cmdClearAll.Caption = "Clear All"
		Me.cmdClearAll.TabIndex = 4
		Me.cmdClearAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClearAll.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdSelectAll
		'
		Me.cmdSelectAll.Name = "cmdSelectAll"
		Me.cmdSelectAll.Size = New System.Drawing.Size(69, 27)
		Me.cmdSelectAll.Location = New System.Drawing.Point(10, 163)
		Me.cmdSelectAll.Caption = "Select All"
		Me.cmdSelectAll.TabIndex = 3
		Me.cmdSelectAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSelectAll.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' lstLayerNames
		'
		Me.lstLayerNames.Name = "lstLayerNames"
		Me.lstLayerNames.Size = New System.Drawing.Size(156, 124)
		Me.lstLayerNames.Location = New System.Drawing.Point(9, 19)
		Me.lstLayerNames.TabIndex = 2
		Me.lstLayerNames.CheckOnClick = True
		Me.lstLayerNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		'
		' CmdIncreaseLine
		'
		Me.CmdIncreaseLine.Name = "CmdIncreaseLine"
		Me.CmdIncreaseLine.Size = New System.Drawing.Size(69, 27)
		Me.CmdIncreaseLine.Location = New System.Drawing.Point(302, 56)
		Me.CmdIncreaseLine.Caption = "Wider"
		Me.CmdIncreaseLine.TabIndex = 7
		Me.CmdIncreaseLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CmdIncreaseLine.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdDecreaseLine
		'
		Me.cmdDecreaseLine.Name = "cmdDecreaseLine"
		Me.cmdDecreaseLine.Size = New System.Drawing.Size(69, 27)
		Me.cmdDecreaseLine.Location = New System.Drawing.Point(387, 56)
		Me.cmdDecreaseLine.Caption = "Narrower"
		Me.cmdDecreaseLine.TabIndex = 8
		Me.cmdDecreaseLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDecreaseLine.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdIncreaseLabel
		'
		Me.cmdIncreaseLabel.Name = "cmdIncreaseLabel"
		Me.cmdIncreaseLabel.Size = New System.Drawing.Size(69, 27)
		Me.cmdIncreaseLabel.Location = New System.Drawing.Point(302, 94)
		Me.cmdIncreaseLabel.Caption = "Larger"
		Me.cmdIncreaseLabel.TabIndex = 9
		Me.cmdIncreaseLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdIncreaseLabel.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdDecreaseLabel
		'
		Me.cmdDecreaseLabel.Name = "cmdDecreaseLabel"
		Me.cmdDecreaseLabel.Size = New System.Drawing.Size(69, 27)
		Me.cmdDecreaseLabel.Location = New System.Drawing.Point(387, 94)
		Me.cmdDecreaseLabel.Caption = "Smaller"
		Me.cmdDecreaseLabel.TabIndex = 10
		Me.cmdDecreaseLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDecreaseLabel.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdBoldText
		'
		Me.cmdBoldText.Name = "cmdBoldText"
		Me.cmdBoldText.Size = New System.Drawing.Size(69, 27)
		Me.cmdBoldText.Location = New System.Drawing.Point(302, 131)
		Me.cmdBoldText.Caption = "Bold"
		Me.cmdBoldText.TabIndex = 11
		Me.cmdBoldText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBoldText.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdNormalText
		'
		Me.cmdNormalText.Name = "cmdNormalText"
		Me.cmdNormalText.Size = New System.Drawing.Size(69, 27)
		Me.cmdNormalText.Location = New System.Drawing.Point(387, 131)
		Me.cmdNormalText.Caption = "Normal"
		Me.cmdNormalText.TabIndex = 12
		Me.cmdNormalText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNormalText.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdDecreaseMarker
		'
		Me.cmdDecreaseMarker.Name = "cmdDecreaseMarker"
		Me.cmdDecreaseMarker.Size = New System.Drawing.Size(69, 27)
		Me.cmdDecreaseMarker.Location = New System.Drawing.Point(387, 19)
		Me.cmdDecreaseMarker.Caption = "Smaller"
		Me.cmdDecreaseMarker.TabIndex = 6
		Me.cmdDecreaseMarker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDecreaseMarker.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' cmdIncreaseMarker
		'
		Me.cmdIncreaseMarker.Name = "cmdIncreaseMarker"
		Me.cmdIncreaseMarker.Size = New System.Drawing.Size(69, 27)
		Me.cmdIncreaseMarker.Location = New System.Drawing.Point(302, 19)
		Me.cmdIncreaseMarker.Caption = "Larger"
		Me.cmdIncreaseMarker.TabIndex = 5
		Me.cmdIncreaseMarker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdIncreaseMarker.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label7
		'
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(83, 18)
		Me.Label7.Location = New System.Drawing.Point(190, 135)
		Me.Label7.Caption = "Make text style:"
		Me.Label7.TabIndex = 21
		Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.AutoSize = False
		Me.Label7.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label4
		'
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(110, 18)
		Me.Label4.Location = New System.Drawing.Point(190, 98)
		Me.Label4.Caption = "Make labels 10%:"
		Me.Label4.TabIndex = 20
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.AutoSize = False
		Me.Label4.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label2
		'
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(111, 18)
		Me.Label2.Location = New System.Drawing.Point(191, 61)
		Me.Label2.Caption = "Make lines 10%:"
		Me.Label2.TabIndex = 19
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.AutoSize = False
		Me.Label2.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label1
		'
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(124, 18)
		Me.Label1.Location = New System.Drawing.Point(190, 23)
		Me.Label1.Caption = "Make markers 10%:"
		Me.Label1.TabIndex = 18
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.AutoSize = False
		Me.Label1.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' Label5
		'
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(415, 51)
		Me.Label5.Location = New System.Drawing.Point(14, 8)
		Me.Label5.Caption = "This command lets you reset the symbology and classifications for the Fire layers to their defaults, or change symbol properties on a layer-by-layer basis."
		Me.Label5.TabIndex = 16
		Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.AutoSize = False
		Me.Label5.BackColor = FromOleColor6(CInt(&H8000000F))
		'
		' frmSymLbl
		'
		Me.Name = "frmSymLbl"
		Me.BorderStyle = CodeArchitects.VB6Library.VBRUN.FormBorderStyleConstants.vbFixedDouble
		Me.Caption = "Symbols and Labels"
		Me.ControlBox = False
		Me.Icon = Nothing
		Me.MaxButton = False
		Me.MinButton = False
		Me.ShowInTaskbar = False
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpOwner
		Me.Location = New System.Drawing.Point(410, 224)
		Me.ClientSize = New System.Drawing.Size(472, 372)

		Me.Controls.Add(cmdClose)
		Me.Controls.Add(fraColor)
		Me.fraColor.Controls.Add(cmdBW)
		Me.fraColor.Controls.Add(cmdColor)
		Me.fraColor.Controls.Add(Label6)
		Me.Controls.Add(fraSize)
		Me.fraSize.Controls.Add(cmdClearAll)
		Me.fraSize.Controls.Add(cmdSelectAll)
		Me.fraSize.Controls.Add(lstLayerNames)
		Me.fraSize.Controls.Add(CmdIncreaseLine)
		Me.fraSize.Controls.Add(cmdDecreaseLine)
		Me.fraSize.Controls.Add(cmdIncreaseLabel)
		Me.fraSize.Controls.Add(cmdDecreaseLabel)
		Me.fraSize.Controls.Add(cmdBoldText)
		Me.fraSize.Controls.Add(cmdNormalText)
		Me.fraSize.Controls.Add(cmdDecreaseMarker)
		Me.fraSize.Controls.Add(cmdIncreaseMarker)
		Me.fraSize.Controls.Add(Label7)
		Me.fraSize.Controls.Add(Label4)
		Me.fraSize.Controls.Add(Label2)
		Me.fraSize.Controls.Add(Label1)
		Me.Controls.Add(Label5)
		Me.fraColor.ResumeLayout(False)
		Me.fraSize.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
