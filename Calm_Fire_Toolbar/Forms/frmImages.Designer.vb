<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmImages
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
	Public WithEvents legend_restorer As CodeArchitects.VB6Library.VB6Image
	Public WithEvents Image1 As CodeArchitects.VB6Library.VB6Image
	Public WithEvents Saveimg As CodeArchitects.VB6Library.VB6Image
	Public WithEvents modifypoly As CodeArchitects.VB6Library.VB6Image
	Public WithEvents label As CodeArchitects.VB6Library.VB6Image
	Public WithEvents legend_limiter As CodeArchitects.VB6Library.VB6Image
	Public WithEvents log As CodeArchitects.VB6Library.VB6Image
	Public WithEvents export_Map As CodeArchitects.VB6Library.VB6Image
	Public WithEvents copyline As CodeArchitects.VB6Library.VB6Image
	Public WithEvents copypoint As CodeArchitects.VB6Library.VB6Image
	Public WithEvents copypoly As CodeArchitects.VB6Library.VB6Image
	Public WithEvents multilabel As CodeArchitects.VB6Library.VB6Image
	Public WithEvents exportgdb As CodeArchitects.VB6Library.VB6Image
	Public WithEvents symbols As CodeArchitects.VB6Library.VB6Image
	Public WithEvents email_view As CodeArchitects.VB6Library.VB6Image
	Public WithEvents email_gdb As CodeArchitects.VB6Library.VB6Image
	Public WithEvents email_boundary As CodeArchitects.VB6Library.VB6Image
	Public WithEvents email_layout As CodeArchitects.VB6Library.VB6Image
	Public WithEvents email As CodeArchitects.VB6Library.VB6Image
	Public WithEvents rulers As CodeArchitects.VB6Library.VB6Image
	Public WithEvents fire_div As CodeArchitects.VB6Library.VB6Image
	Public WithEvents new_line As CodeArchitects.VB6Library.VB6Image
	Public WithEvents new_poly As CodeArchitects.VB6Library.VB6Image
	Public WithEvents new_point As CodeArchitects.VB6Library.VB6Image
	Public WithEvents find As CodeArchitects.VB6Library.VB6Image
	Public WithEvents wizard As CodeArchitects.VB6Library.VB6Image
	Public WithEvents Open As CodeArchitects.VB6Library.VB6Image
	Public WithEvents Logon As CodeArchitects.VB6Library.VB6Image

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
 		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmImages))
		Me.legend_restorer = New CodeArchitects.VB6Library.VB6Image()
		Me.Image1 = New CodeArchitects.VB6Library.VB6Image()
		Me.Saveimg = New CodeArchitects.VB6Library.VB6Image()
		Me.modifypoly = New CodeArchitects.VB6Library.VB6Image()
		Me.label = New CodeArchitects.VB6Library.VB6Image()
		Me.legend_limiter = New CodeArchitects.VB6Library.VB6Image()
		Me.log = New CodeArchitects.VB6Library.VB6Image()
		Me.export_Map = New CodeArchitects.VB6Library.VB6Image()
		Me.copyline = New CodeArchitects.VB6Library.VB6Image()
		Me.copypoint = New CodeArchitects.VB6Library.VB6Image()
		Me.copypoly = New CodeArchitects.VB6Library.VB6Image()
		Me.multilabel = New CodeArchitects.VB6Library.VB6Image()
		Me.exportgdb = New CodeArchitects.VB6Library.VB6Image()
		Me.symbols = New CodeArchitects.VB6Library.VB6Image()
		Me.email_view = New CodeArchitects.VB6Library.VB6Image()
		Me.email_gdb = New CodeArchitects.VB6Library.VB6Image()
		Me.email_boundary = New CodeArchitects.VB6Library.VB6Image()
		Me.email_layout = New CodeArchitects.VB6Library.VB6Image()
		Me.email = New CodeArchitects.VB6Library.VB6Image()
		Me.rulers = New CodeArchitects.VB6Library.VB6Image()
		Me.fire_div = New CodeArchitects.VB6Library.VB6Image()
		Me.new_line = New CodeArchitects.VB6Library.VB6Image()
		Me.new_poly = New CodeArchitects.VB6Library.VB6Image()
		Me.new_point = New CodeArchitects.VB6Library.VB6Image()
		Me.find = New CodeArchitects.VB6Library.VB6Image()
		Me.wizard = New CodeArchitects.VB6Library.VB6Image()
		Me.Open = New CodeArchitects.VB6Library.VB6Image()
		Me.Logon = New CodeArchitects.VB6Library.VB6Image()
		CType(Me.legend_restorer, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Saveimg, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.modifypoly, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.label, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.legend_limiter, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.log, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.export_Map, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.copyline, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.copypoint, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.copypoly, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.multilabel, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.exportgdb, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.symbols, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.email_view, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.email_gdb, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.email_boundary, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.email_layout, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.email, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.rulers, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.fire_div, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.new_line, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.new_poly, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.new_point, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.find, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.wizard, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Open, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Logon, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		' legend_restorer
		'
		Me.legend_restorer.Name = "legend_restorer"
		Me.legend_restorer.Size = New System.Drawing.Size(19, 17)
		Me.legend_restorer.Location = New System.Drawing.Point(560, 16)
		Me.legend_restorer.Picture = CType(resources.GetObject("legend_restorer.Picture"), System.Drawing.Image)
		Me.legend_restorer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.legend_restorer.Stretch = False
		'
		' Image1
		'
		Me.Image1.Name = "Image1"
		Me.Image1.Size = New System.Drawing.Size(16, 16)
		Me.Image1.Location = New System.Drawing.Point(32, 16)
		Me.Image1.Picture = CType(resources.GetObject("Image1.Picture"), System.Drawing.Image)
		Me.Image1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Image1.Stretch = False
		'
		' Saveimg
		'
		Me.Saveimg.Name = "Saveimg"
		Me.Saveimg.Size = New System.Drawing.Size(16, 16)
		Me.Saveimg.Location = New System.Drawing.Point(616, 0)
		Me.Saveimg.Picture = CType(resources.GetObject("Saveimg.Picture"), System.Drawing.Image)
		Me.Saveimg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Saveimg.Stretch = False
		'
		' modifypoly
		'
		Me.modifypoly.Name = "modifypoly"
		Me.modifypoly.Size = New System.Drawing.Size(24, 24)
		Me.modifypoly.Location = New System.Drawing.Point(592, 0)
		Me.modifypoly.Picture = CType(resources.GetObject("modifypoly.Picture"), System.Drawing.Image)
		Me.modifypoly.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.modifypoly.Stretch = False
		'
		' label
		'
		Me.label.Name = "label"
		Me.label.Size = New System.Drawing.Size(16, 16)
		Me.label.Location = New System.Drawing.Point(576, 0)
		Me.label.Picture = CType(resources.GetObject("label.Picture"), System.Drawing.Image)
		Me.label.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label.Stretch = False
		'
		' legend_limiter
		'
		Me.legend_limiter.Name = "legend_limiter"
		Me.legend_limiter.Size = New System.Drawing.Size(16, 16)
		Me.legend_limiter.Location = New System.Drawing.Point(560, 0)
		Me.legend_limiter.Picture = CType(resources.GetObject("legend_limiter.Picture"), System.Drawing.Image)
		Me.legend_limiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.legend_limiter.Stretch = False
		'
		' log
		'
		Me.log.Name = "log"
		Me.log.Size = New System.Drawing.Size(32, 32)
		Me.log.Location = New System.Drawing.Point(528, 0)
		Me.log.Picture = CType(resources.GetObject("log.Picture"), System.Drawing.Image)
		Me.log.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.log.Stretch = False
		'
		' export_Map
		'
		Me.export_Map.Name = "export_Map"
		Me.export_Map.Size = New System.Drawing.Size(32, 32)
		Me.export_Map.Location = New System.Drawing.Point(496, 0)
		Me.export_Map.Picture = CType(resources.GetObject("export_Map.Picture"), System.Drawing.Image)
		Me.export_Map.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.export_Map.Stretch = False
		'
		' copyline
		'
		Me.copyline.Name = "copyline"
		Me.copyline.Size = New System.Drawing.Size(16, 16)
		Me.copyline.Location = New System.Drawing.Point(480, 0)
		Me.copyline.Picture = CType(resources.GetObject("copyline.Picture"), System.Drawing.Image)
		Me.copyline.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.copyline.Stretch = False
		'
		' copypoint
		'
		Me.copypoint.Name = "copypoint"
		Me.copypoint.Size = New System.Drawing.Size(16, 16)
		Me.copypoint.Location = New System.Drawing.Point(464, 0)
		Me.copypoint.Picture = CType(resources.GetObject("copypoint.Picture"), System.Drawing.Image)
		Me.copypoint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.copypoint.Stretch = False
		'
		' copypoly
		'
		Me.copypoly.Name = "copypoly"
		Me.copypoly.Size = New System.Drawing.Size(16, 16)
		Me.copypoly.Location = New System.Drawing.Point(448, 0)
		Me.copypoly.Picture = CType(resources.GetObject("copypoly.Picture"), System.Drawing.Image)
		Me.copypoly.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.copypoly.Stretch = False
		'
		' multilabel
		'
		Me.multilabel.Name = "multilabel"
		Me.multilabel.Size = New System.Drawing.Size(16, 16)
		Me.multilabel.Location = New System.Drawing.Point(432, 0)
		Me.multilabel.Picture = CType(resources.GetObject("multilabel.Picture"), System.Drawing.Image)
		Me.multilabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.multilabel.Stretch = False
		'
		' exportgdb
		'
		Me.exportgdb.Name = "exportgdb"
		Me.exportgdb.Size = New System.Drawing.Size(16, 16)
		Me.exportgdb.Location = New System.Drawing.Point(416, 0)
		Me.exportgdb.Picture = CType(resources.GetObject("exportgdb.Picture"), System.Drawing.Image)
		Me.exportgdb.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.exportgdb.Stretch = False
		'
		' symbols
		'
		Me.symbols.Name = "symbols"
		Me.symbols.Size = New System.Drawing.Size(16, 16)
		Me.symbols.Location = New System.Drawing.Point(400, 0)
		Me.symbols.Picture = CType(resources.GetObject("symbols.Picture"), System.Drawing.Image)
		Me.symbols.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.symbols.Stretch = False
		'
		' email_view
		'
		Me.email_view.Name = "email_view"
		Me.email_view.Size = New System.Drawing.Size(32, 32)
		Me.email_view.Location = New System.Drawing.Point(368, 0)
		Me.email_view.Picture = CType(resources.GetObject("email_view.Picture"), System.Drawing.Image)
		Me.email_view.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.email_view.Stretch = False
		'
		' email_gdb
		'
		Me.email_gdb.Name = "email_gdb"
		Me.email_gdb.Size = New System.Drawing.Size(32, 32)
		Me.email_gdb.Location = New System.Drawing.Point(336, 0)
		Me.email_gdb.Picture = CType(resources.GetObject("email_gdb.Picture"), System.Drawing.Image)
		Me.email_gdb.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.email_gdb.Stretch = False
		'
		' email_boundary
		'
		Me.email_boundary.Name = "email_boundary"
		Me.email_boundary.Size = New System.Drawing.Size(32, 32)
		Me.email_boundary.Location = New System.Drawing.Point(304, 0)
		Me.email_boundary.Picture = CType(resources.GetObject("email_boundary.Picture"), System.Drawing.Image)
		Me.email_boundary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.email_boundary.Stretch = False
		'
		' email_layout
		'
		Me.email_layout.Name = "email_layout"
		Me.email_layout.Size = New System.Drawing.Size(32, 32)
		Me.email_layout.Location = New System.Drawing.Point(272, 0)
		Me.email_layout.Picture = CType(resources.GetObject("email_layout.Picture"), System.Drawing.Image)
		Me.email_layout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.email_layout.Stretch = False
		'
		' email
		'
		Me.email.Name = "email"
		Me.email.Size = New System.Drawing.Size(32, 32)
		Me.email.Location = New System.Drawing.Point(240, 0)
		Me.email.Picture = CType(resources.GetObject("email.Picture"), System.Drawing.Image)
		Me.email.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.email.Stretch = False
		'
		' rulers
		'
		Me.rulers.Name = "rulers"
		Me.rulers.Size = New System.Drawing.Size(32, 32)
		Me.rulers.Location = New System.Drawing.Point(208, 0)
		Me.rulers.Picture = CType(resources.GetObject("rulers.Picture"), System.Drawing.Image)
		Me.rulers.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rulers.Stretch = False
		'
		' fire_div
		'
		Me.fire_div.Name = "fire_div"
		Me.fire_div.Size = New System.Drawing.Size(32, 32)
		Me.fire_div.Location = New System.Drawing.Point(176, 0)
		Me.fire_div.Picture = CType(resources.GetObject("fire_div.Picture"), System.Drawing.Image)
		Me.fire_div.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fire_div.Stretch = False
		'
		' new_line
		'
		Me.new_line.Name = "new_line"
		Me.new_line.Size = New System.Drawing.Size(24, 24)
		Me.new_line.Location = New System.Drawing.Point(152, 0)
		Me.new_line.Picture = CType(resources.GetObject("new_line.Picture"), System.Drawing.Image)
		Me.new_line.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.new_line.Stretch = False
		'
		' new_poly
		'
		Me.new_poly.Name = "new_poly"
		Me.new_poly.Size = New System.Drawing.Size(24, 24)
		Me.new_poly.Location = New System.Drawing.Point(128, 0)
		Me.new_poly.Picture = CType(resources.GetObject("new_poly.Picture"), System.Drawing.Image)
		Me.new_poly.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.new_poly.Stretch = False
		'
		' new_point
		'
		Me.new_point.Name = "new_point"
		Me.new_point.Size = New System.Drawing.Size(24, 24)
		Me.new_point.Location = New System.Drawing.Point(104, 0)
		Me.new_point.Picture = CType(resources.GetObject("new_point.Picture"), System.Drawing.Image)
		Me.new_point.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.new_point.Stretch = False
		'
		' find
		'
		Me.find.Name = "find"
		Me.find.Size = New System.Drawing.Size(24, 24)
		Me.find.Location = New System.Drawing.Point(80, 0)
		Me.find.Picture = CType(resources.GetObject("find.Picture"), System.Drawing.Image)
		Me.find.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.find.Stretch = False
		'
		' wizard
		'
		Me.wizard.Name = "wizard"
		Me.wizard.Size = New System.Drawing.Size(32, 32)
		Me.wizard.Location = New System.Drawing.Point(48, 0)
		Me.wizard.Picture = CType(resources.GetObject("wizard.Picture"), System.Drawing.Image)
		Me.wizard.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.wizard.Stretch = False
		'
		' Open
		'
		Me.Open.Name = "Open"
		Me.Open.Size = New System.Drawing.Size(16, 16)
		Me.Open.Location = New System.Drawing.Point(32, 0)
		Me.Open.Picture = CType(resources.GetObject("Open.Picture"), System.Drawing.Image)
		Me.Open.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Open.Stretch = False
		'
		' Logon
		'
		Me.Logon.Name = "Logon"
		Me.Logon.Size = New System.Drawing.Size(32, 32)
		Me.Logon.Location = New System.Drawing.Point(0, 0)
		Me.Logon.Picture = CType(resources.GetObject("Logon.Picture"), System.Drawing.Image)
		Me.Logon.Font = New System.Drawing.Font("Microsoft Sans Serif", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Logon.Stretch = False
		'
		' frmImages
		'
		Me.Name = "frmImages"
		Me.Caption = "Form1"
		Me.StartUpPosition = CodeArchitects.VB6Library.VBRUN.StartUpPositionConstants.vbStartUpWindowsDefault
		Me.Location = New System.Drawing.Point(4, 30)
		Me.ClientSize = New System.Drawing.Size(638, 40)

		Me.Controls.Add(legend_restorer)
		Me.Controls.Add(Image1)
		Me.Controls.Add(Saveimg)
		Me.Controls.Add(modifypoly)
		Me.Controls.Add(label)
		Me.Controls.Add(legend_limiter)
		Me.Controls.Add(log)
		Me.Controls.Add(export_Map)
		Me.Controls.Add(copyline)
		Me.Controls.Add(copypoint)
		Me.Controls.Add(copypoly)
		Me.Controls.Add(multilabel)
		Me.Controls.Add(exportgdb)
		Me.Controls.Add(symbols)
		Me.Controls.Add(email_view)
		Me.Controls.Add(email_gdb)
		Me.Controls.Add(email_boundary)
		Me.Controls.Add(email_layout)
		Me.Controls.Add(email)
		Me.Controls.Add(rulers)
		Me.Controls.Add(fire_div)
		Me.Controls.Add(new_line)
		Me.Controls.Add(new_poly)
		Me.Controls.Add(new_point)
		Me.Controls.Add(find)
		Me.Controls.Add(wizard)
		Me.Controls.Add(Open)
		Me.Controls.Add(Logon)
		CType(Me.legend_restorer, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Saveimg, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.modifypoly, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.label, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.legend_limiter, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.log, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.export_Map, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.copyline, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.copypoint, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.copypoly, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.multilabel, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.exportgdb, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.symbols, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.email_view, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.email_gdb, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.email_boundary, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.email_layout, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.email, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.rulers, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.fire_div, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.new_line, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.new_poly, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.new_point, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.find, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.wizard, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Open, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Logon, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub

#End Region


End Class
