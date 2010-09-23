' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Module modScaleSymbols

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: FIMT
	'Description
	'Called By
	'Calls
	'Accepts
	'Returns
	'******************************************************************************************************************
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Private Const c_sModuleFileName As String = "Incident Tools\ScaleSymbols.bas"

	'  Functions to scale text, marker, and line symbols up or down.
	
	'  External calls to this module should call only ScaleTextSymbol,
	'  ScaleAnyMarker or ScaleAnyLine.  These functions know how to identify
	'  the different types of symbols and handle them appropriately, and
	'  are also able to decompose multi-layer symbols and scale the individual
	'  layers as needed.
	
	' UPGRADE_INFO (#0551): The 'TextSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ScaleTextSymbol(ByRef TextSymbol As ESRI.ArcGIS.Display.ITextSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler
		
		'  Scale the text symbol and any offsets up or down.
		
19:
		TextSymbol.Size *= Multiplier
		
		Dim pSimpleTextSym As ESRI.ArcGIS.Display.ISimpleTextSymbol
22:
		pSimpleTextSym = TextSymbol
23:
		pSimpleTextSym.XOffset *= Multiplier
24:
		pSimpleTextSym.YOffset *= Multiplier
		
		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleTextSymbol " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'MarkerSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ScaleAnyMarker(ByRef MarkerSymbol As ESRI.ArcGIS.Display.IMarkerSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Determine what kind of marker symbol we have and call the
		'  appropriate specialty function to rescale it.  It turns out
		'  that even though cartographic, picture, and arrow markers
		'  all have additional properties, just setting the
		'  basic size and offset from the iMarkerSymbol interface
		'  is enough to get them to scale properly.  Only the
		'  SimpleMarkerSymbol has to be handled differently to
		'  set its outline width.
		
45:
		If TypeOf MarkerSymbol Is ESRI.ArcGIS.Display.IMultiLayerMarkerSymbol Then
46:
			Call ScaleMultilayerMarker(MarkerSymbol, Multiplier)
47:
		ElseIf TypeOf MarkerSymbol Is ESRI.ArcGIS.Display.ISimpleMarkerSymbol Then
48:
			Call ScaleSimpleMarker(MarkerSymbol, Multiplier)
49:
		Else
50:
			Call ScaleBasicMarker(MarkerSymbol, Multiplier)
51:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleAnyMarker " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'MarkerSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleMultilayerMarker(ByRef MarkerSymbol As ESRI.ArcGIS.Display.IMarkerSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Tear apart a multilayer marker symbol and scale the layers.
		
		Dim pMultiLayerMarker As ESRI.ArcGIS.Display.IMultiLayerMarkerSymbol
		Dim LayerCount As Short
		Dim i As Short
		Dim pThisLayer As ESRI.ArcGIS.Display.IMarkerSymbol
		
70:
		pMultiLayerMarker = MarkerSymbol
71:
		LayerCount = pMultiLayerMarker.LayerCount
72:
		For i = 0 To LayerCount - 1
73:
			pThisLayer = pMultiLayerMarker.Layer(i)
74:
			Call ScaleAnyMarker(pThisLayer, Multiplier)
75:
		Next

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleMultilayerMarker " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'MarkerSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleBasicMarker(ByRef MarkerSymbol As ESRI.ArcGIS.Display.IMarkerSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Rescales basic size, x offset and y Offset implemented by
		'  all marker symbol types.
		
90:
		MarkerSymbol.Size *= Multiplier
91:
		MarkerSymbol.XOffset *= Multiplier
92:
		MarkerSymbol.YOffset *= Multiplier

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleBasicMarker " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'MarkerSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleSimpleMarker(ByRef MarkerSymbol As ESRI.ArcGIS.Display.IMarkerSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler
		
		' Simple markers have an outline
		
105:
		Call ScaleBasicMarker(MarkerSymbol, Multiplier)
		
		Dim pSimpleMarker As ESRI.ArcGIS.Display.ISimpleMarkerSymbol
108:
		If TypeOf MarkerSymbol Is ESRI.ArcGIS.Display.ISimpleMarkerSymbol Then
109:
			pSimpleMarker = MarkerSymbol
110:
			pSimpleMarker.OutlineSize *= Multiplier
111:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleSimpleMarker " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Public Sub ScaleAnyLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Determine what kind of line symbol we have, and call the
		'  appropriate specialty function to rescale it.  Test for
		'  "cartographic" symbol last, because other types may also
		'  be cartographic type.  We want to get the more specific
		'  type, so test for them first.
		
131:
		If TypeOf LineSymbol Is ESRI.ArcGIS.Display.IMultiLayerLineSymbol Then
132:
			Call ScaleMultilayerLine(LineSymbol, Multiplier)
133:
		ElseIf TypeOf LineSymbol Is ESRI.ArcGIS.Display.ISimpleLineSymbol Then
134:
			Call ScaleSimpleLine(LineSymbol, Multiplier)
135:
		ElseIf TypeOf LineSymbol Is ESRI.ArcGIS.Display.IHashLineSymbol Then
136:
			Call ScaleHashLine(LineSymbol, Multiplier)
137:
		ElseIf TypeOf LineSymbol Is ESRI.ArcGIS.Display.IMarkerLineSymbol Then
138:
			Call ScaleMarkerLine(LineSymbol, Multiplier)
139:
		ElseIf TypeOf LineSymbol Is ESRI.ArcGIS.Display.IPictureLineSymbol Then
140:
			Call ScalePictureLine(LineSymbol, Multiplier)
141:
		ElseIf TypeOf LineSymbol Is ESRI.ArcGIS.Display.ICartographicLineSymbol Then
142:
			Call ScaleCartoLine(LineSymbol, Multiplier)
143:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleAnyLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleMultilayerLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Tear apart a multilayer line symbol and scale the layers.
		
		Dim pMultiLayerLine As ESRI.ArcGIS.Display.IMultiLayerLineSymbol
		Dim LayerCount As Short
		Dim i As Short
		Dim pThisLayer As ESRI.ArcGIS.Display.ILineSymbol
		
163:
		pMultiLayerLine = LineSymbol
164:
		LayerCount = pMultiLayerLine.LayerCount
165:
		For i = 0 To LayerCount - 1
166:
			pThisLayer = pMultiLayerLine.Layer(i)
167:
			Call ScaleAnyLine(pThisLayer, Multiplier)
168:
		Next

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleMultilayerLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleCartoLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Carto line symbol has a width, and possibly a template
		'  and offset.
		
		Dim pCartoLineSym As ESRI.ArcGIS.Display.ICartographicLineSymbol
		
185:
		pCartoLineSym = LineSymbol
186:
		pCartoLineSym.Width *= Multiplier
		
188:
		Call ScaleLineTemplate(LineSymbol, Multiplier)
189:
		Call ScaleLineOffset(LineSymbol, Multiplier)

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleCartoLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleSimpleLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Simple line symbol has a width and offset.
		
202:
		LineSymbol.Width *= Multiplier
203:
		Call ScaleLineOffset(LineSymbol, Multiplier)

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleSimpleLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScalePictureLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Picture Line symbol has an offset, xScale and yScale.
		
217:
		LineSymbol.Width *= Multiplier
		
		Dim pPictureLineSym As ESRI.ArcGIS.Display.IPictureLineSymbol
220:
		If TypeOf LineSymbol Is ESRI.ArcGIS.Display.IPictureLineSymbol Then
221:
			pPictureLineSym = LineSymbol
222:
			pPictureLineSym.Offset *= Multiplier
223:
			pPictureLineSym.XScale *= Multiplier
224:
			pPictureLineSym.YScale *= Multiplier
225:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScalePictureLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleHashLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Hash line has a width, template, offset, and a line symbol
		'  used to draw the hashes.
		
		Dim pHashLS As ESRI.ArcGIS.Display.IHashLineSymbol
		Dim pHashSymbol As ESRI.ArcGIS.Display.ILineSymbol
		
		'  Set the width - this controls the distance the hash lines
		'  extend out from the main line.
		
245:
		pHashLS = LineSymbol
246:
		pHashLS.Width *= Multiplier
		
		'  The hash symbol is the line symbol used for the
		'  hash marks.  It could be any type of line.
		
251:
		pHashSymbol = pHashLS.HashSymbol
252:
		Call ScaleAnyLine(pHashSymbol, Multiplier)
		
		'  Set the interval and offset.
		
256:
		Call ScaleLineTemplate(LineSymbol, Multiplier)
257:
		Call ScaleLineOffset(LineSymbol, Multiplier)

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleHashLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSymbol' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleMarkerLine(ByRef LineSymbol As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Marker line has a marker symbol, template, and offset.
		
		Dim pMarkerLS As ESRI.ArcGIS.Display.IMarkerLineSymbol
		Dim pMarker As ESRI.ArcGIS.Display.IMarkerSymbol
		
274:
		pMarkerLS = LineSymbol
275:
		pMarker = pMarkerLS.MarkerSymbol
276:
		Call ScaleAnyMarker(pMarker, Multiplier)
		
278:
		Call ScaleLineTemplate(LineSymbol, Multiplier)
279:
		Call ScaleLineOffset(LineSymbol, Multiplier)

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleMarkerLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSym' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleLineTemplate(ByRef LineSym As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Rescale the line template, if the symbol
		'  has one (some symbol types potentially have a
		'  template, but it might not be used.)
		
		Dim pLineProps As ESRI.ArcGIS.Display.ILineProperties
		Dim pTemplate As ESRI.ArcGIS.Display.ITemplate
		
298:
		If TypeOf LineSym Is ESRI.ArcGIS.Display.ILineProperties Then
299:
			pLineProps = LineSym
300:
			If pLineProps.Template IsNot Nothing Then
301:
				pTemplate = pLineProps.Template
302:
				pTemplate.Interval *= Multiplier
303:
			End If
304:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleLineTemplate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

	' UPGRADE_INFO (#0551): The 'LineSym' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ScaleLineOffset(ByRef LineSym As ESRI.ArcGIS.Display.ILineSymbol, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Rescale the line offset (distance left or right of the feature)
		
		Dim pLineProps As ESRI.ArcGIS.Display.ILineProperties
		
320:
		If TypeOf LineSym Is ESRI.ArcGIS.Display.ILineProperties Then
321:
			pLineProps = LineSym
322:
			pLineProps.Offset *= Multiplier
323:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "ScaleLineOffset " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

End Module
