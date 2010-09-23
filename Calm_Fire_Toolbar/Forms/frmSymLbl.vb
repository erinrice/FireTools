' --------------------------------------------------------------------------------
' Code generated automatically by Code Architects' VB Migration Partner
' --------------------------------------------------------------------------------

Option Strict Off      ' Code migrated from VB6 has Option Strict disabled by default

Friend Class frmSymLbl

	'******************************************************************************************************************
	'Date: 06 / 06 / 06
	'Author: FIMT
	'Description
	'Called By
	'Calls: None
	'Accepts
	'Returns: None
	'******************************************************************************************************************

	Public Event ApplyColor(ByRef UseColor As Boolean, ByRef Points As Boolean, ByRef Lines As Boolean, ByRef Perimeter As Boolean)
	Public Event ApplySymbolProps(ByRef dMarkerSize As Double, ByRef dTextSize As Double, ByRef bBold As Boolean, ByRef Points As Boolean, ByRef Lines As Boolean, ByRef Perimeter As Boolean)
	' UPGRADE_INFO (#0701): The 'Cancelled' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Public Cancelled As Boolean
	
	Private Const c_sModuleFileName As String = "D:\Projects\FIMT9x\Application\Incident Tools\frmSymLbl.frm"
	
	Private p_polyfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_linefeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_pointfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_divfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	Private p_annofeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
	
	Private Sub cmdBW_Click() Handles cmdBW.Click
		On Error GoTo ErrorHandler
		
53:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbHourglass
54:
		Call LoadSymbols(False)
55:
		Call RefreshMap(m_pApp)
56:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbDefault
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdBW_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
61:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbDefault
	End Sub

	Private Sub cmdClearAll_Click() Handles cmdClearAll.Click
		On Error GoTo ErrorHandler
		
67:
		Call SetAllLayers(False)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdClearAll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdClose_Click() Handles cmdClose.Click
		On Error GoTo ErrorHandler
		
78:
		Unload6(Me)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdClose_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdColor_Click() Handles cmdColor.Click
		On Error GoTo ErrorHandler
		
88:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbHourglass
89:
		Call LoadSymbols(True)
90:
		Call RefreshMap(m_pApp)
91:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbDefault
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdColor_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
96:
		' UPGRADE_WARNING (#80F4): The Screen6.MousePointer property sets or returns the MousePointer property of the active form, but only if it's a VB6Form.
		Screen6.MousePointer = VBRUN.MousePointerConstants.vbDefault
	End Sub

	Private Sub cmdDecreaseLabel_Click() Handles cmdDecreaseLabel.Click
		On Error GoTo ErrorHandler
		
102:
		Call ChangeTextSizes(0.9)
103:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdDecreaseLabel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdIncreaseLabel_Click() Handles cmdIncreaseLabel.Click
		On Error GoTo ErrorHandler
		
113:
		Call ChangeTextSizes(1.1)
114:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdIncreaseLabel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdIncreaseMarker_Click() Handles cmdIncreaseMarker.Click
		On Error GoTo ErrorHandler
		
124:
		Call ChangeMarkerSizes(1.1)
125:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdIncreaseMarker_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdDecreaseMarker_Click() Handles cmdDecreaseMarker.Click
		On Error GoTo ErrorHandler
		
135:
		Call ChangeMarkerSizes(0.9)
136:
		Call RefreshMap(m_pApp)

		Exit Sub
ErrorHandler:
		HandleError(True, "cmdDecreaseMarker_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub CmdIncreaseLine_Click() Handles CmdIncreaseLine.Click
		On Error GoTo ErrorHandler
		
147:
		Call ChangeLineWidths(1.1)
148:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "CmdIncreaseLine_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdDecreaseLine_Click() Handles cmdDecreaseLine.Click
		On Error GoTo ErrorHandler
		
158:
		Call ChangeLineWidths(0.9)
159:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdDecreaseLine_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdNormalText_Click() Handles cmdNormalText.Click
		On Error GoTo ErrorHandler
		
169:
		Call ChangeTextStyles(False)
170:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdNormalText_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub cmdBoldText_Click() Handles cmdBoldText.Click
		On Error GoTo ErrorHandler
		
180:
		Call ChangeTextStyles(True)
181:
		Call RefreshMap(m_pApp)
		
		Exit Sub
ErrorHandler:
		HandleError(True, "cmdBoldText_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ChangeMarkerSizes(ByRef Multiplier As Double)
		'On Error GoTo ErrorHandler
		
		'  Get a set of layers whose marker sizes we need to change
		
		Dim pLayerSet As ESRI.ArcGIS.esriSystem.ISet
		
194:
		pLayerSet = GetCheckedLayers()
		
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		
198:
		If pLayerSet IsNot Nothing Then
199:
			If pLayerSet.Count > 0 Then
200:
				pLayerSet.Reset()
201:
				player = pLayerSet.Next()
				
202:
				Do While player IsNot Nothing
203:
					pFeatureLayer = player
					
204:
					If pFeatureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
						'
						'  Here's where we set the symbology
						'
208:
						Call IncrementPointMarkers(pFeatureLayer, Multiplier)
209:
					ElseIf pFeatureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
						'
						'  Set line decorations or hatch length for line symbols
						
213:
						Call IncrementLineMarkers(pFeatureLayer, Multiplier)
214:
					End If ' Shapetype is point
					'
216:
					player = pLayerSet.Next()
					'
218:
				Loop ' for each layer
219:
			End If ' Layercount > 0
220:
		End If ' Layerset not nothing
		
		Exit Sub
		' UPGRADE_INFO (#0521): Unreachable code detected.
ErrorHandler:
		'HandleError False, "ChangeMarkerSizes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1
	End Sub

	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ChangeLineWidths(ByRef Multiplier As Double)
		
		On Error GoTo ErrorHandler
		
		'  Get a set of layers whose line widths we need to change
		
		Dim pLayerSet As ESRI.ArcGIS.esriSystem.ISet
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		
238:
		pLayerSet = GetCheckedLayers()
239:
		If pLayerSet IsNot Nothing Then
240:
			If pLayerSet.Count > 0 Then
241:
				pLayerSet.Reset()
242:
				player = pLayerSet.Next()
243:
				Do While player IsNot Nothing
244:
					pFeatureLayer = player
245:
					If pFeatureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
						' Here's where we set the symbology
249:
						Call IncrementLineWidths(pFeatureLayer, Multiplier)
250:
					End If ' shape type is polyline
252:
					player = pLayerSet.Next()
254:
				Loop ' for each layer
255:
			End If ' Layercount > 0
256:
		End If ' Layerset not nothing
		
		Exit Sub
ErrorHandler:
		HandleError(False, "ChangeLineWidths " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub
	
	' UPGRADE_INFO (#0551): The 'pfeatlyr' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub AnnoUpdate(ByRef pfeatlyr As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Multiplier As Double)
		' UPGRADE_INFO (#05B1): The 'esriEditor' variable wasn't declared explicitly.
		Dim esriEditor As Object = Nothing
        '		On Error GoTo ErrorHandler

        Try


800:
            ' UPGRADE_INFO (#0701): The 'intAnnoSize' member has been removed because it isn't used anywhere in current application.
            ' EXCLUDED: Dim intAnnoSize As Short

801:
            Dim pFeatCur As ESRI.ArcGIS.Geodatabase.IFeatureCursor
802:
            Dim peditor As ESRI.ArcGIS.Editor.IEditor
803:
            Dim pworkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
804:
            Dim pdoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
805:
            Dim pMap As ESRI.ArcGIS.Carto.IMap
806:
            Dim pFeat As ESRI.ArcGIS.Geodatabase.IFeature
807:
            Dim pAnnoFeat As ESRI.ArcGIS.Carto.IAnnotationFeature
808:
            Dim pTextEl As ESRI.ArcGIS.Carto.ITextElement
809:
            Dim pTextSym As ESRI.ArcGIS.Display.ITextSymbol
810:
            Dim pElem As ESRI.ArcGIS.Carto.IElement
811:
            Dim pid As New ESRI.ArcGIS.esriSystem.UIDClass

            'Get a reference to the editor extension
812:
            pid.Value = "esriEditor.Editor"
813:
            peditor = m_pApp.FindExtensionByCLSID(pid)

814:
            pdoc = m_pApp.Document
815:
            pMap = pdoc.FocusMap

            Dim pExt As ESRI.ArcGIS.esriSystem.IExtension = FindMXExtByName(m_pApp, g_FIRE_EXTENSION)

            Dim myIncidentInfo As clsFire_Incident_Info

            If TypeOf pExt Is clsFire_Incident_Info Then

                myIncidentInfo = pExt

            End If

816:
            pworkspace = myIncidentInfo.IncidentWorkspace
817:
            pFeatCur = pfeatlyr.FeatureClass.Update(Nothing, False)

818:
            '                If peditor.EditState <> GetDefaultMember6(esriEditor.esriStateEditing) Then
            If peditor.EditState <> ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
                peditor.StartEditing(pworkspace)
            End If

819:
            peditor.StartOperation()
820:
            pFeat = pFeatCur.NextFeature()

821:
            Do While pFeat IsNot Nothing
822:
                If TypeOf pFeat Is ESRI.ArcGIS.Carto.IAnnotationFeature Then
823:
                    pAnnoFeat = pFeat
824:
                    pElem = pAnnoFeat.Annotation
825:
                    pTextEl = pElem
826:
                    pTextSym = pTextEl.Symbol

                    'intAnnoSize = pTextEl.Symbol.Size
827:
                    pTextSym.Size *= Multiplier
                    'pTextSym.Angle = 93
                    'pTextSym.Size = 20
                    'pTextSym.VerticalAlignment = esriTVACenter
828:
                    If pTextEl Is Nothing Then
829:
                        ' MsgBox "Don't have a text element.  Bye.", vbCritical
830:
                        Exit Sub
                        ' UPGRADE_INFO (#0521): Unreachable code detected.
831:
                    End If
832:
                    pTextEl.Symbol = pTextSym
833:
                    pTextSym = pTextEl.Symbol
834:
                    pElem = pTextEl
835:
                    pAnnoFeat.Annotation = pElem
836:
                    pFeat = pAnnoFeat
837:
                    pFeat.Store()
838:
                    pFeatCur.UpdateFeature(pFeat)
839:
                End If
840:
                pFeat = pFeatCur.NextFeature()
841:
            Loop

842:
            peditor.StopOperation("Update Anno")

            '            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        'ErrorHandler:
        '       HandleError(False, "ChangeTextSizes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub
	
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ChangeTextSizes(ByRef Multiplier As Double)
		On Error GoTo ErrorHandler
		
		'  Get a set of layers whose label sizes we need to change
		
		Dim pLayerSet As ESRI.ArcGIS.esriSystem.ISet
269:
		pLayerSet = GetCheckedLayers()
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		
273:
		If pLayerSet IsNot Nothing Then
274:
			If pLayerSet.Count > 0 Then
275:
				pLayerSet.Reset()
276:
				player = pLayerSet.Next()
277:
				Do While player IsNot Nothing
					pFeatureLayer = player
					If TypeOf player Is ESRI.ArcGIS.Carto.IAnnotationLayer Then
						'Dim pAnnoClass As IAnnoClass
						'Set pAnnoClass = pfeaturelayer.featureclass.Extension
						'Dim ptextsymbol As ITextSymbol
						
						AnnoUpdate(pFeatureLayer, Multiplier)
						
						player = pLayerSet.Next()
					Else
278:
						'
						'  Determine if there are labels for this layer
						'  change the label size if there are.
						'
283:
						Call IncrementLabelSize(pFeatureLayer, Multiplier)
284:
						player = pLayerSet.Next()
					End If
285:
				Loop ' for each layer
286:
			End If ' Layercount > 0
287:
		End If ' Layerset not nothing

		Exit Sub
ErrorHandler:
		HandleError(False, "ChangeTextSizes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'MakeBold' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub ChangeTextStyles(ByRef MakeBold As Boolean)
		On Error GoTo ErrorHandler
		
		'  Get a set of layers whose fonts we need to change
		
		Dim pLayerSet As ESRI.ArcGIS.esriSystem.ISet
302:
		pLayerSet = GetCheckedLayers()
		Dim player As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		
306:
		If pLayerSet IsNot Nothing Then
307:
			If pLayerSet.Count > 0 Then
308:
				pLayerSet.Reset()
309:
				player = pLayerSet.Next()
310:
				Do While player IsNot Nothing
311:
					pFeatureLayer = player
					'
					'  Change the font to bold or normal.
					'
315:
					Call SetBoldText(pFeatureLayer, MakeBold)
316:
					player = pLayerSet.Next()
317:
				Loop ' for each layer
318:
			End If ' Layercount > 0
319:
		End If ' Layerset not nothing

		Exit Sub
ErrorHandler:
		HandleError(False, "ChangeTextStyles " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Function GetCheckedLayers() As ESRI.ArcGIS.esriSystem.ISet
		On Error GoTo ErrorHandler

		Dim LayerSet As ESRI.ArcGIS.esriSystem.ISet
		Dim i As Short
		' UPGRADE_INFO (#0701): The 'iLayerIndex' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim iLayerIndex As Integer
		
336:
		LayerSet = New ESRI.ArcGIS.esriSystem.Set()
		
		'  Go through the list box and return a set of checked FIMT layers.
		
340:
		For i = 0 To lstLayerNames.ListCount - 1
341:
			If lstLayerNames.Selected(i) = True Then
				If lstLayerNames.List(lstLayerNames.ItemData(i)) = modGlobals.g_FIRE_BOUNDARY_NAME Then
					LayerSet.Add(p_polyfeatlayer)
				ElseIf lstLayerNames.List(lstLayerNames.ItemData(i)) = modGlobals.g_FIRE_LINE_NAME Then
					LayerSet.Add(p_linefeatlayer)
				ElseIf lstLayerNames.List(lstLayerNames.ItemData(i)) = modGlobals.g_FIRE_POINT_NAME Then
					LayerSet.Add(p_pointfeatlayer)
				ElseIf lstLayerNames.List(lstLayerNames.ItemData(i)) = modGlobals.g_FIRE_ASSIGNMENT_NAME Then
					LayerSet.Add(p_divfeatlayer)
				ElseIf lstLayerNames.List(lstLayerNames.ItemData(i)) = modGlobals.g_FIRE_ANNO_NAME Then
					LayerSet.Add(p_annofeatlayer)
				End If
368:
			End If
369:
		Next
		
371:
		Return LayerSet
		
ErrorHandler:
		HandleError(False, "GetCheckedLayers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Function

	Private Sub cmdSelectAll_Click() Handles cmdSelectAll.Click
		On Error GoTo ErrorHandler
		
381:
		Call SetAllLayers(True)

		Exit Sub
ErrorHandler:
		HandleError(True, "cmdSelectAll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'Checked' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub SetAllLayers(ByRef Checked As Boolean)
		On Error GoTo ErrorHandler

		Dim i As Short
395:
		For i = 0 To lstLayerNames.ListCount - 1
396:
			lstLayerNames.Selected(i) = Checked
397:
		Next
		
399:
		lstLayerNames.ListIndex = -1
400:
		Call EnableButtons()

		Exit Sub
ErrorHandler:
		HandleError(False, "SetAllLayers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'FLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub IncrementLabelSize(ByRef FLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		Dim pGeoFL As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim colAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerPropertiesCollection
		Dim pAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerProperties
		Dim pLabelEngineProps As ESRI.ArcGIS.Carto.ILabelEngineLayerProperties
		Dim pTextSymbol As ESRI.ArcGIS.Display.ITextSymbol
		' UPGRADE_INFO (#0701): The 'pFont' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFont As Font
		
420:
		pGeoFL = FLayer
421:
		colAnnoLayerProps = pGeoFL.AnnotationProperties
422:
		colAnnoLayerProps.QueryItem(0, pAnnoLayerProps, Nothing, Nothing)
423:
		pLabelEngineProps = pAnnoLayerProps
424:
		pTextSymbol = pLabelEngineProps.Symbol
		
426:
		Call ScaleTextSymbol(pTextSymbol, Multiplier)

		Exit Sub
ErrorHandler:
		HandleError(False, "IncrementLabelSize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub
	
	' UPGRADE_INFO (#0551): The 'FLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub IncrementPointMarkers(ByRef FLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		'  Change the marker size for each class in the specified layer by
		'  multiplying the size by the scale factor.
		
		Dim pGeoFeatureLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim pSimpleRenderer As ESRI.ArcGIS.Carto.ISimpleRenderer
		Dim pUniqueValRenderer As ESRI.ArcGIS.Carto.IUniqueValueRenderer
		Dim pClassBreaksRenderer As ESRI.ArcGIS.Carto.IClassBreaksRenderer
		Dim pMarkerSym As ESRI.ArcGIS.Display.IMarkerSymbol
		Dim ClsIndex As Short
		
447:
		pGeoFeatureLayer = FLayer
448:
		If TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.ISimpleRenderer Then
449:
			pSimpleRenderer = pGeoFeatureLayer.Renderer
			'
			'  Simple renderer has just one symbol.
			'
453:
			pMarkerSym = pSimpleRenderer.Symbol
454:
			Call ScaleAnyMarker(pMarkerSym, Multiplier)
455:
		ElseIf TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
456:
			pUniqueValRenderer = pGeoFeatureLayer.Renderer
			'
			'  Unique value renderer has <ValueCount> symbols.
			'
460:
			For ClsIndex = 0 To pUniqueValRenderer.ValueCount - 1
				
461:
				pMarkerSym = pUniqueValRenderer.Symbol(pUniqueValRenderer.Value(ClsIndex))
				
462:
				Call ScaleAnyMarker(pMarkerSym, Multiplier)
				
463:
			Next
464:
		ElseIf TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.IClassBreaksRenderer Then
465:
			pClassBreaksRenderer = pGeoFeatureLayer.Renderer
			'
			'  Class breaks renderer has <BreakCount> symbols
			'
469:
			For ClsIndex = 0 To pClassBreaksRenderer.BreakCount
470:
				pMarkerSym = pClassBreaksRenderer.Symbol(ClsIndex)
471:
				Call ScaleAnyMarker(pMarkerSym, Multiplier)
472:
			Next
473:
		End If

		Exit Sub
ErrorHandler:
		HandleError(False, "IncrementPointMarkers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'FLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub IncrementLineMarkers(ByRef FLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler

		''MsgBox "incrementing line decorations for " & FLayer.Name

		Exit Sub
ErrorHandler:
		HandleError(False, "IncrementLineMarkers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'FLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'Multiplier' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub IncrementLineWidths(ByRef FLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef Multiplier As Double)
		On Error GoTo ErrorHandler
		
		'  Change the line width for each class in the specified layer by
		'  the number of points specified.  Negative points will decrease the size.
		
		Dim pGeoFeatureLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim pSimpleRenderer As ESRI.ArcGIS.Carto.ISimpleRenderer
		Dim pUniqueValRenderer As ESRI.ArcGIS.Carto.IUniqueValueRenderer
		Dim pClassBreaksRenderer As ESRI.ArcGIS.Carto.IClassBreaksRenderer
		Dim pLineSym As ESRI.ArcGIS.Display.ILineSymbol
		Dim ClsIndex As Short
		
508:
		pGeoFeatureLayer = FLayer
509:
		If TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.ISimpleRenderer Then
510:
			pSimpleRenderer = pGeoFeatureLayer.Renderer
			'
			'  Simple renderer has just one symbol.
			'
514:
			pLineSym = pSimpleRenderer.Symbol
515:
			Call ScaleAnyLine(pLineSym, Multiplier)
516:
		ElseIf TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.IUniqueValueRenderer Then
517:
			pUniqueValRenderer = pGeoFeatureLayer.Renderer
			'
			'  Unique value renderer has <ValueCount> symbols.
			'
			'   For some reason, this blows up on symbol number (valuecount - 1),
			'   so as a workaround I'm only incrementing up to ValueCount - 2.
			'   Not sure why this would be the case.  It may have something to
			'   do with whether the "All other values" category is displayed. (ES)
			'
526:
			For ClsIndex = 0 To pUniqueValRenderer.ValueCount - 2
527:
				pLineSym = pUniqueValRenderer.Symbol(pUniqueValRenderer.Value(ClsIndex))
528:
				Call ScaleAnyLine(pLineSym, Multiplier)
529:
			Next
530:
		ElseIf TypeOf pGeoFeatureLayer.Renderer Is ESRI.ArcGIS.Carto.IClassBreaksRenderer Then
531:
			pClassBreaksRenderer = pGeoFeatureLayer.Renderer
			'
			'  Class breaks renderer has <BreakCount> symbols
			'
535:
			For ClsIndex = 0 To pClassBreaksRenderer.BreakCount
536:
				pLineSym = pClassBreaksRenderer.Symbol(ClsIndex)
537:
				Call ScaleAnyLine(pLineSym, Multiplier)
538:
			Next
539:
		End If
		
		Exit Sub
ErrorHandler:
		HandleError(False, "IncrementLineWidths " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'FLayer' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	' UPGRADE_INFO (#0551): The 'SetToBold' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub SetBoldText(ByRef FLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef SetToBold As Boolean)
		On Error GoTo ErrorHandler
		
		Dim pGeoFL As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		Dim colAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerPropertiesCollection
		Dim pAnnoLayerProps As ESRI.ArcGIS.Carto.IAnnotateLayerProperties
		Dim pLabelEngineProps As ESRI.ArcGIS.Carto.ILabelEngineLayerProperties
		Dim pTextSymbol As ESRI.ArcGIS.Display.ITextSymbol
        'Dim pFont As System.Drawing.Font
        Dim pFont As stdole.IFont
		'exit sub if anno layer
		If FLayer.Name = "Fire Annotation" Then
			Exit Sub
		End If


559:

		pGeoFL = FLayer
560:
		colAnnoLayerProps = pGeoFL.AnnotationProperties
561:
		colAnnoLayerProps.QueryItem(0, pAnnoLayerProps, Nothing, Nothing)
562:
		pLabelEngineProps = pAnnoLayerProps
563:
		pTextSymbol = pLabelEngineProps.Symbol
564:
        pFont = pTextSymbol.Font
		
566:
		FontChangeBold6(pFont, SetToBold)
567:
        '        FontChangeName6(pTextSymbol.Font, pFont.Name)
        pTextSymbol.Font = pFont

		Exit Sub
ErrorHandler:
		HandleError(False, "SetBoldText " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'UseColor' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub LoadSymbols(ByRef UseColor As Boolean)
		On Error GoTo ErrorHandler

		'  Turn symbols to black and white or color for all the FIMT layers
		'  in the current incident.
		
		' UPGRADE_INFO (#0701): The 'i' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim i As Short
		' UPGRADE_INFO (#0701): The 'pFeatureLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim LayerPath As String = ""
		Dim lyr_ex As String = ""
		
		If UseColor Then
			lyr_ex = ""
		Else
			lyr_ex = "_bw"
		End If
		
		If Environ("ARCGISHOME") = "" Then
			MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
			Exit Sub
		End If
		
		LayerPath = Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\"
		modCommon_Functions.Apply_Layer_File_Symbology(p_polyfeatlayer, LayerPath & g_FIRE_BOUNDARY & lyr_ex & ".lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(p_linefeatlayer, LayerPath & g_FIRE_LINE & lyr_ex & ".lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(p_pointfeatlayer, LayerPath & g_FIRE_POINT & lyr_ex & ".lyr")
		modCommon_Functions.Apply_Layer_File_Symbology(p_divfeatlayer, LayerPath & g_FIRE_ASSIGNMENT & lyr_ex & ".lyr")
		
		modCommon_Functions.expand_layer(p_polyfeatlayer, False)
		modCommon_Functions.expand_layer(p_linefeatlayer, False)
		modCommon_Functions.expand_layer(p_pointfeatlayer, False)
		modCommon_Functions.expand_layer(p_divfeatlayer, False)
		
		Exit Sub
ErrorHandler:
		HandleError(False, "LoadSymbols " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0701): The 'LoadLayerSymbols' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Sub LoadLayerSymbols(ByRef FeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef LayerFile As String)
		' EXCLUDED: On Error GoTo ErrorHandler

		' EXCLUDED: Dim pGeoFeatLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		' EXCLUDED: Dim pGXLayer As ESRI.ArcGIS.Catalog.IGxLayer
		' EXCLUDED: Dim pGXFile As ESRI.ArcGIS.Catalog.IGxFile
		' EXCLUDED: Dim FSO As VB6FileSystemObject
		' EXCLUDED: Dim pTempGFLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		
		' EXCLUDED: 811:
		' EXCLUDED: FSO = New VB6FileSystemObject()
		' EXCLUDED: 812:
		' EXCLUDED: If FSO.FileExists(LayerFile) Then
			' EXCLUDED: 813:
			' EXCLUDED: pGXLayer = New ESRI.ArcGIS.Catalog.GxLayer()
			' EXCLUDED: 814:
			' EXCLUDED: pGXFile = pGXLayer
			' EXCLUDED: 815:
			' EXCLUDED: pGXFile.Path = LayerFile
			' EXCLUDED: 816:
			' EXCLUDED: pTempGFLayer = pGXLayer.Layer
			' EXCLUDED: 817:
			' EXCLUDED: If pTempGFLayer.Valid Then
				' EXCLUDED: 818:
				' EXCLUDED: pGeoFeatLayer = FeatLayer
				' EXCLUDED: 819:
				' EXCLUDED: pGeoFeatLayer.Renderer = pTempGFLayer.Renderer
				' EXCLUDED: 820:
				' EXCLUDED: pGeoFeatLayer.AnnotationProperties = pTempGFLayer.AnnotationProperties
				' EXCLUDED: 821:
				' EXCLUDED: pGeoFeatLayer.DisplayAnnotation = pGeoFeatLayer.DisplayAnnotation
				' EXCLUDED: 822:
			' EXCLUDED: Else
				' EXCLUDED: 823:
				' EXCLUDED: MsgBox6("Unable to read symbology from layer file " & ControlChars.CrLf & LayerFile & ".", MsgBoxStyle.Critical, "Invalid Layer File")
				' EXCLUDED: 825:
			' EXCLUDED: End If
			' EXCLUDED: 826:
		' EXCLUDED: Else
			' EXCLUDED: 827:
			' EXCLUDED: MsgBox6("The layer file " & LayerFile & ControlChars.CrLf & "was not found. ", MsgBoxStyle.Critical, "Missing Layer File")
			' EXCLUDED: 829:
		' EXCLUDED: End If
		
		' EXCLUDED: Exit Sub
		' EXCLUDED: ErrorHandler:
		' EXCLUDED: HandleError(False, "LoadLayerSymbols " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	' EXCLUDED: End Sub
	
	Private Sub EnableButtons()
		On Error GoTo ErrorHandler

		Dim bEnable As Boolean
840:
		bEnable = (lstLayerNames.SelCount > 0)
		
842:
		cmdIncreaseMarker.Enabled = bEnable
843:
		cmdDecreaseMarker.Enabled = bEnable
844:
		CmdIncreaseLine.Enabled = bEnable
845:
		cmdDecreaseLine.Enabled = bEnable
846:
		cmdIncreaseLabel.Enabled = bEnable
847:
		cmdDecreaseLabel.Enabled = bEnable
848:
		cmdBoldText.Enabled = bEnable
849:
		cmdNormalText.Enabled = bEnable
		
		Exit Sub
ErrorHandler:
		HandleError(False, "EnableButtons " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	Private Sub Form_Load() Handles MyBase.Load
		
		' UPGRADE_INFO (#0701): The 'pFeatLayer' member has been removed because it isn't used anywhere in current application.
		' EXCLUDED: Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim count As Short
		Dim complayercount As Short
		Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
		pMap = thisdocument.FocusMap
		
		'Loop through all of the maps layers to find the desired one
		For count = 0 To pMap.LayerCount - 1
			If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
				pcomplayer = pMap.Layer(count)
				For complayercount = 0 To pcomplayer.Count - 1
					If pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_BOUNDARY_NAME Then
						p_polyfeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_LINE_NAME Then
						p_linefeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_POINT_NAME Then
						p_pointfeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_ASSIGNMENT_NAME Then
						p_divfeatlayer = pcomplayer.Layer(complayercount)
					ElseIf pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_ANNO_NAME Then
						p_annofeatlayer = pcomplayer.Layer(complayercount)
					End If
				Next
			End If
		Next
		
23:
		With lstLayerNames
24:
			.Clear()
29:
			.AddItem(g_FIRE_BOUNDARY_NAME)
			.ItemData(.NewIndex) = 0
31:
			.Selected(.NewIndex) = True
			
			.AddItem(g_FIRE_LINE_NAME)
			.ItemData(.NewIndex) = 1
32:
			.Selected(.NewIndex) = True
			
			.AddItem(g_FIRE_POINT_NAME)
			.ItemData(.NewIndex) = 2
33:
			.Selected(.NewIndex) = True
			
			.AddItem(g_FIRE_ASSIGNMENT_NAME)
			.ItemData(.NewIndex) = 3
34:
			.Selected(.NewIndex) = True
			
			.AddItem(g_FIRE_ANNO_NAME)
			.ItemData(.NewIndex) = 4
35:
			.Selected(.NewIndex) = True
40:
		End With
		
42:
		'lstLayerNames.ListIndex = -1
43:
		Call EnableButtons()
	End Sub

	Private Sub lstLayerNames_Click() Handles lstLayerNames.Click
		On Error GoTo ErrorHandler
		
863:
		Call EnableButtons()
		
		Exit Sub
ErrorHandler:
		HandleError(True, "lstLayerNames_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0551): The 'm_pApp' parameter is neither assigned in current method nor is passed to methods that modify it. Consider changing its declaration using the ByVal keyword.
	Private Sub RefreshMap(ByRef m_pApp As ESRI.ArcGIS.Framework.IApplication)
		On Error GoTo ErrorHandler

		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
875:
		pMxDoc = m_pApp.Document
876:
		pMxDoc.ActiveView.Refresh()
877:
		pMxDoc.UpdateContents()
878:
		SetDocDirty()

		Exit Sub
ErrorHandler:
		HandleError(False, "RefreshMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	End Sub

	' UPGRADE_INFO (#0701): The 'IncrementInterval' member has been removed because it isn't used anywhere in current application.
	' EXCLUDED: Private Sub IncrementInterval(ByRef LS As ESRI.ArcGIS.Display.ILineSymbol, ByRef Increment As Short)
		' EXCLUDED: On Error GoTo ErrorHandler
		
		' EXCLUDED: Dim pLineProps As ESRI.ArcGIS.Display.ILineProperties
		' EXCLUDED: Dim pTemplate As ESRI.ArcGIS.Display.ITemplate
		
		' EXCLUDED: 893:
		' EXCLUDED: If TypeOf LS Is ESRI.ArcGIS.Display.ILineProperties Then
			' EXCLUDED: 894:
			' EXCLUDED: pLineProps = LS
			' EXCLUDED: 895:
			' EXCLUDED: pTemplate = pLineProps.Template
			' EXCLUDED: 896:
			' EXCLUDED: If (pTemplate.Interval + Increment) > 0 Then
				' EXCLUDED: 897:
				' EXCLUDED: pTemplate.Interval += Increment
				' EXCLUDED: 898:
			' EXCLUDED: End If
			' EXCLUDED: 899:
		' EXCLUDED: End If
		
		' EXCLUDED: Exit Sub
		' EXCLUDED: ErrorHandler:
		' EXCLUDED: HandleError(False, "IncrementInterval " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1)
	' EXCLUDED: End Sub

	Private Sub SetDocDirty()
		On Error GoTo ErrorHandler

		Dim pApp As ESRI.ArcGIS.Framework.IApplication
		Dim pDocDirty As ESRI.ArcGIS.Framework.IDocumentDirty
		
913:
		pApp = m_pApp
914:
		pDocDirty = pApp.Document
915:
		pDocDirty.SetDirty()

		Exit Sub
ErrorHandler:
		HandleError(False, "SetDocDirty " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4)
	End Sub

End Class
