Public Class frmBoundaryFill
    Private p_polyfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private p_linefeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private p_pointfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private p_divfeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private p_annofeatlayer As ESRI.ArcGIS.Carto.IFeatureLayer

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If rdoFill.Checked Then
            LoadSymbols(True)
            RefreshMap(m_pApp)
        ElseIf rdoClear.Checked Then
            LoadSymbols(False)
            RefreshMap(m_pApp)
        End If

        Me.Close()
    End Sub

    Private Sub LoadSymbols(ByRef UseColor As Boolean)
        Dim LayerPath As String = ""
        Dim lyr_ex As String = ""

        If UseColor Then
            lyr_ex = ""
        Else
            lyr_ex = "_Clear"
        End If

        If Environ("ARCGISHOME") = "" Then
            MsgBox6("Some Necessary Environment Variables are missing, try restarting ArcMap", MsgBoxStyle.Exclamation, "Incorrect Configuration")
            Exit Sub
        End If

        LayerPath = Replace(Environ("ARCGISHOME"), "\ArcGIS", "") & "\DEC\GIS\ArcGIS9\Fire\"
        modCommon_Functions.Apply_Layer_File_Symbology(p_polyfeatlayer, LayerPath & g_FIRE_BOUNDARY & lyr_ex & ".lyr")
        modCommon_Functions.expand_layer(p_polyfeatlayer, False)
    End Sub
    Private Sub RefreshMap(ByRef m_pApp As ESRI.ArcGIS.Framework.IApplication)
        Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
        pMxDoc = m_pApp.Document
        pMxDoc.ActiveView.Refresh()
        pMxDoc.UpdateContents()
        SetDocDirty()
    End Sub
    Private Sub SetDocDirty()
        Dim pApp As ESRI.ArcGIS.Framework.IApplication
        Dim pDocDirty As ESRI.ArcGIS.Framework.IDocumentDirty
        pApp = m_pApp
        pDocDirty = pApp.Document
        pDocDirty.SetDirty()
    End Sub

    Private Sub frmBoundaryFill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim pMap As ESRI.ArcGIS.Carto.IMap
        Dim count As Short
        Dim complayercount As Short
        Dim pcomplayer As ESRI.ArcGIS.Carto.ICompositeLayer
        Dim thisdocument As ESRI.ArcGIS.ArcMapUI.IMxDocument = m_pApp.Document
        Dim fireBoundaryExist As Boolean
        pMap = thisdocument.FocusMap
        For count = 0 To pMap.LayerCount - 1
            If TypeOf pMap.Layer(count) Is ESRI.ArcGIS.Carto.IGroupLayer Then
                pcomplayer = pMap.Layer(count)
                For complayercount = 0 To pcomplayer.Count - 1
                    If pcomplayer.Layer(complayercount).Name = modGlobals.g_FIRE_BOUNDARY_NAME Then
                        p_polyfeatlayer = pcomplayer.Layer(complayercount)
                        fireBoundaryExist = True
                    End If
                Next

            End If

        Next
        If fireBoundaryExist = True Then
            btnOk.Enabled = True
        Else
            btnOk.Enabled = False
        End If

    End Sub
End Class