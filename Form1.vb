Imports DotSpatial.Controls
Imports DotSpatial.Data
Imports DotSpatial.Symbology
Imports DotSpatial.Topology
Imports BruTile
Imports DotSpatial.Plugins.Webmap

Public Class Form1

    Public AppPath As String = Application.ExecutablePath
    Public ResourcesPath As String = AppPath.ToUpper.Replace("\Creating_Geospatial_Information_System", "\Resources")
    Public lyrPemerintah As MapPointLayer
    Public lyrAset As MapPointLayer
    Public lyrJalan As MapLineLayer
    Public lyrAdministrasi As MapPolygonLayer
    Public iselect(,) As String
    Public iselectnumd As Integer = 0
    Public totalselected As Integer
    Public selectnext As String = "salah"
    Public fullextentclick As String = "salah"
    Public sedangload As Boolean = False
    Public pointLayerTemplate As MapPointLayer
    Public pointFeatureTemplate As New FeatureSet(FeatureType.Point)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        sedangload = True

        'ADD Layer BATAS ADMIN
        lyrAdministrasi = Map1.Layers.Add(ResourcesPath & "\Database\Spatial\Balam.shp")
        lyrAdministrasi.LegendText = "Batas Administrasi"
        lyrAdministrasi.FeatureSet.AddFid()
        lyrAdministrasi.FeatureSet.Save()
        lyrAdministrasi.SelectionEnabled = False

        'Dim symbolAdmin As New PolygonSymbolizer(Color.FromArgb(255, Color.White), Color.Transparant, o.5)
        'lyrAdmin.Symbolizer = symbolAdmin

        Dim schemeAdmin As New PolygonScheme
        schemeAdmin.EditorSettings.ClassificationType = ClassificationType.UniqueValues
        schemeAdmin.EditorSettings.UseGradient = False
        schemeAdmin.EditorSettings.FieldName = "WADMKC"
        schemeAdmin.CreateCategories(lyrAdministrasi.DataSet.DataTable)

        For Each ifc As IFeatureCategory In schemeAdmin.GetCategories
            ifc.SetColor(Color.FromArgb(255, ifc.GetColor))
        Next

        lyrAdministrasi.Symbology = schemeAdmin

        'ADD LAYER JARINGAN JALAN
        lyrJalan = Map1.Layers.Add(ResourcesPath & "\Database\Spatial\JALAN_LN_50K.shp")
        lyrJalan.LegendText = "Jaringan Jalan"
        lyrJalan.FeatureSet.AddFid()
        lyrJalan.FeatureSet.Save()
        lyrJalan.SelectionEnabled = False

        Dim schemeJalan As New LineScheme
        schemeJalan.ClearCategories()

        Dim symbolizerJalanKolektor As New LineSymbolizer(Color.FromArgb(255, 127, 0), Color.Gray, 3,
                                                            Drawing2D.DashStyle.Solid,
            Drawing2D.LineCap.Flat)
        symbolizerJalanKolektor.ScaleMode = ScaleMode.Simple
        Dim categoryJalanKolektor As New LineCategory(symbolizerJalanKolektor)
        categoryJalanKolektor.FilterExpression = "[REMARK] = 'Jalan Kolektor'"
        categoryJalanKolektor.LegendText = "Jalan Kolektor"
        schemeJalan.AddCategory(categoryJalanKolektor)

        Dim symbolizerJalanLokal As New LineSymbolizer(Color.FromArgb(178, 178, 255), Color.Gray, 2,
                                                            Drawing2D.DashStyle.Solid, Drawing2D.LineCap.Flat)
        symbolizerJalanLokal.ScaleMode = ScaleMode.Simple
        Dim categoryJalanLokal As New LineCategory(symbolizerJalanLokal)
        categoryJalanLokal.FilterExpression = "[REMARK] = 'Jalan Lokal'"
        categoryJalanLokal.LegendText = "Jalan Lokal"
        schemeJalan.AddCategory(categoryJalanLokal)

        Dim symbolizerJalanLain As New LineSymbolizer(Color.FromArgb(232, 190, 255), Color.Gray, 1.5,
                                                            Drawing2D.DashStyle.Solid, Drawing2D.LineCap.Flat)

        symbolizerJalanLain.ScaleMode = ScaleMode.Simple
        Dim categoryJalanLain As New LineCategory(symbolizerJalanLain)
        categoryJalanLain.FilterExpression = "[REMARK] = 'Jalan Lain'"
        categoryJalanLain.LegendText = "Jalan Lain"
        schemeJalan.AddCategory(categoryJalanLain)

        Dim symbolizerJalanSetapak As New LineSymbolizer(Color.FromArgb(232, 190, 255), Color.Gray, 1.5,
                                                           Drawing2D.DashStyle.Solid, Drawing2D.LineCap.Flat)
        symbolizerJalanSetapak.ScaleMode = ScaleMode.Simple
        Dim categoryJalanSetapak As New LineCategory(symbolizerJalanSetapak)
        categoryJalanSetapak.FilterExpression = "[REMARK] = 'Jalan Setapak'"
        categoryJalanSetapak.LegendText = "Jalan Setapak"
        schemeJalan.AddCategory(categoryJalanSetapak)

        For Each ifc As IFeatureCategory In schemeJalan.GetCategories
            ifc.SetColor(Color.FromArgb(255, ifc.GetColor))
        Next

        lyrJalan.Symbology = schemeJalan

        'ADD LAYER TTIK PEMERINTAHAN
        lyrPemerintah = Map1.Layers.Add(ResourcesPath & "\Database\Spatial\PEMERINTAHAN_PT_50K.shp")
        lyrPemerintah.LegendText = "Titik Pemerintahan"
        lyrPemerintah.FeatureSet.AddFid()
        lyrPemerintah.FeatureSet.Save()

        Dim schemePemerintahan As New PointScheme
        schemePemerintahan.ClearCategories()

        'Dim symbolizerKantorCamat As New PointSymbolizer (Color.White, DotSpatial.Symbology.PointShape.Hexagon, 10)
        Dim Kantor_camat As Image = Image.FromFile(ResourcesPath & "\Database\Spatial\KntorCmat.png", False)
        Dim symbolizerKantorCamat As New PointSymbolizer(Kantor_camat, 20)
        symbolizerKantorCamat.ScaleMode = ScaleMode.Simple
        Dim categoryKantorCamat As New PointCategory(symbolizerKantorCamat)
        categoryKantorCamat.FilterExpression = "[REMARK] = 'Kantor Camat'"
        categoryKantorCamat.LegendText = "Kantor Camat"
        schemePemerintahan.AddCategory(categoryKantorCamat)

        'Dim symbolizerKantorLurah As New PointSymbolizer (Color.White, DotSpatial.Symbology.PointShape.Hexagon, 10)
        Dim Kantor_Kelurahan As Image = Image.FromFile(ResourcesPath & "\Database\Spatial\KntorLrah.png", False)
        Dim symbolizerKantorLurah As New PointSymbolizer(Kantor_Kelurahan, 20)
        symbolizerKantorLurah.ScaleMode = ScaleMode.Simple
        Dim categoryKantorLurah As New PointCategory(symbolizerKantorLurah)
        categoryKantorLurah.FilterExpression = "[REMARK] = 'Kantor Lurah'"
        categoryKantorLurah.LegendText = "Kantor Lurah"
        schemePemerintahan.AddCategory(categoryKantorLurah)

        'Dim symbolizerKantorWalikota As New PointSymbolizer (Color.White, DotSpatial.Symbology.PointShape.Hexagon, 10)
        Dim Kantor_Walikota As Image = Image.FromFile(ResourcesPath & "\Database\Spatial\KntorWlkt.png", False)
        Dim symbolizerKantorWalikota As New PointSymbolizer(Kantor_Walikota, 20)
        symbolizerKantorWalikota.ScaleMode = ScaleMode.Simple
        Dim categoryKantorWalikota As New PointCategory(symbolizerKantorWalikota)
        categoryKantorWalikota.FilterExpression = "[REMARK] = 'Kantor Wali Kota'"
        categoryKantorWalikota.LegendText = "Kantor Wali Kota"
        schemePemerintahan.AddCategory(categoryKantorWalikota)

        lyrPemerintah.Symbology = schemePemerintahan

        'ADD LAYER TEMPLATE
        pointLayerTemplate = Map1.Layers.Add(pointFeatureTemplate)
        Dim pointttsymbol As New PointSymbolizer(Color.FromArgb(175, 75, 230, 0), DotSpatial.Symbology.PointShape.Star, 12)
        pointLayerTemplate.Symbolizer = pointttsymbol
        pointLayerTemplate.LegendText = "Point Template"
        pointLayerTemplate.LegendItemVisible = False

        'LOAD ATTRIBUTE
        Dim dt As DataTable
        dt = lyrPemerintah.DataSet.DataTable
        DataGridView1.DataSource = dt


        'LOAD DATA QUERY
        lyrAdministrasi.SelectAll()

        Dim ls1 As List(Of IFeature) = New List(Of IFeature)
        Dim il1 As ISelection = lyrAdministrasi.Selection

        ls1 = il1.ToFeatureList

        KryptonRibbonGroupComboBox_Kecamatan.Items.Clear()
        Dim i As Integer = 0
        Do While (i < il1.Count)
            Dim Name As String = (ls1(i).DataRow.ItemArray.GetValue(5).ToString)
            KryptonRibbonGroupComboBox_Kecamatan.Items.Insert(i, Name)
            i = (i + 1)
        Loop

        KryptonRibbonGroupComboBox_Kecamatan.Sorted = True
        Dim cboNumber As Integer = KryptonRibbonGroupComboBox_Kecamatan.Items.Count - 1

        Try
            For j = 1 To cboNumber
                If j > (KryptonRibbonGroupComboBox_Kecamatan.Items.Count - 1) Then Exit For
                If KryptonRibbonGroupComboBox_Kecamatan.Items(j) = KryptonRibbonGroupComboBox_Kecamatan.Items(j - 1) Then
                    KryptonRibbonGroupComboBox_Kecamatan.Items.RemoveAt(j)
                    j = j - 1
                    cboNumber = cboNumber - 1
                End If
            Next
        Catch ex As Exception

        End Try

        KryptonRibbonGroupComboBox_Kecamatan.Sorted = True

        lyrAdministrasi.UnSelectAll()


        sedangload = False

    End Sub

    Private Sub Map1_Mouse(sender As Object, e As MouseEventArgs) Handles Map1.MouseMove
        Try
            Dim coord As Coordinate = Map1.PixelToProj(e.Location)
            lblXY.Text = String.Format(“X: {0} Y: {1}", coord.X, coord.Y)
        Catch ex As Exception
            MsgBox("Mohon maaf, terjadi kesalahan. " & ex.ToString & ". Error Number (" & Err.Number & ")
                : " & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Sistem Informasi Pemerintahan Kota Bandarlampung")
        End Try
    End Sub

    Private Sub Map1_SelectionChanged(sender As Object, e As EventArgs) Handles Map1.SelectionChanged
        Try
            If sedangload = True Then Exit Sub
            If KryptonRibbonGroupButton_Identify.Checked = True Then
                If lyrPemerintah.Selection.Count = 0 Then
                    FormPopUp.Map1.ClearLayers()
                    Call RemoveSelection()
                    Exit Sub
                Else
                    FormPopUp.Show()
                    Call ShowPhoto()
                    FormPopUp.BringToFront()
                    FormPopUp.Activate()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Map1_MouseUp(sender As Object, e As MouseEventArgs) Handles Map1.MouseUp
        If KryptonRibbonGroupButton_AddPoint.Checked = True And Map1.Cursor = Cursors.Cross Then
            If formAddPoint.RadioButton_Kursor.Checked = True Then
                If e.Button = MouseButtons.Left Then
                    sedangload = True
                    pointLayerTemplate.SelectAll()
                    pointLayerTemplate.RemoveSelectedFeatures()
                    Dim coord As Coordinate = Map1.PixelToProj(e.Location)

                    Dim point As New DotSpatial.Topology.Point(coord)
                    Dim currentFeature As IFeature = pointFeatureTemplate.AddFeature(point)
                    pointFeatureTemplate.AddFeature(point)
                    formAddPoint.txtTitikX.Text = coord.X
                    formAddPoint.txtTitikY.Text = coord.Y
                    sedangload = False
                End If
                pointFeatureTemplate.InitializeVertices()
                pointLayerTemplate.DataSet.InitializeVertices()
                pointLayerTemplate.AssignFastDrawnStates()
                pointFeatureTemplate.UpdateExtent()
                pointLayerTemplate.DataSet.UpdateExtent()
                Map1.Refresh()
                Map1.ResetBuffer()
            End If
        End If

    End Sub

    Private Sub KryptonRibbonGroupButton_Normal_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Normal.Click
        If KryptonRibbonGroupButton_Normal.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
            'KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If

    End Sub

    Private Sub kryptonRibbonGroupButton_Zoomin_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_ZoomIn.Click
        If KryptonRibbonGroupButton_ZoomIn.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.ZoomIn
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If

    End Sub

    Private Sub KryptonRibbonGroupButton_Zoomout_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_ZoomOut.Click
        If KryptonRibbonGroupButton_ZoomOut.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.ZoomOut
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If

    End Sub

    Private Sub KryptonRibbonGroupButton_Pan_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Pan.Click
        If KryptonRibbonGroupButton_Pan.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.Pan
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If

    End Sub


    Private Sub KryptonRibbonGroupButton_Identify_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Identify.Click
        If KryptonRibbonGroupButton_Identify.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.Select
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If

    End Sub

    Private Sub KryptonRibbonGroupButtonLength_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Length.Click
        If KryptonRibbonGroupButton_Length.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.Label
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If
    End Sub

    Private Sub KryptonRibbonGroupButtonArea_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Area.Click
        If KryptonRibbonGroupButton_Area.Checked = True Then
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.Label
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_AddPoint.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If
    End Sub

    Private Sub KryptonRibbonGroupButton_addpoint_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_AddPoint.Click
        Map1.Cursor = Cursors.Cross
        formAddPoint.Show()
        formAddPoint.BringToFront()
        formAddPoint.Activate()

        If KryptonRibbonGroupButton_AddPoint.Checked = True Then
            Map1.Cursor = Cursors.Cross
            formAddPoint.Show()
            formAddPoint.BringToFront()
            formAddPoint.Activate()
            KryptonRibbonGroupButton_Normal.Checked = False
            KryptonRibbonGroupButton_ZoomIn.Checked = False
            KryptonRibbonGroupButton_ZoomOut.Checked = False
            KryptonRibbonGroupButton_Pan.Checked = False
            KryptonRibbonGroupButton_Length.Checked = False
            KryptonRibbonGroupButton_Area.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
            KryptonRibbonGroupButton_Identify.Checked = False
        Else
            KryptonRibbonGroupButton_Normal.Checked = True
            Map1.FunctionMode = DotSpatial.Controls.FunctionMode.None
        End If
    End Sub

    Private Sub KryptonRibbonGroupButtonZoomIn_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_ZoomIn.Click
        Map1.ZoomIn()
    End Sub

    Private Sub KryptonRibbonGroupButtonZoomOut_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_ZoomOut.Click
        Map1.ZoomOut()
    End Sub

    Private Sub KryptonRibbonGroupButtonPrev_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Prev.Click
        Map1.ZoomToPrevious()
    End Sub

    Private Sub KryptonRibbonGroupButtonNext_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_Next.Click
        Map1.ZoomToNext()
    End Sub
    Private Sub KryptonRibbonGroupButtonFullExtent_Click(sender As Object, e As EventArgs) Handles KryptonRibbonGroupButton_FullExtent.Click
        Map1.ZoomToMaxExtent()
    End Sub


    Private Sub KryptonRibbonGroupComboBoxQueryKecamatan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles KryptonRibbonGroupComboBox_Kecamatan.SelectedIndexChanged
        If KryptonRibbonGroupComboBox_Kecamatan.Text = "Cari Kecamatan..." Then Exit Sub

        sedangload = True

        Dim StrKecamatan As String = KryptonRibbonGroupComboBox_Kecamatan.Text
        lyrAdministrasi.SelectByAttribute("[WADMKC] ='" & StrKecamatan & "'")
        lyrAdministrasi.ZoomToSelectedFeatures(0.01, 0.01)
        Map1.Refresh()

        Dim ls1 As List(Of IFeature) = New List(Of IFeature)
        Dim il1 As ISelection = lyrAdministrasi.Selection

        ls1 = il1.ToFeatureList

        KryptonRibbonGroupComboBox_Desa.Items.Clear()
        Dim i As Integer = 0
        Do While (i < il1.Count)
            Dim Name As String = (ls1(i).DataRow.ItemArray.GetValue(4).ToString)
            KryptonRibbonGroupComboBox_Desa.Items.Insert(i, Name)
            i = (i + 1)
        Loop

        KryptonRibbonGroupComboBox_Desa.Sorted = True
        Dim cboNumber As Integer = KryptonRibbonGroupComboBox_Desa.Items.Count - 1
        Try
            For j = 1 To cboNumber
                If j > (KryptonRibbonGroupComboBox_Desa.Items.Count - 1) Then Exit For
                If KryptonRibbonGroupComboBox_Desa.Items(j) = KryptonRibbonGroupComboBox_Desa.Items(j - 1) Then
                    j = j - 1
                    cboNumber = cboNumber - 1
                End If
            Next
        Catch ex As Exception
        End Try

        KryptonRibbonGroupComboBox_Desa.Sorted = True

        'lyrAdministrasi.UnSelectAll()

        sedangload = False

    End Sub

    Private Sub KryptonRibbonGroupComboBox_Desa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles KryptonRibbonGroupComboBox_Desa.SelectedIndexChanged
        If KryptonRibbonGroupComboBox_Kecamatan.Text = "Cari Kecamatan..." Then Exit Sub
        If KryptonRibbonGroupComboBox_Kecamatan.Text = "Cari Desa..." Then Exit Sub

        sedangload = True

        Dim StrKecamatan As String = KryptonRibbonGroupComboBox_Kecamatan.Text
        Dim StrDesa As String = KryptonRibbonGroupComboBox_Desa.Text
        lyrAdministrasi.SelectByAttribute("[WADMKC] ='" & StrKecamatan & "' AND [NAMOBJ] ='" & StrDesa & " ' ")
        lyrAdministrasi.ZoomToSelectedFeatures(0.01, 0.01)
        Map1.Refresh()

        lyrPemerintah.SelectByAttribute("[WADMKC] ='" & StrKecamatan & "' AND [WADMKC] = '" & StrDesa & "'")
        Dim ls1 As List(Of IFeature) = New List(Of IFeature)
        Dim il1 As ISelection = lyrAset.Selection

        ls1 = il1.ToFeatureList

        KryptonRibbonGroupComboBox_Desa.Items.Clear()
        Dim i As Integer = 0
        Do While (i < il1.Count)
            Dim Name As String = (ls1(i).DataRow.ItemArray.GetValue(5).ToString)
            KryptonRibbonGroupComboBox_Desa.Items.Insert(i, Name)
            i = (i - 1)
        Loop

        KryptonRibbonGroupComboBox_Aset.Sorted = True
        Dim cboNumber As Integer = KryptonRibbonGroupComboBox_Desa.Items.Count - 1
        Try
            For j = 1 To cboNumber
                If j > (KryptonRibbonGroupComboBox_Desa.Items.Count - 1) Then Exit For
                If KryptonRibbonGroupComboBox_Desa.Items(j) = KryptonRibbonGroupComboBox_Desa.Items(j - 1) Then
                    KryptonRibbonGroupComboBox_Desa.Items.RemoveAt(j)
                    j = j - 1
                    cboNumber = cboNumber - 1
                End If
            Next
        Catch ex As Exception

        End Try

        KryptonRibbonGroupComboBox_Desa.Sorted = True

        'lyrAdmin UnSelectAll()

        sedangload = False

    End Sub

    Private Sub KryptonRibbonGroupComboBox_Aset_SelectedIndexChanged(sender As Object, e As EventArgs) Handles KryptonRibbonGroupComboBox_Aset.SelectedIndexChanged
        If KryptonRibbonGroupComboBox_Kecamatan.Text = "Cari Kecamatan..." Then Exit Sub
        If KryptonRibbonGroupComboBox_Desa.Text = "Cari Desa..." Then Exit Sub
        If KryptonRibbonGroupComboBox_Aset.Text = "Cari Aset..." Then Exit Sub

        sedangload = True

        Dim StrKecamatan As String = KryptonRibbonGroupComboBox_Kecamatan.Text
        Dim StrDesa As String = KryptonRibbonGroupComboBox_Desa.Text
        Dim StrAset As String = KryptonRibbonGroupComboBox_Aset.Text
        lyrPemerintah.SelectByAttribute("[WADMKC] '" & StrKecamatan & " ' AND [NAMOBJ] ='" & StrDesa & " ' And [Penggunaan] = '" & StrAset & "'")
        lyrPemerintah.ZoomToSelectedFeatures(0.01, 0.01)
        Map1.Refresh()

        lyrAdministrasi.UnSelectAll()

        sedangload = False
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        sedangload = True
        If DataGridView1.SelectedRows.Count = 0 Then Exit Sub
        Map1.ClearSelection()
        lyrPemerintah.Select(CInt(DataGridView1.SelectedRows.Item(0).Cells("FID").Value))
        lyrPemerintah.ZoomToSelectedFeatures(0.01, 0.01)
        Map1.Refresh()
        sedangload = False
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs)
        sedangload = True
        If DataGridView1.SelectedRows.Count = 0 Then Exit Sub
        Map1.ClearSelection()
        For i = 0 To DataGridView1.SelectedRows.Count - 1
            'lyrPemerintah.SelectByAttribute("[FID] =" & DataGridView1.SelectedRows.Item(i).Cells.Item("FID").Value)
            lyrPemerintah.Select(CInt(DataGridView1.SelectedRows.Item(i).Cells.Item("FID").Value))
        Next
        lyrPemerintah.ZoomToSelectedFeatures(0.01, 0.01)
        Map1.Refresh()
        sedangload = True
    End Sub

    'POP UP 
    Public Sub ShowPhoto()
        Try
            Dim ls1 As List(Of IFeature) = New List(Of IFeature)
            Dim il1 As ISelection = lyrPemerintah.Selection

            Dim dt As DataTable
            dt = lyrPemerintah.DataSet.DataTable

            Dim Kode As String = ""
            Dim NamaAset As String = ""
            Dim JenisAset As String = ""
            Dim AtasNama As String = ""
            Dim Foto As String = ""
            Dim ShapeIndex As Integer = ""
            Dim inputString As String = ""

            ls1 = il1.ToFeatureList

            Kode = (ls1(0).DataRow.ItemArray.GetValue(2).ToString)
            NamaAset = (ls1(0).DataRow.ItemArray.GetValue(7).ToString)
            JenisAset = (ls1(0).DataRow.ItemArray.GetValue(2).ToString)
            AtasNama = (ls1(0).DataRow.ItemArray.GetValue(8).ToString)
            Foto = (ls1(0).DataRow.ItemArray.GetValue(43).ToString)
            ShapeIndex = (ls1(0).DataRow.ItemArray.GetValue(dt.Columns.Count - 1))


            FormPopUp.txtKode.Text = Kode
            FormPopUp.txtNamaAset.Text = NamaAset
            FormPopUp.txtJenisAset.Text = JenisAset
            FormPopUp.txtAtasNama.Text = AtasNama
            FormPopUp.txtFoto.Text = Foto
            FormPopUp.txtShapeIndex.Text = ShapeIndex


            Dim AlamatFoto As String = ResourcesPath & "\Database\Spatial\AlmtKtBalam.png" & Foto
            FormPopUp.Map1.AddLayer(AlamatFoto)

            If NamaAset = "" Then
                Call RemoveSelection()
                Exit Sub
            End If

            Map1.Refresh()
            Me.Refresh()
            FormPopUp.Refresh()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub RemoveSelection()
        Try
            sedangload = True

            FormPopUp.txtKode.Text = ""
            FormPopUp.txtNamaAset.Text = ""
            FormPopUp.txtJenisAset.Text = ""
            FormPopUp.txtAtasNama.Text = ""
            FormPopUp.txtFoto.Text = ""
            FormPopUp.txtFoto.Text = ""
            FormPopUp.txtShapeIndex.Text = ""

            lyrAdministrasi.UnSelectAll()
            lyrPemerintah.UnSelectAll()

            FormPopUp.Map1.ClearLayers()

            Me.Refresh()
            FormPopUp.Refresh()

            sedangload = False

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub TabPage_Map_Click(sender As Object, e As EventArgs)

    End Sub

    Friend Shared Function ResourcePath() As String
        Throw New NotImplementedException()
    End Function

    Private Sub FormMainWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class