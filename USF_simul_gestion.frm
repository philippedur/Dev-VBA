VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_simul_gestion 
   Caption         =   "simul_recette_vs_stock"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18180
   OleObjectBlob   =   "USF_simul_gestion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_simul_gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public WithEvents Groupelbl As MSForms.TextBox    'USF_simul_gestion         !!!!!!!!!!!!!!!
Attribute Groupelbl.VB_VarHelpID = -1
Dim ctrl As Control
Dim sub_stock As Double
Dim lbl() As New Classe_USF_simul_gestion
Private Sub Validation_STOCK_Click()
'    Call update_stock_for_all_products_into_recette(nb_rec)
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    For Each chrt In Workbooks("Facturation-auto-mail-MIDI-SERVICES-01.xlsm").Sheets("Gestion").ChartObjects
        chrt.Delete
    Next
    Unload USF_simul_gestion
    Set USF_simul_gestion = Nothing
    '    On Error Resume Next
End Sub
Public Sub Userform_initialize()
    Call init_rep2
    '    SpinButton1.max = 100
    '    SpinButton1.Min = -10
    '    SpinButton1.SmallChange = 1
    '    nb_rec = 1
    ''    simul_gestion.Label80.Caption = "     DASHBOARD  MIDI-SERVICES  " & Year(Date)
'    TextBox46.text = "nb recettes à déduire du stock: " & nb_rec & " X"
    '    Call refresh(nb_rec)
'''        Call Calc_pmt_clients_graph_charge_table
    Call Calc_pmt_clients_graph_cmpt_table
    Call Graph_Tab_init(Tabmin1(), tabmin2())
'        USF_simul_gestion.Show
End Sub
Private Sub Write_V()
    sstr1 = t1
    For L = 1 To Len(t1)
        sstr2 = Left(sstr1, L) & vbCrLf
    Next
End Sub
Public Sub refresh(nb_rec)
    Set c3 = Sheets("CLIENTS")
    '    tolerance = 0.05
    nbrowmax = c3.Range("A65000").End(xlUp).Row
    '        Call Calc_pmt_clients_graph_charge_table
End Sub
Sub effacecharts()
    For Each chrt In Sheets("Gestion").ChartObjects
        chrt.Delete
    Next
End Sub
Public Sub Graph_Tab_init(Tabmin1(), tabmin2())
    Dim Grph As Chart
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Gestion"
    With ActiveChart
    ''''    .ApplyChartTemplate ("C:\Users\R\AppData\Roaming\Microsoft\Templates\Charts\Graphique3_modele.crtx")
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With ActiveChart.Axes(xlValue)
            .MinimumScale = -100000
            .MaximumScale = 50000    ' max     ' Val(Workbooks("stock.xlsm").Worksheets("ProduitsStock").Cells.Range("D" & pos).Value)
        End With
        j = 0
        For i = 2 To Month(Date) + 1
            j = j + 1
            .SeriesCollection.NewSeries
            .SeriesCollection(j).Values = IIf(tabmin2(1, i) <> "", tabmin2(1, i), "")
            .ChartType = xlColumnClustered
            .SeriesCollection(j).Name = IIf(Tabmin1(1, i) <> "", Tabmin1(1, i - 1), "")
            .SeriesCollection(j).HasDataLabels = True
            .SeriesCollection(j).DataLabels.ShowValue = True
            .SeriesCollection(j).Format.Fill.ForeColor.RGB = GetLongFromRGB(64, 192, 224)
    '            .SeriesCollection(j).DataLabels.NumberFormat = GetFormat(Split(.Formula, ",")(2))
    '            If max < tabmin(j) Then max = tabmin(1, i)
            j = j + 1
            .SeriesCollection.NewSeries
            .SeriesCollection(j).Values = IIf(tabmin2(2, i) <> "", tabmin2(2, i), "")
            .ChartType = xlColumnClustered
            .SeriesCollection(j).Name = IIf(Tabmin1(2, i) <> "", Tabmin1(2, i - 1), "")
            .SeriesCollection(j).HasDataLabels = True
            .SeriesCollection(j).DataLabels.ShowValue = True
            .SeriesCollection(j).Format.Fill.ForeColor.RGB = GetLongFromRGB(102, 153, 255)
        Next i
        With ActiveChart.Axes(xlValue)
    '            .Visible = msoTrue
    '            .UserPicture path2 & "\Images\20190803170330_215A2613.JPG"
    '            .UserPicture Path2 & "Images\ustensile2.JPG"
    '        End With
            .MinimumScale = -100000
            .MaximumScale = 100000     ' Val(Workbooks("stock.xlsm").Worksheets("ProduitsStock").Cells.Range("D" & pos).Value)
        End With
        With ActiveChart
            .HasTitle = True
            .ChartTitle.text = "  DASHBOARD  MIDI-SERVICES  " & Year(Date)
        End With
        With ActiveChart.ChartTitle.Font
            .Name = "Arial"
            .FontStyle = "BOLD"
            .Size = 14
            .color = GetLongFromRGB(41, 0, 204)    ' 1
        End With
    End With
    With ActiveChart.Parent
        .Height = 250
        .Width = i * 85    ' 90
    End With
    Set Grph = ActiveChart    ' ActiveSheet.ChartObjects(1).Chart
    ActiveChart.Export fileName:=Path2 & "Images\Gestion_MIDI-services.gif", FilterName:="GIF"
    USF_simul_gestion.Picture = LoadPicture(Path2 & "Images\Gestion_MIDI-services.gif")
    Call taillimage
    USF_simul_gestion.Repaint
End Sub
Private Sub taillimage()
    Dim centimetres As Double, pouces As Double, points As Double
    Dim pict_height, pict_width
    centimetres = USF_simul_gestion.Picture.Height / 1000
    pouces = centimetres / 2.54
    points = pouces * 72
    pict_height = points
    centimetres = USF_simul_gestion.Picture.Width / 1000
    pouces = centimetres / 2.54
    points = pouces * 72
    pict_width = points
    USF_simul_gestion.Width = pict_width + (i * 0.85)
End Sub
Public Sub Calc_pmt_clients_graph_charge_table()    '____TEST____TEST____TEST____TEST____TEST____TEST
    Application.ScreenUpdating = True
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c8 = Sheets("Gestion")
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    L = 1
    mois_fact = Array("", "JANVIER", "JANVIER", "FEVRIER", "FEVRIER", "MARS", "MARS", "AVRIL", "AVRIL", "MAI", "MAI", "JUIN", "JUIN", "JUILLET", "JUILLET", "AOUT", "AOUT", "SEPTEMBRE", "SEPTEMBRE", "OCTOBRE", "OCTOBRE", "NOVEMBRE", "NOVEMBRE", "DECEMBRE", "DECEMBRE")
    For k = 1 To Val(Month(Date)) * 2
        c8.Cells(1, k + 1) = mois_fact(k) & "/MIDI"
        c8.Cells(1, k + 2) = mois_fact(k + 1) & "/EBP"
        For i = 2 To nbrowmax
            res1 = 0
            res2 = 0
            res3 = 0
            res1 = L * c3.Range("S" & i)
            res2 = calc_trav_For_all_Year(k, i)
            res3 = c3.Range("K" & i)
            c8.Cells(i, k + 1) = res1 + res2
            c8.Cells(i, k + 2) = res3   ' res1 + res2
            c8.Range("A" & i) = c3.Range("N" & i)
        Next i
        L = L + 1
        k = k + 1
    Next k
    '    nbrowmax = c7.Range("G65000").End(xlUp).Row
    '    res = 0
    '    For i = 2 To nbrowmax
    '        sstr1 = c3.Range("N" & i)
    ''        res = calc_Xtrac_Dom(i)
    '    Next i
End Sub
Public Sub Calc_pmt_clients_graph_cmpt_table()    '____TEST____TEST____TEST____TEST____TEST____TEST
    Erase Tabmin1, tabmin2
    Application.ScreenUpdating = True
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c8 = Sheets("Gestion")
    Worksheets("Gestion").Activate
    nbrowmax = c8.Range("A65000").End(xlUp).Row
    L = 1
    ReDim Tabmin1(1 To 2, 1 To 12)  ' TAB DES ENTETES MIDI/EBP
    ReDim tabmin2(1 To 2, 1 To 12)  ' TAB VALEURS MIDI/EBP
    For k = 2 To 13
        nbrowmax = c8.Cells(Rows.Count, k).End(xlUp).Row
    '    mois_fact = Array("", "JANVIER", "JANVIER", "FEVRIER", "FEVRIER", "MARS", "MARS", "AVRIL", "AVRIL", "MAI", "MAI", "JUIN", "JUIN", "JUILLET", "JUILLET", "AOUT", "AOUT", "SEPTEMBRE", "SEPTEMBRE", "OCTOBRE", "OCTOBRE", "NOVEMBRE", "NOVEMBRE", "DECEMBRE", "DECEMBRE")
        mois_fact = Array("", "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE")
        Tabmin1(1, L) = Left(mois_fact(L), 3) & "/MIDI"
        Tabmin1(2, L) = Left(mois_fact(L), 3) & "/EBP"
        c8.Cells(1, k).Value = Tabmin1(1, L)
        c8.Cells(1, k + 1).Value = Tabmin1(2, L)
        If nbrowmax > 1 Then
            res1 = 0
            res2 = 0
            res3 = 0
            For i = 2 To nbrowmax
                res1 = res1 + (L * c3.Range("S" & i))                'Total mois domiciliation
                res2 = res2 + calc_trav_For_LastMonth(k, i)    'Total travaux de l'année
                res3 = res3 + c8.Cells(i, k + 1)           'Total EBP à ce mois de l'année
                If res2 > 0 Then
    '''            c8.Cells(i, k) = res1 + res2
    '            c8.Cells(i, k).Interior.color = GetLongFromRGB(204, 255, 51)
                    c8.Range(Cells(i, 1), Cells(i, 12)).Interior.color = GetLongFromRGB(204, 255, 51)
                End If
            Next i
            L = L + 1
            k = k + 1
    '    ReDim Preserve Tabmin(1 To 2, 1 To L)
            tabmin2(1, L) = res1 + res2  '  Mois de dom + travaux
            tabmin2(2, L) = res3    '  nb_EBP
        Else
            Exit For
        End If
    Next k
End Sub




