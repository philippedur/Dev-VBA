Attribute VB_Name = "calculs_affichage"
Public Sub set_rep()
    sdate = Format(Date, "mmmm")
    init_path = path3
    If Exist_ssrep(ssrep_pdf) = False Then
        ssrep_pdf = init_path & Year(Date) & "\"
        MkDir (ssrep_pdf)
    ElseIf Exist_rep(rep_pdf) = False Then
        rep_pdf = ssrep_pdf & "\" & sdate
        MkDir (rep_pdf)
    End If
End Sub
Public Sub Deactivation_PDFCreator()
    Dim svc As Object
    Dim sQuery As String
    Dim oproc, ProcessName
'    On Error GoTo Command1_Click_Error
    Set svc = GetObject("winmgmts:root\cimv2")
    sQuery = "select * from win32_process"
    For Each oproc In svc.execquery(sQuery)
    If oproc.Name = "PDFCreator.exe" Then
      MsgBox "Desactivation du process " & oproc.Name
        oproc.Terminate
    End If
    Next
End Sub
Public Function Exist_rep(rep_pdf) As String
    sdate = Format(Date, "mmmm")
    rep_pdf = init_path & Year(Date) & "\" & sdate & "\"
    If Len(Dir(rep_pdf, vbDirectory)) > 0 Then
        Exist_rep = True
    Else
        Exist_rep = False
    End If
End Function
Public Function Exist_ssrep(ssrep_pdf) As String
    ssrep_pdf = init_path & Year(Date)
    If Len(Dir(ssrep_pdf, vbDirectory)) > 0 Then
        Exist_ssrep = True
    Else
        Exist_ssrep = False
    End If
End Function
Public Function strf15(ByVal Data As Variant)
    Dim rf15 As String * 15
        rf15 = Data
        strf15 = rf15
End Function
Public Function strf25(ByVal Data As Variant)
    Dim rf25 As String * 25
        rf25 = Data
        strf25 = rf25
End Function
Public Function strf30(ByVal Data As Variant)
    Dim rf30 As String * 30
        rf30 = Data
        strf30 = rf30
End Function
Public Function strf60(ByVal Data As Variant)
    Dim rf60 As String * 60
        rf60 = Data
        strf60 = rf60
End Function
Public Function strf90(ByVal Data As Variant)
    Dim rf90 As String * 90
        rf90 = Data
        strf90 = rf90
End Function
Function GetLongFromRGB(Red As Integer, Green As Integer, Blue As Integer) As Long
    GetLongFromRGB = RGB(Red, Green, Blue)
End Function
Public Function calc_couleur() As Long
    Dim ctnu, ctnt, calcul As Integer
    ctnu = s4.Cells(ligne, nbCol)
    ctnt = s5.Cells(ligne, nbCol)
    If (ctnu Or ctnt) = 0 Then
        calc_couleur = RGB(255, 255, 255)
        Exit Function
    Else
        calcul = Round(ctnu / ctnt * 100, 1)
        If (calcul > 80 And calcul <= 100) Then
            calc_couleur = RGB(224, 255, 32)    ' &HB0FEB2 ' 65280  'RGB(128, 160, 32)    '
        ElseIf (calcul > 60 And calcul <= 80) Then
            calc_couleur = RGB(224, 224, 32)    '&H7CFC81   ' 57344  'RGB(128, 224, 32)    '
        ElseIf (calcul > 30 And calcul <= 60) Then
            calc_couleur = RGB(224, 192, 32)    '&H48FD4C  ' 40960  'RGB(160, 255, 32)  '
        ElseIf (calcul > 20 And calcul <= 30) Then
            calc_couleur = RGB(224, 160, 32)    '&H1CFE20   ' 32768    'RGB(192, 128, 0)
        ElseIf (calcul > 10 And calcul <= 20) Then
            calc_couleur = RGB(224, 128, 32)    '&HFD8D65   ' 16639    'RGB(255, 64, 0)    '
        ElseIf (calcul > 0 And calcul <= 10) Then
            calc_couleur = RGB(224, 96, 32)    '&HFE820E    ' 2105599  'RGB(255, 32, 32)   '
        Else
            calc_couleur = RGB(224, 64, 32)    '2105599  'RGB(255, 32, 32)   '
        End If
    End If
End Function
Public Function calc_period(ByVal k As Integer, ByVal periodicite As Integer, ByVal date_creation As Date) As Boolean
    calc = Abs(Abs(Month(date_creation)) - Abs(Month(Date)))
    Select Case periodicite
    Case Is = 12
        calc_period = True
    Case Is = 6
        Select Case calc
        Case Is = 2
        calc_period = True
        Case Else
        calc_period = False
        End Select
    Case Is = 4
        Select Case calc
        Case Is = 3
        calc_period = True
        Case Else
        calc_period = False
        End Select
    Case Is = 3
        Select Case calc
        Case Is = 4
        calc_period = True
        Case Else
        calc_period = False
        End Select
    Case Is = 2
        Select Case calc
        Case Is = 6
        calc_period = True
        Case Else
        calc_period = False
        End Select
    Case Is = 1
        Select Case calc
        Case Is = 12
        calc_period = True
        Case Else
        calc_period = False
        End Select
    Case Is = 0
        calc_period = False
    Case Else
        calc_period = False
    End Select

End Function
Public Function calc_couleur2() As Long
    Dim ctnu, ctnt, calcul As Integer
'    ctnu = c2.Cells(ligne, nbCol)
'    ctnt = c2.Cells(ligne, nbCol)
    If (ctnu Or ctnt) = 0 Then
        calc_couleur2 = RGB(255, 255, 255)
        Exit Function
    Else
        calcul = Round(ctnu / ctnt * 100, 1)
        If (calcul > 80 And calcul <= 100) Then
            calc_couleur2 = RGB(224, 255, 32)    ' &HB0FEB2 ' 65280  'RGB(128, 160, 32)    '
        ElseIf (calcul > 60 And calcul <= 80) Then
            calc_couleur2 = RGB(224, 224, 32)    '&H7CFC81   ' 57344  'RGB(128, 224, 32)    '
        ElseIf (calcul > 30 And calcul <= 60) Then
            calc_couleur2 = RGB(224, 192, 32)    '&H48FD4C  ' 40960  'RGB(160, 255, 32)  '
        ElseIf (calcul > 20 And calcul <= 30) Then
            calc_couleur2 = RGB(224, 160, 32)    '&H1CFE20   ' 32768    'RGB(192, 128, 0)
        ElseIf (calcul > 10 And calcul <= 20) Then
            calc_couleur2 = RGB(224, 128, 32)    '&HFD8D65   ' 16639    'RGB(255, 64, 0)    '
        ElseIf (calcul > 0 And calcul <= 10) Then
            calc_couleur2 = RGB(224, 96, 32)    '&HFE820E    ' 2105599  'RGB(255, 32, 32)   '
        Else
            calc_couleur2 = RGB(224, 64, 32)    '2105599  'RGB(255, 32, 32)   '
        End If
    End If
End Function
Public Function affich_info_stock() As Long
'    Workbooks("Calcul prix et recettes.xlsb").Activate
    Set c7 = Sheets("Produits")
    Set plage = c7.Range("F1:F2000")
    cle_rech = c3.Cells(ActiveCell.Row, 6)
    tt = yy
    Nbre = Application.CountIf(plage, cle_rech)
    If Nbre > 0 Then
    '    ligne = Columns("A").Find(plage).row
    '        pos = nbrow
        nbrowmax = c7.Range("H65000").End(xlUp).Row
        Set plage_rech = c7.Range("H1" & ":H" & nbrowmax)
    '        Set plage_rech = c7.Range("A" & Trim(ActiveCell.row - 1) & ":A2000")
        Set c = plage_rech.Find(cle_rech, , , xlPart)
        If Not c Is Nothing Then
            With c
                fnd1 = c.Row
                Lig = c.Row
                ligne = c.Row
                pos = c.Row
    '        Set c = plage_rech.FindNext(c)
    '        fnd2 = c.row
            End With
    '            UserForm3.Show
        End If
        Workbooks("Stock.xlsm").Activate
        Set s4 = Sheets("Contenu")
        Set s5 = Sheets("Contenance")
        Dim ctnu, ctnt, calcul As Integer
        nbCol = s4.Cells(ligne, Columns.Count).End(xlToLeft).Column
        ctnu = SommeCol(s4, s4.Cells(ligne, 1), s4.Cells(ligne, nbCol))
        ctnt = SommeCol(s5, s5.Cells(ligne, 1), s5.Cells(ligne, nbCol))
        If (ctnu Or ctnt) = 0 Then
            affich_info_stock = 0
            Exit Function
        Else
            affich_info_stock = Round(ctnu / ctnt * 100, 1)
        End If
    End If
End Function
Public Function affich_info_stock_Global() As Long
'    Workbooks("Calcul prix et recettes.xlsb").Activate
    Set c7 = Sheets("Produits")
    Set plage = c7.Range("H1:H2000")
    '    cle_rech = c7.Cells(ActiveCell.row, ActiveCell.Column - 3)
    tt = yy
    Nbre = Application.CountIf(plage, cle_rech)
    If Nbre > 0 Then
    '    ligne = Columns("A").Find(plage).row
        nbrowmax = c7.Range("H65000").End(xlUp).Row
        Set plage_rech = c7.Range("H1" & ":H" & nbrowmax)
    '        Set plage_rech = c7.Range("A" & Trim(ActiveCell.row - 1) & ":A2000")
        Set c = plage_rech.Find(cle_rech, , , xlPart)
        If Not c Is Nothing Then
            With c
                fnd1 = c.Row
                Lig = c.Row
                ligne = c.Row
                pos = c.Row
    '        Set c = plage_rech.FindNext(c)
    '        fnd2 = c.row
            End With    '
    '            UserForm3.Show

            Workbooks("Stock.xlsm").Activate
            Set s4 = Sheets("Contenu")
    '        Set s5 = Sheets("Contenance")
    '        Dim ctnu, ctnt, calcul As Integer
            nbCol = s4.Cells(ligne, Columns.Count).End(xlToLeft).Column
            ctnu = SommeCol(s4, s4.Cells(ligne, 1), s4.Cells(ligne, nbCol))
            ctnt = SommeCol(s5, s5.Cells(ligne, 1), s5.Cells(ligne, nbCol))
            If (ctnu Or ctnt) = 0 Then
                affich_info_stock_Global = RGB(255, 255, 255)
                Exit Function
            Else
                affich_info_stock_Global = Round(ctnu / ctnt * 100, 1)
                If (ctnu Or ctnt) = 0 Then
                    affich_info_stock_Global = RGB(255, 255, 255)
                    Exit Function
                Else
                    calcul = Round(ctnu / ctnt * 100, 1)
                    If (calcul > 80 And calcul <= 100) Then
                        affich_info_stock_Global = RGB(224, 255, 32)    ' &HB0FEB2 ' 65280  'RGB(128, 160, 32)    '
                    ElseIf (calcul > 60 And calcul <= 80) Then
                        affich_info_stock_Global = RGB(224, 224, 32)    '&H7CFC81   ' 57344  'RGB(128, 224, 32)    '
                    ElseIf (calcul > 30 And calcul <= 60) Then
                        affich_info_stock_Global = RGB(224, 192, 32)    '&H48FD4C  ' 40960  'RGB(160, 255, 32)  '
                    ElseIf (calcul > 20 And calcul <= 30) Then
                        affich_info_stock_Global = RGB(224, 160, 32)    '&H1CFE20   ' 32768    'RGB(192, 128, 0)
                    ElseIf (calcul > 10 And calcul <= 20) Then
                        affich_info_stock_Global = RGB(224, 128, 32)    '&HFD8D65   ' 16639    'RGB(255, 64, 0)    '
                    ElseIf (calcul > 0 And calcul <= 10) Then
                        affich_info_stock_Global = RGB(224, 96, 32)    '&HFE820E    ' 2105599  'RGB(255, 32, 32)   '
                    Else
                        affich_info_stock_Global = RGB(224, 64, 32)    '2105599  'RGB(255, 32, 32)   '
                    End If
                End If
            End If
        End If
    End If
End Function
Public Function Stru(chaine, max As Integer)
    Dim MyArray(16, 16) As Integer
    MyArray = Array("Textbox1-", "Textbox2-", "Textbox3-", "Textbox4-", "Textbox5-", "Textbox6-", "Textbox7-", "Textbox8-", "Textbox9-", "Textbox10-", "Textbox11-", "Textbox12-", , "Textbox13-", "Textbox14-""Textbox15-", "Textbox16-")
    On Error Resume Next
    sstr = InStr(1, chaine, "-", vbBinaryCompare)
    Stru = Right(chaine, (Len(chaine) - sstr))
    Exit Function
    If tt > 2 Then Exit Function
End Function
Public Sub total_ph()
    t1 = 0
    t2 = 0
    t3 = 0
    t5 = 0
    t6 = 0
    t7 = 0
    t8 = 0
    t9 = 0
    t10 = 0
    n = 0
    L = 0
    Top = IIf((fnd4 = 0) And (fnd5 = 0) And (fnd6 = 0), 3, 4)
    For ligne = Top To (r4)
        With R_t4
            If (r1 < ligne) And (ligne <= r2) Then    '  TOTAL T1 BLEU
    '                If USF6_recettes_stock("Textbox8-" & ligne).Value <> "" Then
                t1 = t1 + Val(USF6_recettes_stock("Textbox8-" & ligne).Value)
    '                End If
            ElseIf (r2 < ligne) And (ligne <= r3) Then    '  TOTAL T2 JAUNE
    '                If USF6_recettes_stock("Textbox8-" & ligne).Value <> "" Then
                t2 = t2 + Val(USF6_recettes_stock("Textbox8-" & ligne).Value)
    '                End If
            ElseIf (r3 < ligne) And (ligne <= r4) Then    '  TOTAL T3 VIOLET
    '                If USF6_recettes_stock("Textbox8-" & ligne).Value <> "" Then
                t3 = t3 + Val(USF6_recettes_stock("Textbox8-" & ligne).Value)
    '                End If
            End If
        End With
    '            If USF6_recettes_stock("Textbox9-" & r4 + 1).Value <> "" Then
        t5 = t5 + Val(USF6_recettes_stock("Textbox9-" & ligne).Value)
    '            End If
    '            If USF6_recettes_stock("Textbox10-" & r4 + 1).Value <> "" Then
        t6 = t6 + Val(USF6_recettes_stock("Textbox10-" & ligne).Value)
    '            End If
    '            If USF6_recettes_stock("Textbox11-" & r4 + 1).Value <> "" Then
        t7 = t7 + Val(USF6_recettes_stock("Textbox11-" & ligne).Value)
    '            End If
    '            If USF6_recettes_stock("Textbox13-" & r4 + 1).Value <> "" Then
        t8 = t8 + Val((Replace(USF6_recettes_stock("TextBox13-" & ligne).Value, ",", ".")))
    '                USF6_recettes_stock("Textbox13-" & r4 + 1).ForeColor = RGB(47, 117, 181)
    '            End If
    Next ligne
    t9 = 0
    For ligne = (r4 + 2) To (r5)
        With R_t4
    '            If USF6_recettes_stock("Textbox13-" & r4 + 1).Value <> "" Then
            t9 = t9 + Val((Replace(USF6_recettes_stock("TextBox13-" & ligne).Value, ",", ".")))
    '            End If
        End With
    Next ligne
    L = t1 + t2 + t3
    t10 = (t9 + t8)
    If ((fnd4 = 0) And (fnd5 = 0) And (fnd6 = 0)) Then
        USF6_recettes_stock("Textbox1A-1").Value = affich_pct(1)
        USF6_recettes_stock("Textbox1B-1").Value = affich_pct(2)
        USF6_recettes_stock("Textbox1C-1").Value = affich_pct(3)
        USF6_recettes_stock("Textbox1A-1").Value = "" & vbCrLf & Format(L, "0.0") & vbCrLf & "%"
        USF6_recettes_stock("Textbox8-" & r4 + 1).Value = Format(L, "0.0") & "%"
        USF6_recettes_stock("Textbox8-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox9-" & r4 + 1).Value = Format(t5, "###0.00") & "g"
        USF6_recettes_stock("Textbox9-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox10-" & r4 + 1).Value = Format(t6, "###0.00") & "ml"
        USF6_recettes_stock("Textbox10-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox11-" & r4 + 1).Value = t7
        USF6_recettes_stock("Textbox11-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox13-" & r4 + 1).Value = Format(t8, "###0.00") & "€"  '  Total partiel recette
        USF6_recettes_stock("Textbox13-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox13-" & r5 + 1).Value = Format(t10, "###0.00") & "€"  '  Total global recette
        USF6_recettes_stock("Textbox13-" & r5 + 1).ForeColor = RGB(255, 255, 255)
        USF6_recettes_stock("Textbox13-" & r5 + 1).BackColor = RGB(47, 117, 181)
    Else
        USF6_recettes_stock("Textbox1A-1").Value = affich_pct(1)
        USF6_recettes_stock("Textbox1B-1").Value = affich_pct(2)
        USF6_recettes_stock("Textbox1C-1").Value = affich_pct(3)
        USF6_recettes_stock("Textbox8-" & r4 + 1).Value = Format(L, "0.0") & "%"
        USF6_recettes_stock("Textbox8-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox9-" & r4 + 1).Value = Format(t5, "###0.0") & "g"
        USF6_recettes_stock("Textbox9-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox10-" & r4 + 1).Value = Format(t6, "###0.0") & "ml"
        USF6_recettes_stock("Textbox10-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox11-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox11-" & r4 + 1).Value = t7
        USF6_recettes_stock("Textbox13-" & r4 + 1).Value = Format(t8, "###0.00") & "€"  '  Total partiel recette
        USF6_recettes_stock("Textbox13-" & r4 + 1).ForeColor = RGB(47, 117, 181)
        USF6_recettes_stock("Textbox13-" & r5 + 1).Value = Format(t10, "###0.00") & "€"  '  Total global recette
        USF6_recettes_stock("Textbox13-" & r5 + 1).ForeColor = RGB(255, 255, 255)
        USF6_recettes_stock("Textbox13-" & r5 + 1).BackColor = RGB(47, 117, 181)
    End If

End Sub
Function affiche_phase(ph As String)
    Dim res As Integer
    Select Case ph
    Case Is = "A"
        res = Round((height_A / 2), 1)
        If res > 0 Then
            affiche_phase = ""    'IIf(res > 2, String(res - 1, vbCrLf) & ph & vbCrLf & t1 & vbCrLf & "%", IIf(res > 1, ph & vbCrLf & "%", ph & "%"))
        Else
            affiche_phase = ""  ' IIf(res > 1, ph & vbCrLf & "%", ph & "%")
        End If
    Case Is = "B"
        res = Round((height_B / 2), 1)
        If res > 0 Then
            affiche_phase = ""  ' IIf(res > 2, String(res - 1, vbCrLf) & ph & vbCrLf & t2 & vbCrLf & "%", IIf(res = 2, ph & vbCrLf & "%", ph & "%"))
        Else
            affiche_phase = ""
        End If
    Case Is = "C"
        res = Round((height_C / 2), 1)
        If res > 0 Then
            affiche_phase = ""  'IIf(res > 2, String(res - 1, vbCrLf) & ph & vbCrLf & t3 & vbCrLf & "%", IIf(res = 2, ph & vbCrLf & "%", ph & "%"))
        End If
    Case Is = "D"
        res = Round(Abs((fnd5 - fnd2) / 2), 1)
        If res > 0 Then
            affiche_phase = ""  'IIf(res > 2, String(res - 1, vbCrLf) & ph & vbCrLf & "%", IIf(res = 2, ph & vbCrLf & "%", ph & "%"))
        Else
            affiche_phase = ""
        End If
    Case Else
        affiche_phase = ""

    End Select
End Function
Public Function affich_pct(ByRef typt As Double) As String
    If typt = 1 Then
        If Abs(fnd5 - fnd4) = 2 Then
            affich_pct = "A" & Format(t1, "0.0") & "%"
        ElseIf Abs(fnd5 - fnd4) = 3 Then
            affich_pct = "A " & vbCrLf & Format(t1, "0.0") & "%"
        ElseIf Abs(fnd5 - fnd4) > 3 Then
            affich_pct = "A " & vbCrLf & Format(t1, "0.0") & vbCrLf & "%"
        End If
    ElseIf typt = 2 Then
        If Abs(fnd6 - fnd5) = 1 Then
            affich_pct = "B" & Format(t2, "0.0") & "%"
        ElseIf Abs(fnd6 - fnd5) = 2 Then
            affich_pct = "B " & vbCrLf & Format(t2, "0.0") & "%"
        ElseIf Abs(fnd6 - fnd5) > 2 Then
            affich_pct = "B " & vbCrLf & Format(t2, "0.0") & vbCrLf & "%"
        End If
    ElseIf typt = 3 Then
        If Abs(fnd2 - fnd6) = 1 Then
            affich_pct = "C " & Format(t3, "0.0") & "%"
        ElseIf Abs(fnd2 - fnd6) = 2 Then
            affich_pct = "C" & vbCrLf & Format(t3, "0.0") & "%"
        ElseIf Abs(fnd2 - fnd6) > 2 Then
            affich_pct = "C" & vbCrLf & Format(t3, "0.0") & vbCrLf & "%"
        End If
    End If
End Function
Function NumCoulCel(c As Object)
'Application.Volatile True
'NumCoulCel = Abs(c.Interior.ColorIndex)
End Function
Public Function SommeCol(ByVal wsht As Worksheet, plage As Range, Couleur As Long) As Double
    Application.Volatile
    Dim cellule As Range
    Dim Stock As Double
    For i = 1 To nbCol
        If wsht.Cells(ligne, i).Interior.color = 16777215 And IsNumeric(wsht.Cells(ligne, i)) Then Stock = Stock + wsht.Cells(ligne, i)
    Next i
    SommeCol = Stock
End Function
Public Function FlaconMax(ByVal wsht As Worksheet, plage As Range, Couleur As Long, ligne) As Double
    Application.Volatile
    Dim cellule As Range
    For i = 1 To nbCol
        If wsht.Cells(ligne, i).Interior.color = 16777215 And IsNumeric(wsht.Cells(ligne, i) And FlaconMax < (wsht.Cells(ligne, i))) Then
            FlaconMax = wsht.Cells(ligne, i)
        End If
    Next i
    fMax = FlaconMax
End Function
Public Function decod_cle(inp As String) As String
    sstr1 = InStr(1, inp, "-", vbTextCompare)
    If sstr1 > 0 Then
        decod_cle = Trim(Left(inp, sstr1 - 2))
    End If
End Function
Public Sub clean_rec_data()
    With R_t4
        For ligne = 1 To max - 3    ' 1 + ctr1
            For col = 8 To 15
                If USF6_recettes_stock("Combobox7-" & ligne).Value = Empty Then
                    USF6_recettes_stock("Textbox" & col & "-" & ligne).Value = ""
                End If
            Next col
        Next ligne
    End With
    With R_t4
        For ligne = (r4) + 1 To (r5) + 1
            For col = 8 To 15
                If USF6_recettes_stock("Combobox7-" & ligne).Value = Empty Then
                    USF6_recettes_stock("Textbox" & col & "-" & ligne).Value = ""
                End If
            Next col
        Next ligne
    End With
End Sub
Sub tri_colB_Societe(col As Integer)
    Set c3 = Sheets("CLIENTS")
    Worksheets("CLIENTS").Activate
    Dim titre, ordretri
    Set titre = [A1:N1]
    Cells(1, col).Select
    ordretri = IIf(Cells(1, 2).Interior.colorIndex = 3, xlDescending, xlAscending)
    Cells(1, col).CurrentRegion.Sort Key1:=Cells(2, Cells(1, col).Column), Order1:=xlAscending, header:=xlYes
    c2 = IIf(Cells(1, col).Interior.colorIndex = 2, 2, 17)
    titre.Interior.colorIndex = 2
    Cells(1, col).Interior.colorIndex = 17
End Sub
Sub tri_colB_Fourniss(col As Integer)
    Workbooks("Calcul prix et recettes.xlsb").Activate
    Sheets("Fournisseurs").Select
    Dim titre, ordretri
    Set titre = [A1:I1]
    Cells(1, col).Select
    ordretri = IIf(Cells(1, 2).Interior.colorIndex = 3, xlDescending, xlAscending)
    Cells(1, col).CurrentRegion.Sort Key1:=Cells(2, Cells(1, col).Column), Order1:=xlAscending, header:=xlYes
    c2 = IIf(Cells(1, col).Interior.colorIndex = 2, 2, 17)
    titre.Interior.colorIndex = 2
    Cells(1, col).Interior.colorIndex = 17
    Sheets("Produits").Select
End Sub
Sub tri_colB_ProduitsStock(col As Integer)
    Workbooks("Stock.xlsm").Activate
    Sheets("ProduitsStock").Select
    Dim titre, ordretri
    Set titre = [A1:H1]
    Cells(1, col).Select
    ordretri = IIf(Cells(1, 2).Interior.colorIndex = 3, xlDescending, xlAscending)
    Cells(1, col).CurrentRegion.Sort Key1:=Cells(2, Cells(1, col).Column), Order1:=xlAscending, header:=xlYes
    c2 = IIf(Cells(1, col).Interior.colorIndex = 2, 2, 17)
    titre.Interior.colorIndex = 2
    Cells(1, col).Interior.colorIndex = 17
End Sub
Sub tri_col_generic(ByVal osht As Worksheet, ByRef col As Integer)
    Dim titre, ordretri
    osht.Activate
    osht.Rows.RowHeight = 18.5
    Set titre = [A1:X1]
    Cells(1, col).Select
    ordretri = IIf(Cells(1, 2).Interior.colorIndex = 3, xlDescending, xlAscending)
    Cells(1, col).CurrentRegion.Sort Key1:=Cells(2, Cells(1, col).Column), Order1:=xlAscending, header:=xlYes
    c2 = IIf(Cells(1, col).Interior.colorIndex = 2, 2, 17)
    titre.Interior.colorIndex = 2
    Cells(1, col).Interior.colorIndex = 17
End Sub
Sub gen_test4()
    Workbooks("Calcul prix et recettes.xlsb").Activate
    Dim nbcol_s, nbcol_d As Integer
    Set c1 = Sheets("Produits")
    nbrowmax = c1.Range("F65000").End(xlUp).Row
    nbcol_s = 2  '  (largeur de la region exprimee en colonnes (légendes + commentaires compris)
    nbcol_d = 1  '  (largeur de la region exprimee en colonnes (légendes + commentaires compris)
    Application.ScreenUpdating = True
    source_range = c1.Range(Cells(1, nbcol_s), Cells(nbrowmax, nbcol_s)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Worksheets("Produits").Range(source_range).Copy
    Workbooks("Stock.xlsm").Activate
    Set s1 = Sheets("ProduitsStock")
    dest_range = Range(Cells(1, nbcol_d), Cells(nbrowmax, nbcol_d)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    With dest_range
    ' copie valeurs dans range
        Worksheets("ProduitsStock").Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                                                                   xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
    Workbooks("Calcul prix et recettes.xlsb").Activate
    Set c1 = Sheets("Produits")
    nbrowmax = c1.Range("F65000").End(xlUp).Row
    nbcol_s = 6  '  (largeur de la region exprimee en colonnes (légendes + commentaires compris)
    nbcol_d = 2  '  (largeur de la region exprimee en colonnes (légendes + commentaires compris)
    Application.ScreenUpdating = True
    source_range = c1.Range(Cells(1, nbcol_s), Cells(nbrowmax, nbcol_s)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Worksheets("Produits").Range(source_range).Copy
    Workbooks("Stock.xlsm").Activate
    Set s1 = Sheets("ProduitsStock")
    dest_range = Range(Cells(1, nbcol_d), Cells(nbrowmax, nbcol_d)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    With dest_range
    ' copie valeurs dans range
        Worksheets("ProduitsStock").Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                                                                   xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
End Sub
Function EstForm(c As Range)
    Application.Volatile
    EstForm = c.HasFormula
    '    Debug.Print "Formule: " & c.Formula & " *  " & ligne
End Function
Public Function decod_cle4(inp As String) As String
    p = 1
    Do While p <= Len(inp)
        If Mid(inp, p, 1) Like "#" Then decod_cle4 = Left(inp, p)
        p = p + 1
    Loop
    On Error Resume Next
End Function
Public Function Affich_gen(ByVal i As Integer, ByVal max_affich As Integer)
    USF_affich_gen.Show 0
    USF_affich_gen.L11.Width = (i / max_affich) * 330
    USF_affich_gen.L11.Caption = Mess & " " & Format(Round((i / max_affich) * 100)) & CStr(" %")
    If Round((i / max_affich) * 100) < 40 Then
        USF_affich_gen.L11.ForeColor = vbBlack
    Else
        USF_affich_gen.L11.ForeColor = vbWhite
    End If
End Function
Public Function FindAll(ByVal sText As String, ByRef osht As Worksheet, ByRef sRange As String, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' FindAll - To find all instances of the1 given string and return the row numbers.
' If there are not any matches the function will return false
' --------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Trap
    '    Dim rfnd As Range    ' Range Object
    Dim iarr, i As Integer   ' Counter for Array
    Dim rFirstAddress    ' Address of the First Find
    ' -----------------
    ' Clear the Array
    ' -----------------
    Erase Armatches
    Set rfnd = osht.Range(sRange).Find(What:=sText, LookIn:=xlValues, LookAt:=xlPart)

    If Not rfnd Is Nothing Then
        rFirstAddress = rfnd.Address
        Do Until rfnd Is Nothing
            iarr = iarr + 1
            ReDim Preserve Armatches(iarr)
            Armatches(iarr) = rfnd.Row   'rFnd.Address pour adresse complete ' rFnd.Row Pour N° de ligne
'            Tab1(iArr, iArr) = rfnd.Row
            pos = rfnd.Row
            Set rfnd = osht.Range(sRange).FindNext(rfnd)
            If rfnd.Address = rFirstAddress Then Exit Do    ' Do not allow wrapped search
        Loop
        FindAll = True
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        FindAll = False
    End If
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        FindAll = False
        Exit Function
    End If
End Function
Public Function FindAll_OneRec(ByVal sText As String, ByVal osht As Worksheet, ByRef sRange As String, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' FindAll - To find all instances of the1 given string and return the row numbers.
' If there are not any matches the function will return false
' --------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Trap
    '    Dim rfnd As Range    ' Range Object
    Dim iarr, i As Integer   ' Counter for Array
    Dim rFirstAddress    ' Address of the First Find
    ' -----------------
    ' Clear the Array
    ' -----------------
    Erase Armatches
    Set rfnd = osht.Range(sRange).Find(What:=sText, LookIn:=xlValues, LookAt:=xlPart)
    If Not rfnd Is Nothing Then
        rFirstAddress = rfnd.Address
    '            Do Until rfnd Is Nothing
        iarr = iarr + 1
        ReDim Preserve Armatches(iarr)
        Armatches(iarr) = rfnd.Row   'rFnd.Address pour adresse complete ' rFnd.Row Pour N° de ligne
        pos = rfnd.Row
        Set rfnd = osht.Range(sRange).FindNext(rfnd)
    '                If rfnd.Address = rFirstAddress Then Exit Do    ' Do not allow wrapped search
    '            Loop
        FindAll_OneRec = True
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        FindAll_OneRec = False
    End If
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        FindAll_OneRec = False
        Exit Function
    End If
End Function
Public Function FindAll_ByArea(rng As Range, ByVal What As Variant, Optional LookIn As XlFindLookIn = xlFormulas, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False, Optional iDoEvents As Boolean = False) As Boolean
    Dim NextResult As Range, Result As Range, area As Range
    Dim FirstMatch As String
    
    If Len(What) > 255 Then Err.Raise 1, "FindAll", "Parameter 'What' must not have more than 255 characters"
    For Each area In rng.Areas
      FirstMatch = ""
      pos = 0
        With area
''            Set NextResult = .Find(What:=What, After:=.Cells(.Cells.Count), LookIn:=LookIn, _
''                                    LookAt:=LookAt, SearchOrder:=SearchOrder, SearchDirection:=xlDown, MatchCase:=MatchCase, MatchByte:=MatchByte, SearchFormat:=SearchFormat)
            Set NextResult = .Find(What:=What, After:=.Cells(.Cells.Count), LookIn:=LookIn, LookAt:=LookAt)
''            Set NextResult = .Find(What:=What, LookIn:=LookIn, LookAt:=LookAt)
''     d5 = InStr(1, c3.Range("N" & k), What, 0) > 0
            If Not NextResult Is Nothing Then
                FirstMatch = NextResult.Address
                pos = NextResult.Row
'                Do
'                    If Result Is Nothing Then
'                        Set Result = NextResult
'                    Else
'                        Set Result = Union(Result, NextResult)
'                    End If
'                    Set NextResult = .FindNext(NextResult)
'
'                    If iDoEvents Then DoEvents
'                Loop While Not NextResult Is Nothing ' And NextResult.Address <> FirstMatch
'            End If
        FindAll_ByArea = True ' Result
         Else
        FindAll_ByArea = False ' Result
         End If
         End With
     Next

End Function
Public Function FindMinTab(ByRef otab As Variant, ByVal sRange As Integer, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' FindTab - To find MIN instance into an array().
' return min value into array and corresponding product name
' trigs alarm set to 2 times recipy_alarm
' If there are not any matches the function will return false
' --------------------------------------------------------------------------------------------------------------
    Dim rfnd As String
    On Error GoTo Err_Trap
    rfnd = ""
    Dim iarr As Integer    ' Counter for Array
    Erase Armatches
    If rfnd = "" Then
        min2 = otab(sRange, 1)
        For i = 2 To j
            If Application.Min(Tabmin(sRange, i)) < min2 Then
                pos = otab(2, i)
                vol = otab(7, i)
                iarr = iarr + 1
                min2 = Application.Min(Tabmin(4, i))
                ctr3 = otab(1, i)
                ReDim Preserve Armatches(iarr)
                Armatches(iarr) = min2  'rFnd.Address pour adresse complete
                rfnd = min2
                FindMinTab = IIf(ctr3 < 3, True, False)
            End If
        Next
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        FindMinTab = False
    End If
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        FindMinTab = False
        Exit Function
    End If
End Function
Public Function Find_Tarif(ByVal sText As Double, ByRef osht As Worksheet, ByRef sRange As String, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' FindAll - To find all instances of the1 given string and return the row numbers.
' If there are not any matches the function will return false
' --------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Trap
    '    Dim rfnd As Range    ' Range Object
    Dim iarr, i As Integer   ' Counter for Array
    Dim rFirstAddress    ' Address of the First Find
    ' -----------------
    ' Clear the Array
    ' -----------------
    Erase Armatches
    i = 0
    nbrowmax = c5.Range("D65000").End(xlUp).Row
'    sstr1 = Replace(sText, ".", ",")
    For i = 2 To nbrowmax
    If sText = c5.Range("D" & CStr(i)) Then
        iarr = iarr + 1
        ReDim Preserve Armatches(iarr)
        Armatches(iarr) = i   'rFnd.Address pour adresse complete ' rFnd.Row Pour N° de ligne
        pos = i
        Find_Tarif = True
        Exit For
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        Find_Tarif = False
    End If
    
    Next
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        Find_Tarif = False
        Exit Function
    End If
End Function
Public Function Find_fact(ByVal sText As String, ByRef tblbd As Variant, ByRef sRange As String, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' FindAll - To find all instances of the1 given string and return the row numbers.
' If there are not any matches the function will return false
' --------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Trap
    '    Dim rfnd As Range    ' Range Object
    Dim iarr, i As Integer   ' Counter for Array
    Dim rFirstAddress    ' Address of the First Find
    ' -----------------
    ' Clear the Array
    ' -----------------
    Erase Armatches
    i = 0
    nbrowmax = c5.Range("D65000").End(xlUp).Row
    For i = 2 To nbrowmax
    If InStr(1, Format(c5.Range("D" & CStr(i)), "###0.00"), CStr(sText), 0) > 0 Then
        iarr = iarr + 1
        ReDim Preserve Armatches(iarr)
        Armatches(iarr) = i   'rFnd.Address pour adresse complete ' rFnd.Row Pour N° de ligne
        pos = i
        Find_fact = True
        Exit For
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        Find_fact = False
    End If
    
    Next
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        Find_fact = False
        Exit Function
    End If
End Function
Public Function Find_All_EBP(ByVal sText As String, ByRef tblbd As Variant, ByRef sRange As String, ByRef Armatches() As String) As Boolean
' --------------------------------------------------------------------------------------------------------------
' Find pour recherche sans filtres
' --------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Trap
    '    Dim rfnd As Range    ' Range Object
    Dim iarr, i As Integer   ' Counter for Array
    Dim rFirstAddress    ' Address of the First Find
    ' -----------------
    ' Clear the Array
    ' -----------------
    Erase Armatches
    i = 0
    nbrowmax = c6.Range("B65000").End(xlUp).Row
    For i = 3840 To nbrowmax
        If (Left(c6.Range("B" & i), 3) = "411" And InStr(1, cle_rech, c6.Range("G" & i))) > -1 Then
        iarr = iarr + 1
        ReDim Preserve Armatches(iarr)
        Armatches(iarr) = i   'rFnd.Address pour adresse complete ' rFnd.Row Pour N° de ligne
        Debug.Print Armatches(iarr)
        pos = i
        Find_All_EBP = True
    Else
    ' ----------------------
    ' No Value is Found
    ' ----------------------
        Find_All_EBP = False
    End If
     Next i
    ' -----------------------
    ' Error Handling
    ' -----------------------
Err_Trap:
    If Err <> 0 Then
    '        MsgBox ("Erreur " & Str(Err.Number) & " " & "Find All-Non trouvé")
        Err.Clear
        Find_factFind_All_EBP = False
        Exit Function
    End If
End Function
'
'Function NumCoulCel(c As Object)
'Application.Volatile True
'NumCoulCel = Abs(c.Interior.ColorIndex)
'End Function
'Function SommeCol(plage As Range, Couleur As Long) As Double
'  Application.Volatile
'  Dim Cellule As Range
'  Dim Somme As Double
'  For Each Cellule In plage
'    If Cellule.Interior.Color = Couleur And IsNumeric(Cellule.Value) Then Somme = Somme + Cellule.Value
'  Next Cellule
'  SommeCol = Somme
'End Function
'_______________________________________________________________________________________________________________
'With Worksheets("Feuil1").Shapes.AddShape(msoShapeRectangle, 40, 80, 140, 50)
'    .Name = "NomForme"
'    .TextFrame.Characters.Text = "Le texte dans la forme"
'End With




