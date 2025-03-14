Attribute VB_Name = "Xtract_1"
Option Explicit
Public Tab4(6) As Integer
Public Tab3(2) As String
Public Sub Calc_Theo_cumul_clients()  ' MISE A JOUR COL J DANS CLIENTS  MISE A JOUR COL J DANS CLIENTS
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Set c6 = Sheets("EBP-Xtract-expert")
    Set c7 = Sheets("Buff2")
    Set c8 = Sheets("Gestion")
    nbrowmax = c3.Range("I65000").End(xlUp).Row
    For i = 2 To 10 ' nbrowmax
        c8.Cells(1, 1) = Val(Month(Date)) * c3.Range("S" & i)
        res1 = 0
        res2 = 0
        c3.Range("J" & i) = Val(Month(Date)) * c3.Range("S" & i)
    '        sstr1 = c3.Range("N" & i)
        res1 = calc_Xtrac_Dom(i)
        res2 = calc_trav_For_all_Year(i)
        c3.Range("K" & i) = res1 + res2
    '        c3.Range("J" & i) = res + c3.Range("J" & i)
    Next i
    nbrowmax = c7.Range("G65000").End(xlUp).Row
    '    res = 0
    '    For i = 2 To nbrowmax
    '        sstr1 = c3.Range("N" & i)
    ''        res = calc_Xtrac_Dom(i)
    '    Next i
End Sub
Public Sub Calc_Reel_cumul_clients()  ' MISE A JOUR COL K DANS CLIENTS A PARTIR DE Xtract-expert
    Call init_rep2
    ficopen = Path2 & "Listing_Travaux_Log - v2.html"
    USF_affich_gen.Show 0
'    Open ficopen For Output As #1
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Set c6 = Sheets("EBP-Xtract-expert")
    Set c7 = Sheets("Buff2")
    nbrowmax = c3.Range("F65000").End(xlUp).Row
    max_affich = nbrowmax
    For i = 2 To nbrowmax
        res1 = 0
        res2 = 0
        j = i
        cle_rech = UCase(c3.Range("N" & i))
        cle_rech2 = c3.Range("O" & i)
    '        Workbooks("Facturation-auto-mail-MIDI-SERVICES-01.XLSM").Worksheets("EBP-Xtract-expert").Range("A3:J25000").Select
        c6.Activate
        c6.AutoFilterMode = False
        c6.Range("A2:J22000").Select
        If InStr(1, cle_rech, c6.Range("G" & i)) > -1 Then
    '        If ((Left(c6.Range("B" & i), 3) = "411" And InStr(1, cle_rech, c6.Range("G" & i)) > -1)) Then
    '            If Find_All_EBP(cle_rech2, Sheets("EBP-Xtract-expert"), "G3840:G" & nbrowmax, Armatches()) <> Empty Then
    '''''            For j = 1 To UBound(Armatches())
            Selection.AutoFilter field:=1, Criteria1:="=" & "*"
            Selection.AutoFilter field:=2, Criteria1:="=" & "411*"
            Selection.AutoFilter field:=7, Criteria1:="=" & cle_rech & "*", Operator:=xlAnd
            LastFilterRow = c6.Range(["B65535"]).End(xlUp).Row
    '''''        If InStr(1, cle_rech, c6.Range("G" & LastFilterRow)) > -1 Then
    '''''            Debug.Print cle_rech, c6.Range("G" & LastFilterRow)
            c7.Range("A3:J2000").ClearContents
            If LastFilterRow > 1 Then
                Range("A2", Cells(LastFilterRow, 10)).Copy _
                        Destination:=c7.Cells(2, 1)
                c7.AutoFilterMode = False
                nbrowmax2 = c7.Range("B65000").End(xlUp).Row
                c3.Range("O" & j) = c7.Range("G" & 5)
                c3.Range("J" & i) = res2
                res1 = calc_Xtrac_Dom(i)
                res2 = calc_trav_For_all_Year(i) + res1
                c3.Range("O" & j) = c7.Range("G" & 5)
                c3.Range("J" & i) = res2
    '                Debug.Print cle_rech; res2
    '            Debug.Print c7.Range("G" & i); c7.Range("H" & i); c7.Range("I" & i); c7.Range("J" & i); cle_rech; "  "; Format(res2, "###0.00") & "€"
                c3.Range("K" & j) = res2
            ElseIf InStr(1, cle_rech2, c6.Range("G" & i)) > -1 Then
                Selection.AutoFilter field:=1, Criteria1:="=" & "*"
                Selection.AutoFilter field:=2, Criteria1:="=" & "411*"
                Selection.AutoFilter field:=7, Criteria1:="=" & cle_rech2 & "*", Operator:=xlAnd
                LastFilterRow = c6.Range(["B65535"]).End(xlUp).Row
    '''''        If InStr(1, cle_rech, c6.Range("G" & LastFilterRow)) > -1 Then
    '''''            Debug.Print cle_rech, c6.Range("G" & LastFilterRow)
                c7.Range("A3:J2000").ClearContents
                Range("A2", Cells(LastFilterRow, 10)).Copy _
                        Destination:=c7.Cells(2, 1)
                c7.AutoFilterMode = False
                nbrowmax2 = c7.Range("B65000").End(xlUp).Row
'                c3.Range("O" & j) = cle_rech2
                res1 = calc_Xtrac_Dom(i)
                res2 = calc_trav_For_all_Year(i) + res1
            Else
                c3.Range("K" & j) = ""
    '           c3.Range("O" & j) = ""
'                Debug.Print c7.Range("G" & i); cle_rech; cle_rech2; " ..Pas de record EBP"
'                Print #1, c7.Range("G" & i); cle_rech; "    ERREUR ..PAS D'ENREGT EBP"
            End If
    '''''            Next j
    '        c3.Range("J" & i) = res + c3.Range("J" & i)
        End If
        X = Affich_gen(i, max_affich)
    Next i
'    Close #1
End Sub
Public Sub OpenUrl()
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "file:///E:\Dev-VBA\Midi-services\Send_Facturation/Etat_Clients%20-%20v1.html", path3, 1)
End Sub
Public Function cmpdt(ByVal t0 As String, ByVal t1 As String, ByVal t2 As String) As String
    If (t0 = "") And (t1 = "") Or (t0 = 0) And (t1 = 0) Then Exit Function
    t10 = Round((t1 / t0) * 100)
    Debug.Print t10
    If (t10 > 80 And t10 <= 100) Then  ' Calcul retard paiement de 1 mois
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#336600"
    ElseIf ((t10 > 60) And (t10 <= 80)) Then
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#33FF00"
    ElseIf (t10 > 40 And t10 <= 60) Then
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#669900"
    ElseIf (t10 > 20 And t10 <= 40) Then
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#666600"
    ElseIf (t10 > 0 And t10 <= 20) Then
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#660000"
    Else
        cmpdt = CStr(Abs(t1 - t0))
        oColor = "#FF0000"
    End If
End Function
Public Function gm(texte As String) As String
    Dim G As String
    G = """"
    gm = G & texte & G
End Function
Function MSA$(ByVal chaine$)
    Const VAccent = "àáâãäåéêëèìíîïðòóôõöùúûüç-° ", VSsAccent = "aaaaaaeeeeiiiioooooouuuuc . "
    Dim Bcle&
    For Bcle = 1 To Len(VAccent)
        chaine = Replace(chaine, Mid(VAccent, Bcle, 1), Mid(VSsAccent, Bcle, 1))
    Next Bcle
    MSA = chaine
End Function
Sub filrage_1()
    Range("A3:J25000").Select
    Selection.AutoFilter field:=1, Criteria1:="=" & Range("B2").text & "706001", _
                         Criteria2:="=" & Range("B2").text & "706003", _
                         Operator:=xlOr
    sstr2 = "SMA"
    Selection.AutoFilter field:=7, Criteria1:="=" & sstr2 & "*", Operator:=xlAnd
    ''''Set plage = [_filterdatabase].offset(1).Resize(, 1)
    ''''Set plage = plage.Resize(plage.Count - 1).SpecialCells(xlCellTypeVisible)
    ''''MsgBox "Nombre de lignes affichées = " & plage.Count
    ''''MsgBox "Première ligne affichée = " & plage.Row
    ''''plage(plage.Count).Select
    ''''MsgBox "Dernière ligne affichée = " & plage.SpecialCells(xlCellTypeLastCell).Row
End Sub
Public Function calc_trav_For_all_Year(ByVal k As Integer, Optional ByVal i As Integer)
    If i = 0 Then i = k
    Set c3 = Worksheets("CLIENTS")
    cle_rech = c3.Range("N" & i)
    Set c2 = Sheets("Travaux")
    nbrowmax2 = c2.Range("B65000").End(xlUp).Row
    t6 = c2.Range("G" & i)
    res2 = 0
    If FindAll(cle_rech, Sheets("Travaux"), "B2:B" & nbrowmax2, Armatches()) <> Empty Then
        For i = 1 To UBound(Armatches)
            res2 = res2 + Val(c2.Range("D" & Armatches(i))) * Val(c2.Range("E" & Armatches(i)))
        Next i
    End If
    calc_trav_For_all_Year = res2
End Function
Public Function calc_trav_For_LastMonth(ByVal k As Integer, Optional ByVal i As Integer)
    cle_rech = c3.Range("N" & i)
    Set c2 = Sheets("Travaux")
    nbrowmax2 = c2.Range("B65000").End(xlUp).Row
    t6 = c2.Range("G" & i)
    res2 = 0
    If FindAll(cle_rech, Sheets("Travaux"), "B1:B" & nbrowmax2, Armatches()) <> Empty Then
        For j = 1 To UBound(Armatches)
            If comp_mois(k, c2.Range("G" & Armatches(j))) = k Then
                res2 = res2 + Val(c2.Range("D" & Armatches(j))) * Val(c2.Range("E" & Armatches(j)))
            End If
        Next j
    End If
    calc_trav_For_LastMonth = res2
End Function
Private Function calc_Xtrac_Dom(ByVal i As Integer)
    Set c7 = Sheets("Buff2")
    nbrowmax = c7.Range("G65000").End(xlUp).Row
    For i = 3 To nbrowmax
        If c7.Range("H" & i) <> "" Then
            nbrowmax = i
        ElseIf c7.Range("H" & i) = "" Then
            Exit For
        End If
    Next i
    res1 = 0
    For i = 3 To nbrowmax
        If c7.Range("H" & i) = "C" Then
            res1 = res1 + c7.Range("I" & i)
        ElseIf c7.Range("H" & i) = "D" Then
            res1 = res1 - c7.Range("I" & i)
        ElseIf c7.Range("H" & i) = "" Then Exit For
        End If
    Next i
    '        End With
    '    End If
    '        End With
    calc_Xtrac_Dom = Format(res1, "###0.00")
End Function
Function filter_has_results() As Boolean
    nbrow = 0
    nbrow = c7.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    Set plage = [_filterdatabase].offset(1).Resize(, 1)
    Set plage = plage.Resize(plage.Count - 1).SpecialCells(xlCellTypeVisible)
    nbrow = plage.Count
    If nbrow = 0 Then
        filter_has_results = True
    Else
        filter_has_results = False
    End If
End Function
