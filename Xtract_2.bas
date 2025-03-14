Attribute VB_Name = "Xtract_2"
Option Explicit
Public Tab4(6) As Integer
Public Tab3(2) As String
Public Sub Calc_pmt_clients_2()    '____TEST____TEST____TEST____TEST____TEST____TEST
    Application.ScreenUpdating = True
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Set c6 = Sheets("EBP-Xtract-expert")
    Set c7 = Sheets("Buff2")
    Set c8 = Sheets("Gestion")
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    L = 1
    For k = 1 To (2 * Val(Month(Date))) Step 2
        c8.Cells(1, k) = Left(comp_mois_rev(k), 5) & "/MIDI"
        c8.Cells(1, k + 1) = Left(comp_mois_rev(k, Val(Month(Date))), 5) & "/EBP"
        For i = 2 To nbrowmax
            res1 = 0
            res2 = 0
            res1 = L * c3.Range("S" & i)
            res2 = calc_trav_For_all_Year(k, i)
            c8.Cells(i, k) = res1 + res2
            c8.Cells(i, k + 1) = ""    ' res1 + res2
    '        c3.Range("J" & i) = res + c3.Range("J" & i)
        Next i
        L = L + 1
    Next k
    '    nbrowmax = c7.Range("G65000").End(xlUp).Row
    '    res = 0
    '    For i = 2 To nbrowmax
    '        sstr1 = c3.Range("N" & i)
    ''        res = calc_Xtrac_Dom(i)
    '    Next i
End Sub
Public Sub Calc_pmt_EBP_clients()
    Call init_rep2
    ficopen = Path2 & "Listing_Travaux_Log - v2.html"
    USF_affich_gen.Show 0
    Open ficopen For Output As #1
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
        If InStr(1, cle_rech2, c6.Range("G" & i)) > -1 Then
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
    '''''
                c3.Range("O" & j) = c7.Range("G" & 5)
                res1 = calc_Xtrac_Dom(i)
                res2 = calc_trav_For_all_Year(i) + res1
    '                Debug.Print cle_rech; res2
    '            Debug.Print c7.Range("G" & i); c7.Range("H" & i); c7.Range("I" & i); c7.Range("J" & i); cle_rech; "  "; Format(res2, "###0.00") & "€"
                c3.Range("K" & j) = res2
            Else
                c3.Range("K" & j) = ""
    '             c3.Range("O" & j) = ""
                Debug.Print c7.Range("G" & i); cle_rech; "    ERREUR ..PAS D'ENREGT EBP"
                Print #1, c7.Range("G" & i); cle_rech; "    ERREUR ..PAS D'ENREGT EBP"
    '                Beep
            End If
    '''''            Next j
    '        c3.Range("J" & i) = res + c3.Range("J" & i)
        End If
        T = Affich_gen(i, max_affich)
    Next i
    '        nbrowmax = c7.Range("G65000").End(xlUp).Row
    '    For i = 2 To nbrowmax
    '        sstr1 = c3.Range("N" & i)
    '        res = calc_Xtrac_Dom(i)
    '        c3.Range("K" & i) = res
    '    Next i
    Close #1
End Sub
Public Sub OpenUrl()
    Dim lSuccess As Long
    Dim zlparent
'    lSuccess = ShellExecute(0, "Open", "file:///E:\Dev-VBA\Midi-services\Send_Facturation/Listing_Travaux_Log%20-%20v2.html", path3, 1)
    lSuccess = ShellExecute(zlparent, "open", "file:///E:\Dev-VBA\Midi-services\Send_Facturation/Etat_Clients - v1.html", "", "", SW_SHOWNORMAL)
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
Private Function calc_Xtrac_Dom2(ByVal i As Integer)
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
'        With Selection
'    cle_rech = UCase(cle_rech)
'    If FindAll(cle_rech, Sheets("Buff2"), "G2:G" & nbrowmax, Armatches()) <> Empty Then
'            With Worksheets("EBP-Xtract-expert").AutoFilter.Range
'
'    Set plage = [_filterdatabase].offset(1).Resize(, 1)
'    Set plage = plage.Resize(plage.Count - 1).SpecialCells(xlCellTypeVisible)
'MsgBox "Nombre de lignes affichées = " & plage.Count
'MsgBox "Première ligne affichée = " & plage.Row
'plage(plage.Count).Select
'    MsgBox "Dernière ligne affichée = " & plage.SpecialCells(xlCellTypeLastCell).Row




