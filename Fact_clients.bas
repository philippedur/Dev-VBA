Attribute VB_Name = "Fact_clients"
Public Function Facture_clients(ByRef sh As Worksheet)
    Call init_rep2
    lnk = Path2 & "Facturation-auto-mail-MIDI-SERVICES-01.xlsm"
    '    Workbooks.Open Filename:=lnk
    Dim bb, cc, NumFactAb, NumFactAn, comp As Integer
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    '    Call RAZ
    Dim CurrentWorkbook As String
    Dim CurrentFormat As Long
    CurrentWorkbook = ThisWorkbook.FullName
    CurrentFormat = ThisWorkbook.FileFormat
    Application.ScreenUpdating = True
    nbCol = c3.[c1].CurrentRegion.Columns.Count
    nbrow = c3.Range("N65000").End(xlUp).Row
    r1 = nbrow
    With Sheets("modele1")    ' .Open(Filename:=lnk).Sheets("modele1")
        Dim name_facture, num_facture, name_pdf As String
        Sheets("modele1").Range("E8") = CStr(Format(Date, "dd mmmm yyyy"))
        Sheets("modele1").Range("E8").HorizontalAlignment = xlHAlignLeft
        Sheets("modele1").Range("E8").Font.Bold = True
        Sheets("modele1").Range("E8").Font.Size = 22
        Sheets("modele1").Range("H11") = UCase(Format(Date, "mmmm"))
        Sheets("modele1").Range("H11").HorizontalAlignment = xlHAlignCenter
        Sheets("modele1").Range("H11").Font.Bold = True
        Sheets("modele1").Range("H11").Font.Size = 12
        For k = 2 To nbrow
            periodicite = c3.Range("X" & k)
            date_creation = c3.Range("D" & k)
            Societe = c3.Range("N" & k)
            If date_creation <> 0 And (periodicite <> 0) Then
                res = (calc_period2(Societe, k, Date, periodicite, date_creation))
                If trig_fact = True Then
                    .[champ1] = "Société:   " & c3.Range("N" & k)
                    sstr1 = c3.Range("N" & k)
                    .[champ1].Font.Bold = True
                    .[champ1].Font.Size = 11
                    .[champ1].Font.Name = "Calibri"
                    .[champ2] = "Gérant:  Mr " & c3.Range("F" & k)
                    .[champ2].Font.Bold = False
                    .[champ2].Font.Size = 11
                    .[champ2].Font.Name = "Calibri"
                    name_facture = c3.Range("E" & k) & " " & sh.Range("F" & k)
                    .[adresse1] = c3.Range("A" & k) & " " & c3.Range("B" & k) & " " & c3.Range("C" & k)
                    .[adresse1].Font.Size = 11
                    .[adresse1].Font.Name = "Calibri"
                    .[CP] = ""    ' c3.Range("B" & k) & " " & c3.Range("K" & k)
                    .[CP].Font.Size = 11
                    .[CP].Font.Name = "Calibri"
                    '                    .[Ville] = ""  ' c3.Range("S" & k)
                    .[TYP_CLIENT] = c3.Range("R" & k)
                    .[TYP_CLIENT].Font.Size = 11
                    .[TYP_CLIENT].Font.Name = "Calibri"
                    .[num_client] = c3.Range("G" & k)
                    .[num_client].Font.Size = 11
                    .[num_client].Font.Name = "Calibri"
                    .[num_facture] = c3.Range("A" & k)
                    Call etat_gene
                    .[date_facture] = CStr(Format(Date, "mm/dd/yy"))
                    .[date_facture].Font.Size = 11
                    .[date_facture].Font.Name = "Calibri"
                    .[echeance] = UCase(Format(Date, "mmmm"))   ' c3.Range("E" & k)
                    .[echeance].Font.Size = 11
                    .[echeance].Font.Name = "Calibri"
                    .[echeance].Font.Size = 11
                    .[echeance].Font.Name = "Calibri"
                    ''                    Sheets("modele1").Range("B14") = "Code"
                    ''                    Sheets("modele1").Range("C14") = "Libellé"
                    ''                    Sheets("modele1").Range("F14") = "PU/HT"
                    ''                    Sheets("modele1").Range("G14") = "Nb"
                    ''                    Sheets("modele1").Range("H14") = ""
                    ''                    Sheets("modele1").Range("I14") = "TOTAL/HT"
                    Call calcul_travaux
                    '            .[Prix_HT] = tarif(c3.Range("R" & k)) * c2.Range("D" & ArMatches(i))
                    '            .[Prix_HT].Font.Size = 11
                    '            .[Prix_HT].Font.Name = "Calibri"
                    '            .[PU_HT] = tarif(c3.Range("R" & k))
                    '            .[PU_HT].Font.Size = 11
                    '            .[PU_HT].Font.Name = "Calibri"
                    total = 0
                    For i = 12 To 28 ' 32
                        total = total + Val(c1.Cells(i, 8))
                    Next i
                    '            c3.Cells(Target.Row, 19).NumberFormat = "# ##0.00__€"
                    .[Total_HT].NumberFormat = "# ##0.00__€"
                    .[Total_HT] = total
                    .[Total_HT].Font.Size = 11
                    .[Total_HT].Font.Name = "Calibri"
                    .[TVA_20].NumberFormat = "# ##0.00__€"
                    .[TVA_20] = total / 100 * 20
                    .[TVA_20].Font.Size = 11
                    .[TVA_20].Font.Name = "Calibri"
                    .[Total_TTC].NumberFormat = "# ##0.00__€"
                    .[Total_TTC] = .[Total_HT] + .[TVA_20]
                    .[Total_TTC].Font.Size = 11
                    .[Total_TTC].Font.Name = "Calibri"
                    '            c3.Range("J" & k) = .[Prix_HT]
                    '            c3.Range("K" & k) = .[Prix_HT]
                    ctr1 = 0
                    '            For i = 2 To 1000
                    '                If ctr1 < Val(Right(c3.Range("H" & i), 4)) Then
                    '                    ctr1 = Val(Right(c3.Range("H" & i), 4))
                    '                    NumFactAb = ctr1
                    '                End If
                    '            Next
                    '            NumFactAn = ctrl
                    ''            .[num_facture] = "F" & c3.Range("G" & k) & "/" & CStr(Format(c3.Range("E" & k), "mm")) & CStr(Format(c3.Range("E" & k), "yy"))
                    .[num_facture] = "F" & c3.Range("G" & k) & "/" & CStr(Format(Date, "mm")) & CStr(Format(Date, "yy"))
                    .[num_facture].Font.Size = 11
                    .[num_facture].Font.Name = "Calibri"
                    .[num_facture].Font.Size = 11
                    .[num_facture].Font.Name = "Calibri"
                    .[champ1]
                    Row_IPub = 20
                    '            name_xls = CStr("Société: " & c3.Range("N" & k) & "-" & num_facture & "-" & c3.Range("E" & k) & ".csv")
                    Call set_rep
                    '            rep_pdf = path3
                    name_pdf = CStr("Fact. Société:" & "__" & c3.Range("N" & k) & "__" & .[num_facture] & "__" & ".pdf")
                    Set pdfjob = CreateObject("PDFCreator.clsPDFCreator")
                    '            nomexcel = Worksheets(1).Select
                    '            nompdf = Left(nomexcel, Len(nomexcel) - 4) & ".pdf"
                    With pdfjob
                        If .cStart("/NoProcessingAtStartup") = False Then
                            MsgBox "Can't initialize PDFCreator.", vbCritical + vbOKOnly, "PrtPDFCreator"
                        End If
                        .cOption("UseAutosave") = 1
                        .cOption("UseAutisaveDirectory") = 1
                        .cOption("AutosaveDirectory") = rep_pdf
                        .cOption("AutosaveFilename") = name_pdf
                        .cOption("AutosaveFormat") = 0
                        .cClearCache
                    End With
                    Worksheets("modele1").PrintOut From:=1, To:=1, copies:=1, ActivePrinter:="PDFCreator"
                    '            Worksheets(1).PrintOut Copies:=1, ActivePrinter:="PDFCreator"
                    Do Until pdfjob.cCountOfPrintjobs = 1
                        DoEvents
                    Loop
                    pdfjob.cPrinterStop = False
                    Do Until pdfjob.cCountOfPrintjobs = 0
                        DoEvents
                    Loop
                    With pdfjob
                        .cDefaultPrinter = DefaultPrinter
                        .cClearCache
                        .cClose
                    End With
                    Set pdfjob = Nothing
                    '            Workbook.Close
                    '                    Call Affiche_pct("Génération des pdf.. ")
                    USF61b.Label8.Visible = True
                End If
            End If
        Next k
    End With
    '    Call RAZ
    '        ActiveWorkbook.Close
End Function
Sub Macro1()
    Set c2 = Sheets("Travaux")
    cle_rech = "CAD Construction"
    nbrowmax = Sheets("Travaux").Range("B5000").End(xlUp).Row
    If FindAll(cle_rech, Sheets("Travaux"), "B2:B" & nbrowmax, Armatches()) Then
        j = 1
        For i = 1 To UBound(Armatches())
            If (Year(Sheets("Travaux").Cells(Armatches(i), 8)) = Year(Date)) And (Sheets("Travaux").Cells(Armatches(i), 7) = "FEVRIER") Then
                Debug.Print c2.Cells(Armatches(i), 1), c2.Cells(Armatches(i), 2), c2.Cells(Armatches(i), 5), c2.Cells(Armatches(i), 6), c2.Cells(Armatches(i), 7), c2.Cells(Armatches(i), 8), j
                j = j + 1
            End If
        Next i
    End If
End Sub
Public Sub RAZ()
    Set c1 = Sheets("modele1")
    With Sheets("modele1")
        .[champ1] = ""
        .[champ2] = ""
        .[adresse1] = ""
        .[CP] = ""
        '        .[Ville] = ""
        .[TYP_CLIENT] = ""
        .[num_client] = ""
        .[num_facture] = ""
        .[date_facture] = ""
        .[echeance] = ""
        .Range("E8") = ""
        '            .[Prix_HT] = tarif(c3.Range("R" & k)) * c2.Range("D" & ArMatches(i))
        '            .[Prix_HT].Font.Size = 11
        '            .[Prix_HT].Font.Name = "Calibri"
        .[PU_HT] = ""
        .[Total_HT] = ""
        .[TVA_20] = ""
        offset = 11
        For i = 1 To 100
            Select Case (offset + i)

                Case Is = 38
                    c1.Cells(offset + i, 1) = ""
                    c1.Cells(offset + i, 2) = ""
                    c1.Cells(offset + i, 3) = ""
                    c1.Cells(offset + i, 4) = ""
                    c1.Cells(offset + i, 5) = ""
                    c1.Cells(offset + i, 6) = ""
                Case Is = 39
                    c1.Cells(offset + i, 1) = ""
                    c1.Cells(offset + i, 2) = ""
                    c1.Cells(offset + i, 3) = ""
                    c1.Cells(offset + i, 4) = ""
                    c1.Cells(offset + i, 5) = ""
                    c1.Cells(offset + i, 6) = ""
                Case Is = 40
                    c1.Cells(offset + i, 1) = ""
                    c1.Cells(offset + i, 2) = ""
                    c1.Cells(offset + i, 3) = ""
                    c1.Cells(offset + i, 4) = ""
                    c1.Cells(offset + i, 5) = ""
                    c1.Cells(offset + i, 6) = ""
                Case Is < 42
                    c1.Cells(offset + i, 1) = ""
                    c1.Cells(offset + i, 2) = ""
                    c1.Cells(i, 3) = ""
                    c1.Cells(offset + i, 4) = ""
                    c1.Cells(offset + i, 5) = ""
                    c1.Cells(ooffsetffset + i, 6) = ""
                    c1.Cells(offset + i, 7) = ""
                    c1.Cells(offset + i, 8) = ""
                Case Is > 46
                    c1.Cells(offset + i, 1) = ""
                    c1.Cells(offset + i, 2) = ""
                    c1.Cells(offset + i, 3) = ""
                    c1.Cells(offset + i, 4) = ""
                    c1.Cells(offset + i, 5) = ""
                    c1.Cells(offset + i, 6) = ""
                    c1.Cells(offset + i, 7) = ""
                    c1.Cells(offset + i, 8) = ""
            End Select
        Next i

        c1.Range("A38:I46") = ""

        c1.Cells(29, 7) = "Total_HT"
        c1.Cells(30, 7) = "TVA_20"
        c1.Cells(31, 7) = "Total_TTC"
        c1.Cells(32, 8) = ""
        c1.Cells(22, 8) = ""
        c1.Cells(34, 8) = ""


        c1.Cells(32, 2) = "Vous réglez par virement ? Indiquez votre n° de facture sur l'ordre de virement"
        c1.Cells(33, 2) = "IBAN : FR76 1460 7000 6505 2215 1348 002 - BIC : CCBPFRPPMAR"
        c1.Cells(34, 2) = "Sarl au capital de 7.622 € - 69, rue du Rouet 13008 MARSEILLE - Tél.: 04 91 79 35 38"
        c1.Cells(35, 2) = " RC 902 B 654 - Code NAF : 8299Z - SIRET 377 491 154 00011"
        c1.Cells(36, 2) = "Agrément Préfecture des BDR n° 2010/AEFDJ/13/015"
    End With
End Sub
Private Function tarif(inp As String) As Integer
    Select Case inp
        Case Is = "A"
            tarif = c4.Range("D2")
        Case Is = "B"
            tarif = c4.Range("D3")
        Case Is = "C"
            tarif = c4.Range("D4")
        Case Is = "D"
            tarif = c4.Range("D5")
        Case Is = "E"
            tarif = c4.Range("D6")
        Case Is = "F"
            tarif = c4.Range("D7")
        Case Is = "G"
            tarif = c4.Range("D8")
        Case Is = "H"
            tarif = c4.Range("D9")
        Case Is = "I"
            tarif = c4.Range("D10")
        Case Is = "J"
            tarif = c4.Range("D11")
        Case Is = "K"
            tarif = c4.Range("D12")
        Case Is = "L"
            tarif = c4.Range("D13")
        Case Is = "M"
            tarif = c4.Range("D14")
        Case Is = "N"
            tarif = c4.Range("D15")
        Case Is = "O"
            tarif = c4.Range("D16")
        Case Is = "P"
            tarif = c4.Range("D17")
        Case Is = "Q"
            tarif = c4.Range("D18")
        Case Is = "R"
            tarif = c4.Range("D19")
        Case Is = "S"
            tarif = c4.Range("D20")
        Case Is = "T"
            tarif = c4.Range("D21")
        Case Is = "U"
            tarif = c4.Range("D22")
        Case Is = "V"
            tarif = c4.Range("D23")
        Case Is = "W"
            tarif = c4.Range("D24")
        Case Is = "X"
            tarif = c4.Range("D25")
    End Select
End Function
Private Sub calcul_travaux()
    Set c1 = Sheets("modele1")
    Set c7 = Sheets("Buff2")
    Set c2 = Sheets("Travaux")
    With c1
        offset = 13 ' 15
        If trig_fact = True Then
            c1.Cells(offset, 2) = "DOM"
            c1.Cells(offset, 3) = dom_type(k, periodicite)
            c1.Cells(offset, 7) = ind_fact    ' 1
            c1.Cells(offset, 6) = Worksheets("CLIENTS").Range("S" & k).Value

            c1.Cells(offset, 8) = Worksheets("CLIENTS").Range("S" & k).Value * ind_fact

            For i = 1 To 17 ' (36 - offset)
                c1.Cells(offset + i, 2) = ""
                c1.Cells(offset + i, 3) = ""
                c1.Cells(offset + i, 7) = ""
                c1.Cells(offset + i, 6) = ""
                c1.Cells(offset + i, 8) = ""
            Next i
            '            Call RAZ
            Set c2 = Sheets("Travaux")
            Smois = comp_mois_rev(Month(Date) - 1)

            '               If Find_trav(societe, Smois, Sheets("Travaux"), "B2:B10000", Armatches()) <> Empty Then
            Call filter_travaux(Societe, Smois, Armatches)
            Set c7 = Sheets("Buff2")
            nbrowmax2 = c7.Range("A65000").End(xlUp).Row
            c1.Cells(38, 7) = "Total_HT"
            c1.Cells(39, 7) = "TVA_20"
            c1.Cells(40, 7) = "Total_TTC"
            If nbrowmax2 >= 1 Then
                j = 2
                c1.Cells(offset + j, 2) = "TRAV"
                c1.Cells(offset + j, 3) = "TRAVAUX ADDITIONNELS DIVERS"
                j = 4
                '''                xx = Date - 31
                '''                yy = "01/" & Right(xx, 7)
                '''                d_date = CDate(yy)
                '''                sstr2 = comp_mois_rev(Month(Date - 31))
                For i = 1 To nbrowmax2
                    c1.Cells(offset + j, 2) = c7.Range("F" & i)
                    c1.Cells(offset + j, 3) = c7.Range("C" & i)
                    c1.Cells(offset + j, 7) = c7.Range("D" & i)
                    c1.Cells(offset + j, 6) = c7.Range("E" & i)
                    c1.Cells(offset + j, 8) = c7.Range("D" & i) * (c7.Range("E" & i))
                    j = j + 1
                Next i
            End If
        End If
    End With
End Sub
Private Sub etat_gene()
    USF61b.Label5.Caption = "Gérant: " & c3.Range("F" & k)
    USF61b.Label5.Visible = True
    USF61b.Label5.ForeColor = RGB(15, 5, 107)
    USF61b.Label5.Font.Size = 11
    USF61b.Label5.Font.Name = "Calibri"
    USF61b.Label6.Caption = "Generation pdf en cours:"
    USF61b.Label6.Visible = True
    USF61b.Label6.ForeColor = RGB(15, 5, 107)
    USF61b.Label6.Font.Size = 11
    USF61b.Label6.Font.Name = "Calibri"
    USF61b.Label7.Caption = "Société: " & c3.Range("N" & k)
    USF61b.Label7.Visible = True
    USF61b.Label7.ForeColor = RGB(15, 5, 107)
    USF61b.Label7.Font.Size = 11
    USF61b.Label7.Font.Name = "Calibri"
    '    USF61b.Label8.Caption = "Expédition terminée."
    USF61b.Label8.Visible = False
    USF61b.Label8.ForeColor = RGB(15, 5, 107)
    USF61b.Label8.Font.Size = 11
    USF61b.Label8.Font.Name = "Calibri"
End Sub
Function comp_mois(ByVal mois_Trav As String) As Integer
    Select Case mois_Trav
        Case Is = "JANVIER"
            comp_mois = 1
        Case Is = "FEVRIER"
            comp_mois = 2
        Case Is = "MARS"
            comp_mois = 3
        Case Is = "AVRIL"
            comp_mois = 4
        Case Is = "MAI"
            comp_mois = 5
        Case Is = "JUIN"
            comp_mois = 6
        Case Is = "JUILLET"
            comp_mois = 7
        Case Is = "AOUT"
            comp_mois = 8
        Case Is = "SEPTEMBRE"
            comp_mois = 9
        Case Is = "OCTOBRE"
            comp_mois = 10
        Case Is = "NOVEMBRE"
            comp_mois = 11
        Case Is = "DECEMBRE"
            comp_mois = 12
    End Select
End Function
Function comp_mois_rev(ByVal k As Integer) As String
    comp_mois_rev = ""
    k = Month(Date)
    If k = 1 Then
    k = Month(Date) + 11
    Else
    k = Month(Date) - 1
    End If
    Select Case k
        Case Is = 1
            comp_mois_rev = "JANVIER"
        Case Is = 2
            comp_mois_rev = "FEVRIER"
        Case Is = 3
            comp_mois_rev = "MARS"
        Case Is = 4
            comp_mois_rev = "AVRIL"
        Case Is = 5
            comp_mois_rev = "MAI"
        Case Is = 6
            comp_mois_rev = "JUIN"
        Case Is = 7
            comp_mois_rev = "JUILLET"
        Case Is = 8
            comp_mois_rev = "AOUT"
        Case Is = 9
            comp_mois_rev = "SEPTEMBRE"
        Case Is = 10
            comp_mois_rev = "OCTOBRE"
        Case Is = 11
            comp_mois_rev = "NOVEMBRE"
        Case Is = 12
            comp_mois_rev = "DECEMBRE"
    End Select
End Function
Function comp_date(k, ByVal mois_Trav As String) As Integer
    mois_Trav = c2.Range("G" & i)
    Select Case mois_Trav
        Case Is = "JANVIER"
            comp_date = 1
        Case Is = "FEVRIER"
            comp_date = 2
        Case Is = "MARS"
            comp_date = 3
        Case Is = "AVRIL"
            comp_date = 4
        Case Is = "MAI"
            comp_date = 5
        Case Is = "JUIN"
            comp_date = 6
        Case Is = "JUILLET"
            comp_date = 7
        Case Is = "AOUT"
            comp_date = 8
        Case Is = "SEPTEMBRE"
            comp_date = 9
        Case Is = "OCTOBRE"
            comp_date = 10
        Case Is = "NOVEMBRE"
            comp_date = 11
        Case Is = "DECEMBRE"
            comp_date = 12
    End Select
End Function
Function dom_type(ByVal k As Integer, ByVal periodicite As Integer) As String
    Set c3 = Sheets("CLIENTS")
    periodicite = c3.Range("X" & k)
    Select Case periodicite
        Case Is = 12
            dom_type = "DOMICILIATION MENSUELLE"
        Case Is = 6
            dom_type = "DOMICILIATION SEMESTRIELLE"
        Case Is = 4
            dom_type = "DOMICILIATION TRIMESTRIELLE"
        Case Is = 3
            dom_type = "DOMICILIATION QUADRIMESTRIELLE"
        Case Is = 2
            dom_type = "DOMICILIATION BIMENSUELLE"
        Case Is = 1
            dom_type = "DOMICILIATION ANNUELLE"
    End Select
End Function
Private Sub clean_Modele1()
    Sheets("Modele1").Range("B13", "J30").ClearContents
    Sheets("Modele1").Range("B13", "J30").Borders.Value = 0
    Sheets("Modele1").Range("D1", "J10").ClearContents
End Sub
Public Function reset_buffer_commandes()
    Set s5 = Sheets("Buffer commandes")
    s5.Activate
    nbrowmax = s5.Range("A65000").End(xlUp).Row
    For i = nbrowmax To 2 Step -1
        Rows(i).EntireRow.Delete
    Next i
    Set s1 = Sheets("Tarifs labos")
    s1.Activate
End Function
Public Sub filter_travaux(ByVal Societe As String, ByVal Smois As String, ByRef Armatches() As String)
    Set c2 = Sheets("Travaux")
    Set c7 = Sheets("Buff2")
    Dim crit1 As String
    c2.Range("J1") = Societe
    crit1 = c2.Range("J1").Value
    Dim iarr As Integer
    Call deactivate_C2_filters
    Call erase_buff2(nbrowmin, nbrowmax)
    c2.Activate
    Dim Sht As Worksheet
    crit1 = Societe
    nbrowmax = c2.Range("A65535").End(xlUp).Row
    With Sheets("Travaux").Range("$A$2:$H" & nbrowmax)
        c2.Range("$A$2:$H$65535").AutoFilter field:=2, Criteria1:= _
                  crit1
        c2.Range("$A$2:$H" & nbrowmax).AutoFilter field:=7, Criteria1:=Smois '
        nbrowmax = c2.Range("A65535").End(xlUp).Row
        nbrowmin = NbLignesFiltrées(Worksheets("Travaux").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
        DateSup = DateSerial(Year(Date), Month(Date) - 1, Day(Date))
        dateInf = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
        '''     c2.[A1].AutoFilter field:=8, Criteria1:=">" & CDbl(dateInf) _
                                                         '''     , Operator:=xlAnd, Criteria2:="<=" & CDbl(DateSup)             '  ##### 2 criteres

        c2.[A1].AutoFilter field:=8, Criteria1:=">" & CDbl(dateInf)    '  ##### 2 criteres

        nbrowmax = c2.Range("A65535").End(xlUp).Row
        If nbrowmax > 2 Then
            nbrowmin = NbLignesFiltrées(Worksheets("Travaux").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
            Call copy_range_to_buff(Filter_Start, Filter_End)
        End If
    End With
    Call deactivate_C2_filters
End Sub
Sub filtre2()
    dateInf = DateSerial(Year(Date), Month(Date) - 3, Day(Date))
    Sheets(1).[A1].AutoFilter field:=1, Criteria1:=">" & CDbl(dateInf), Operator:=xlAnd, _
                                                    Criteria2:="<=" & CDbl(Date)
End Sub
Public Sub copy_range_to_buff(ByVal Filter_Start As Integer, ByVal Filter_End As Integer)
    Application.CutCopyMode = False
    Set c2 = Sheets("Travaux")
    Set c7 = Sheets("Buff2")
    c2.Activate
    nbcol_s = 8 ' c2.Range("A1").End(xlRight).Column
    nbcol_d = 8 ' c3.Range("A65000").End(xlRight).Column
    res = Filter_End - Filter_Start + 1
    If res = 1 Then
        source_range = c2.Range(Cells(Filter_Start, 1), Cells(Filter_End, nbcol_s)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Else
        source_range = c2.Range(Cells(Filter_Start, 1), Cells(Filter_End, nbcol_s)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
    c2.Range(source_range).Copy
    c7.Activate
    c7.Range("A1").Select
    dest_range = Range(Cells(1, 1), Cells(1, nbcol_d)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    With dest_range
        ' copie valeurs dans range
        c7.Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                 xlNone, SkipBlanks:=False, Transpose:=False
    End With

End Sub
Public Sub erase_buff2(ByVal nbrowmin As Integer, ByVal nbrowmax As Integer)
    Set c7 = Sheets("Buff2")
    nbrowmax = c7.Range("B65535").End(xlUp).Row
    For i = nbrowmax To 1 Step -1
        c7.Rows(i).EntireRow.Delete
    Next i

End Sub
Function NbLignesFiltrées(ByVal sRange As String) As Long
    ' Calcule le nombre de lignes d'un range avec des lignes contiguës ou non
    Dim i As Long, iLigDep As Long, iLigArr As Long
    Dim ro, visiblerang
    Dim tab_intermédiaire As Variant, tab_inter2 As Variant
    NbLignesFiltrées = 0
    tab_intermédiaire = Split(sRange, ",")
    sstr1 = c3.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = 0 To UBound(tab_intermédiaire)
        tab_inter2 = Split(tab_intermédiaire(i), ":")
        iLigDep = Split(tab_inter2(0), "$")(2)
        iLigArr = Split(tab_inter2(1), "$")(2)
        '
        NbLignesFiltrées = NbLignesFiltrées + iLigArr - iLigDep + 1
    Next i
    Filter_Start = iLigDep
    Filter_End = iLigArr
End Function
Public Sub FilterAnCopyData()
    Dim td As Range, a, b, d
    Set c2 = Sheets("Travaux")
    Set c7 = Sheets("Buff2")
    Set td = c2.Range("$A$2:$H$65535")
    c2.AutoFilter.ShowAllData
    d = c2.Range("J1").Value
    a = Application.Match(c2.Cells(2, 2).Value, td.Columns(1), 0)
    b = Application.Match(c2.Cells(2, 3).Value, td.Columns(3), 0)

End Sub '        If Not IsError(a) And Not IsError(b) Then
'            With td.ListObject.Range
c2.AutoFilter field:=2, Criteria1:=d
c2.AutoFilter field:=3, Criteria1:=ws.Cells(2, 3).Value
'            End With
'            With Range("Résultat").ListObject
'                If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
td.CurrentRegion.offset(1).Copy c2.Range.Cells(2, 1)
'            End With
td.ListObject.AutoFilter.ShowAllData
'        End If
End Sub
Sub FilterOnCellValue()
    Set c2 = Sheets("Travaux")
    Set c7 = Sheets("Buff2")
    Dim crit1 As String
    '    crit1 = c2.Range("J1").Value
    crit1 = "ALPHA CONSTRUCTION"
    With Sheets("Travaux").Range("A1:H10000")
        '         c2.AutoFilter.ShowAllData
        .AutoFilter field:=2, Criteria1:=crit1
        '    .AutoFilter Field:=23, Criteria1:=Sheets("ControlPlanning").Range("C4").Value
    End With
End Sub
Public Sub deactivate_C2_filters()
    c2.Activate
    c2.AutoFilterMode = False
    ' Worksheets("CLIENTS").Activate
End Sub
