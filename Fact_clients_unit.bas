Attribute VB_Name = "Fact_clients_unit"
Option Explicit
'Private Declare ptrsafe Function FindWindow Lib "User32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    Global Const HWND_TOPMOST = -1
    Global Const SWP_NOACTIVATE = &H10
    Global Const SWP_SHOWWINDOW = &H40
    Public Function Facture_clients_unitaire(ByRef sh As Worksheet, ByVal k As Integer)
    Call init_rep2
    lnk = Path2 & "Facturation-auto-mail-MIDI-SERVICES-01.xlsm"
    '    Workbooks.Open Filename:=lnk
    Dim bb, cc, NumFactAb, NumFactAn, comp As Integer
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Call RAZ
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
        Sheets("modele1").Range("H11") = UCase(Format(USF_Inst_fact("ComboBox3").Value, "mmmm"))
        Sheets("modele1").Range("H11").HorizontalAlignment = xlHAlignCenter
        Sheets("modele1").Range("H11").Font.Bold = True
        Sheets("modele1").Range("H11").Font.Size = 12
        '        For k = k To k
        periodicite = c3.Range("X" & k)
        date_creation = c3.Range("D" & k)
        Societe = c3.Range("N" & k)
        If date_creation <> 0 And (periodicite <> 0) Then
            res = (calc_period2(Societe, k, Date, periodicite, date_creation))
            '''                If trig_fact = True Then
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
            '                    Call etat_gene
            .[date_facture] = CStr(Format(Date, "mm/dd/yy"))
            .[date_facture].Font.Size = 11
            .[date_facture].Font.Name = "Calibri"
            If no_record = True Then .[echeance] = UCase(Format(USF_Inst_fact("ComboBox3").Value, "mmmm"))   ' c3.Range("E" & k)
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
            For i = 12 To 32
                total = total + Val(c1.Cells(i, 8))
            Next i
            .[Total_HT] = total
            .[Total_HT].Font.Size = 11
            .[Total_HT].Font.Name = "Calibri"
            .[TVA_20] = total / 100 * 20
            .[TVA_20].Font.Size = 11
            .[TVA_20].Font.Name = "Calibri"
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
            '''                End If
        End If
        '        Next k
    End With
    Call RAZ
    '        ActiveWorkbook.Close
End Function
Private Sub RAZ()
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
        '            .[Prix_HT] = tarif(c3.Range("R" & k)) * c2.Range("D" & ArMatches(i))
        '            .[Prix_HT].Font.Size = 11
        '            .[Prix_HT].Font.Name = "Calibri"
        '        .[PU_HT] = ""
        '        .[Total_HT] = 0
        '        .[Total_TVA] = 0
        offset = 13
        c1.Cells(offset, 2) = ""
        c1.Cells(offset, 3) = ""
        c1.Cells(offset, 7) = ""
        c1.Cells(offset, 6) = ""
        c1.Cells(offset, 8) = ""
        For i = 1 To 9
            c1.Cells(offset + i, 2) = ""
            c1.Cells(offset + i, 3) = ""
            c1.Cells(offset + i, 7) = ""
            c1.Cells(offset + i, 6) = ""
            c1.Cells(offset + i, 8) = ""
        Next i
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
    End Select
End Function
Private Sub calcul_travaux()
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c10 = Sheets("Buff3")
    ind_fact = 1
    With c1
        offset = 13 ' 15
        c1.Cells(offset, 2) = "DOM"
        c1.Cells(offset, 3) = dom_type(k, periodicite)
        c1.Cells(offset, 7) = ind_fact    ' 1
        c1.Cells(offset, 6) = Worksheets("CLIENTS").Range("S" & k).Value
        c1.Cells(offset, 8) = Worksheets("CLIENTS").Range("S" & k).Value * ind_fact

        For i = 1 To 20 ' (36 - offset)
            c1.Cells(offset + i, 2) = ""
            c1.Cells(offset + i, 3) = ""
            c1.Cells(offset + i, 7) = ""
            c1.Cells(offset + i, 6) = ""
            c1.Cells(offset + i, 8) = ""
        Next i
        If no_record = False Then
            If FindAll(sstr1, Sheets("Buff3"), "B2:B2000", Armatches()) <> Empty Then
                .[echeance] = UCase(Format(c10.Range("H" & Armatches(1)), "MMMM")) ' c3.Range("E" & k)
                sstr4 = c2.Range("G" & UBound(Armatches()))
                '            Call filter_travaux(sstr1)
                nbrowmax = c2.Range("B2").End(xlUp).Row
                If UBound(Armatches()) > 1 Then j = 1
                c1.Cells(offset + j, 2) = "TRAV"
                c1.Cells(offset + j, 3) = "TRAVAUX ADDITIONNELS DIVERS"
                j = 2
                Set c10 = Sheets("Buff3")
                For i = 1 To UBound(Armatches())
''                    If (Year(c2.Cells(Armatches(i), 8)) = Year(Date)) And (Sheets("Travaux").Cells(Armatches(i), 7) = comp_mois_rev(k))) Then
                    If USF_Inst_fact.ListBox2.Selected(i) = True Then
                        c1.Cells(offset + j, 2) = c2.Range("F" & Armatches(i))
                        c1.Cells(offset + j, 3) = c2.Range("C" & Armatches(i))
                        c1.Cells(offset + j, 7) = c2.Range("D" & Armatches(i))
                        c1.Cells(offset + j, 6) = c2.Range("E" & Armatches(i))
                        c1.Cells(offset + j, 8) = c2.Range("D" & Armatches(i)) * (c2.Range("E" & Armatches(i)))
'''                        Debug.Print c2.Cells(Armatches(i), 2), c2.Cells(Armatches(i), 3), c2.Cells(Armatches(i), 4), c2.Cells(Armatches(i), 5), c2.Cells(Armatches(i), 6), c2.Cells(Armatches(i), 7)
                        j = j + 1
                    End If
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
    USF61b.Label8.Caption = "Expédition terminée."
    USF61b.Label8.Visible = False
    USF61b.Label8.ForeColor = RGB(15, 5, 107)
    USF61b.Label8.Font.Size = 11
    USF61b.Label8.Font.Name = "Calibri"
End Sub

Public Sub Print_to_DefaultPrinter(name_pdf)
    If printer_enable = True Then
        Call TstPdfCreator
    End If
End Sub








