VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Classe_USF_simul_gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Groupelbl As MSForms.Label
Attribute Groupelbl.VB_VarHelpID = -1
'Public Sub Groupelbl_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Sub Groupelbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cle_rech3 = Groupelbl.Name
    sstr1 = Groupelbl.Name
    sstr2 = 5 + InStr(1, sstr1, "Label")
    sstr3 = Mid(sstr1, sstr2, sstr2 - Len(sstr2))
    tt = Mid(sstr1, sstr2, sstr3)
    offset = 31 - fnd1   '  calcul offset entre 1ere progressbar(31) et fnd1
    j = tt - offset
    If Workbooks("Facturation-auto-mail-MIDI-SERVICES-01.xlsm").Sheets("CLIENTS").Cells.Range("F" & j).Value <> Empty Then
        t1 = Workbooks("Calcul prix et recettes.xlsb").Sheets("Recettes en atelier").Cells.Range("F" & j)                                                    '  t1-Source-name
        t2 = Val(decod_cle6("Facturation-auto-mail-MIDI-SERVICES-01.xlsm").Sheets("CLIENTS").Cells.Range("H" & j).Value)                                      '  t2-Source-value
        If FindAll_OneRec(t1, Workbooks("Stock.xlsm").Worksheets("ProduitsStock"), "B1" & ":B" & nbrowmax, Armatches()) Then    '  t3-Dest-name
            t3 = Workbooks("Stock.xlsm").Sheets("CLIENTS").Cells.Range("A" & pos).Value                                        '  t3-Dest-value
            t4 = Workbooks("Stock.xlsm").Sheets("CLIENTS").Cells.Range("A" & pos).Value
        End If
    End If
    If ((t2 > 0) And (t3 > 0)) Then
        simul_gestion("Textbox5").text = t1 & " ---------------   " & Round(t3 / t2) & "  Recettes possibles"
    Else
        simul_gestion("Textbox5").text = t1 & "  Recettes possibles"
    End If

End Sub
