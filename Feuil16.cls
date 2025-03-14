VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_beforedoubleclick(ByVal Target As Range, Cancel As Boolean)
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_trav")
    If (Target.Row = 1) Then
        Set titre = [A1:S1]
        ordretri = IIf(Target.Interior.colorIndex = 3, xlDescending, xlAscending)
        Target.CurrentRegion.Sort Key1:=Cells(1, Target.Column), Order1:=ordretri, header:=xlYes
        xx = IIf(Target.Interior.colorIndex = 2, 2, 17)
        titre.Interior.colorIndex = 2
        Target.Interior.colorIndex = 17
    ElseIf Target.Row > 1 And (Target.Column = 19) Then
        Lig = c3.Cells(Rows.Count, 1).End(xlUp).Row
        liste = tab1_liste_tarifs(list_rech)
        With Target.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=list_rech
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
'        Target.Font.FontStyle = 1
        nbrowmax = c5.Range("D65000").End(xlUp).Row
        c3.Range("S:S").NumberFormat = "###0.00 €"
        cle_rech = Format(c3.Cells(Target.Row, Target.Column), "###0.00")
        If Find_Tarif(cle_rech, Sheets("TYP_dom"), "D2" & ":D" & nbrowmax, Armatches()) Then
            Cells(Target.Row, 18) = CStr(Sheets("TYP_dom").Cells(pos, 1))
            Application.EnableEvents = True
        End If
    Else
        Exit Sub
    End If
End Sub
Public Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
    If Target.Row > 1 And (Target.Column = 19) Then
        Set c3 = Sheets("CLIENTS")
        Set c5 = Sheets("TYP_dom")
        nbrowmax = c5.Range("D65000").End(xlUp).Row
        cle_rech = Format(c3.Cells(Target.Row, Target.Column), "###0.00")
        If Find_Tarif(cle_rech, Sheets("TYP_dom"), "D2" & ":D" & nbrowmax, Armatches()) Then
            c3.Cells(Target.Row, 18) = c5.Cells(pos, 1)
'            c3.Cells(Target.Row, 19) = Format(CStr(c5.Cells(pos, 4)), "#.## [0.00 €]"
            c3.Range("S:S").NumberFormat = "###0.00 €"
            c3.Cells(Target.Row, 19).HorizontalAlignment = xlCenter
                End If
    End If
End Sub
Public Function tab1_liste_tarifs(list_rech)
    Set c3 = Sheets("CLIENTS")
    Set c5 = Sheets("TYP_dom")
    nbrowmax = c5.Range("D65000").End(xlUp).Row
    '    cle_rech = c3.Cells(Target.Row, Target.Column)
    '    If FindAll(cle_rech, Sheets("TYP_dom"), "D1" & ":D" & nbrowmax, ArMatches()) Then
    list_rech = ""    ' ENTETE pour permettre de comparer avec une chaine vide
    For i = 2 To nbrowmax
        sstr1 = CStr(CStr(Sheets("TYP_dom").Cells(i, 4))) & " €"
        sstr2 = CStr(Sheets("TYP_dom").Cells(i, 1))
        list_rech = list_rech & Format(sstr1, "###0.00 €") & ","
    Next i
End Function
Public Function tab3_rech_typ_dom() As Boolean
    Set c5 = Sheets("TYP_dom")
    nbrowmax = c5.Range("D65000").End(xlUp).Row
    cle_rech = Trim(Mid(c3.Cells(i, 19), 1, 2))
    If FindAll_OneRec(cle_rech, Sheets("TYP_dom"), "D1" & ":D" & nbrowmax, Armatches()) Then
        tab3_rech_typ_dom = True
    End If
End Function
Sub tri()
    Range("A3:J502").Select
    Selection.AutoFilter
    Selection.AutoFilter field:=1, Criteria1:="=" & Range("E1").text & "*"
End Sub
'Public Sub mesure_BCDE()
'    cle_rech = CStr(Trim(USF66.ComboBox1.Value))
'    Workbooks("Calcul prix et recettes.xlsb").Activate
'    Set c3 = Sheets("recettes en atelier")
'    nbrowmax = c3.Range("A65000").End(xlUp).row
'    cle_rech = decod_cle(USF66.ComboBox1.Value)
'    If FindSingle(cle_rech, Sheets("Recettes en atelier"), "A2:" & "A" & nbrowmax, Armatches()) Then
'        fnd1 = pos
'    End If
'    If FindSingle("Total", Sheets("Recettes en atelier"), "A" & fnd1 + 1 & ":A" & nbrowmax, Armatches()) Then
'        fnd2 = pos
'    End If
'    If FindSingle("Total", Sheets("Recettes en atelier"), "A" & fnd2 + 1 & ":A" & nbrowmax, Armatches()) Then
'        fnd3 = pos
'    End If
'End Sub
'
''                   If EstForm(c3.Cells(ligne, 6)) Then Me("Combobox7-" & j).Value = ""  ' fonction test presence formules dans table recettes









