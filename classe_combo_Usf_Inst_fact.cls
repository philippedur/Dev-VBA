VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_Inst_fact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeCombo As MSForms.ComboBox
Attribute GroupeCombo.VB_VarHelpID = -1

Private Sub GroupeCombo_click()
    combomem = GroupeCombo.Name
    cmbrow = Right(combomem, Len(combomem) - 8)
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    If GroupeCombo.text = "" Then
        Exit Sub
        '    ElseIf cmbrow = 1 Then
        ''        Call tri_col_generic(Sheets("CLIENTS"), 6)
        '        nbrowmax = c3.Range("F65000").End(xlUp).Row
        '        cle_rech = USF_newjob("ComboBox1").Value
        '        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "F2:F" & nbrowmax, Armatches()) Then
        '            ligne = pos2
        ''            Call USF_newjob.affiche1
        '            USF_newjob("ComboBox2").Value = c3.Cells(pos, 14)
        '            trig = False
        '        End If
    ElseIf (cmbrow = 2) Then
        Call tri_col_generic(Sheets("CLIENTS"), 14)
        cle_rech = USF_Inst_fact("ComboBox2").Value
        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            k = pos
            cle_rech2 = CStr(Sheets("Travaux").Range("A" & pos))
            '            Call tri_col_generic(Sheets("Travaux"), 1)
            Set c2 = Sheets("Travaux")
            nbrowmax = Sheets("Travaux").Range("B65000").End(xlUp).Row
            If FindAll(cle_rech, Sheets("Travaux"), "B2:B" & nbrowmax, Armatches()) Then
                '              USF_Inst_fact.ListBox2.RowSource = ""  ' COMMENTER AVEC USAGE LISTVIEW..... COMMENTER AVEC USAGE LISTVIEW..... COMMENTER AVEC USAGE LISTVIEW.....
                USF_Inst_fact.ListBox2.RowSource = ""
                no_record = False
                trig = False
                Call USF_Inst_fact.affiche1
            Else
                no_record = True
                USF_Inst_fact.ListBox2.RowSource = ""
                Call USF_Inst_fact.affiche1
            End If
        End If
    ElseIf (cmbrow = 3) Then
                no_record = True
                USF_Inst_fact.ListBox2.RowSource = ""
                Call USF_Inst_fact.affiche1
End If
    
End Sub





