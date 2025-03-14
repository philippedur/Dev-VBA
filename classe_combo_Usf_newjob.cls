VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_newjob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeCombo2 As MSForms.ComboBox
Attribute GroupeCombo2.VB_VarHelpID = -1
Private Sub GroupeCombo2_click()
    combomem = GroupeCombo2.Name
    cmbrow = Right(combomem, Len(combomem) - 8)
        Set c3 = Sheets("CLIENTS")
        Call tri_col_generic(Sheets("CLIENTS"), 14)
    If GroupeCombo2.text = "" Then
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
'        Call tri_col_generic(Sheets("CLIENTS"), 14)
        nbrowmax = c3.Range("N65000").End(xlUp).Row
        cle_rech = USF_newjob("ComboBox2").Value
        Set rng = c3.Range("N" & k & ":N2000")
        If FindAll_ByArea(rng, cle_rech) Then
'        If FindAll_ByArea(ByVal cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            ligne = pos
             pos2 = pos
            Call USF_newjob.affiche1
            USF_newjob("TextBox1").Value = Sheets("CLIENTS").Cells(pos, 6)
            trig = False
        End If
    ElseIf (cmbrow = 6) Then
    selct = True
        USF_newjob("ComboBox10").Enabled = False
        USF_newjob("TextBox9").Enabled = False
        USF_newjob("TextBox11").Enabled = False
        cle_rech = Left(USF_newjob("ComboBox6").Value, 10)
'        If FindAll_OneRec(cle_rech, Sheets("Typ_trav"), "B2:B" & nbrowmax, Armatches()) Then
        Lig = USF_newjob("ComboBox6").ListIndex + 2
'        End If
    ElseIf (cmbrow = 10) Then
    selct = False
        USF_newjob("ComboBox6").Enabled = False
        USF_newjob("TextBox5").Enabled = False
        Lig = USF_newjob("ComboBox10").ListIndex + 2
'        End If
    End If
End Sub





