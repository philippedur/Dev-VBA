VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_Eff_client"
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
    ElseIf (cmbrow = 2) Then
'        Call tri_col_generic(Sheets("CLIENTS"), 14)
        nbrowmax = c3.Range("N65000").End(xlUp).Row
        cle_rech = USF_Eff_client("ComboBox2").Value
        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            ligne = pos
            Call USF_Eff_client.affiche1
            trig = False
        End If
    End If
End Sub




