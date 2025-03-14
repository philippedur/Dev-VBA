VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeCombo As MSForms.ComboBox
Attribute GroupeCombo.VB_VarHelpID = -1
Private Sub GroupeCombo_click()
    combomem = GroupeCombo.Name
    cmbrow = Right(combomem, Len(combomem) - 8)
    If new_row = True Then
        If cmbrow = 24 Then
            USF_NewClient("Textbox20").Value = Left(CStr(USF_NewClient("ComboBox24")), 1)
        Else
            nbrowmax = c3.Range("N65000").End(xlUp).Row
            cle_rech2 = GroupeCombo.text
        End If
    End If
    If GroupeCombo.text = "" Then
        Exit Sub
    ElseIf cmbrow = 1 Then
        Call tri_col_generic(Sheets("CLIENTS"), 6)
        cle_rech = USF_Client("ComboBox1").Value
        Set rng = c3.Range("F2" & ":F2000")
        If FindAll_ByArea(rng, cle_rech) Then
'        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "F2:F" & nbrowmax, Armatches()) Then
            ligne = pos
            Call USF_Client.affiche
            trig = False
        End If
    ElseIf cmbrow = 2 Then
        Call tri_col_generic(Sheets("CLIENTS"), 14)
        cle_rech = USF_Client("ComboBox2").Value
        Set rng = c3.Range("N2" & ":N2000")
        If FindAll_ByArea(rng, cle_rech) Then
'''        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            ligne = pos
            Call USF_Client.affiche
            trig = False
        End If
    End If
End Sub


