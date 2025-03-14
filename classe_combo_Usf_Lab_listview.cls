VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_Lab_listview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeCombo As MSForms.ComboBox
Attribute GroupeCombo.VB_VarHelpID = -1
Private Sub GroupeCombo_click()
    combomem = GroupeCombo.Name
    cmbrow = Right(combomem, Len(combomem) - 8)
    Set c3 = Sheets("CLIENTS")
    If GroupeCombo.text = "" Then
        Exit Sub
    ElseIf (cmbrow = 2) Then
    Call tri_col_generic(Sheets("Travaux"), 2)
        cle_rech = Usf_Lab_listview("ComboBox2").Value
            nbrowmax = Sheets("Travaux").Range("B65000").End(xlUp).Row
            If FindAll(cle_rech, Sheets("Travaux"), "B2:B" & nbrowmax, Armatches()) Then
               Call Usf_Lab_listview.affiche1
               Call Usf_Lab_listview.affiche3
                trig = False
                Else
                Call Usf_Lab_listview.affiche3
'            End If
          End If

    End If
End Sub





