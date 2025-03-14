VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_combo_Usf_newclient"
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
    ElseIf cmbrow = 20 Then
'        cle_rech = Format(c3.Cells(Target.Row, Target.Column), "###0.00")
        Lig = USF_NewClient("ComboBox20").ListIndex + 2
        USF_NewClient("Textbox18").Value = c5.Cells(Lig, 1)
    End If
    
End Sub



