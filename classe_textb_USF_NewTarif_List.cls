VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_textb_USF_NewTarif_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeTextb As MSForms.TextBox
Attribute GroupeTextb.VB_VarHelpID = -1
Public Sub Groupetextb_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If USF_NewTarif_List.Frame1.ActiveControl.Name <> "" Then
        sstr1 = USF_NewTarif_List.Frame1.ActiveControl.Name
        Set c5 = Sheets("TYP_dom")
        sstr2 = InStr(1, sstr1, "-")
        sstr3 = (Len(sstr1) - sstr2)
        tt = Mid(sstr1, sstr2 + 1, sstr3)
        nbrowmax = Sheets("TYP_dom").Range("B65000").End(xlUp).Row
        For L = 1 To 5
            c = GetLongFromRGB(192, 224, 192)
        Next L
        res1 = USF_NewTarif_List.Frame1("Textbox" & 1 & "-" & tt)
        res2 = USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt)
        If (Val(tt) < 3) Then
            sstr4 = InputBox("Nouvelle valeur du tarif " & CStr(res1) & "  à " & res2 & " ?" & vbCrLf & "(Doit être comprise entre " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt), " 0.00") & " €" & " et " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt + 1), " 0.00") & " €", "Euros")
        ElseIf Val(tt) = nbrowmax Then
            sstr4 = InputBox("Nouvelle valeur du tarif " & CStr(res1) & "  à " & res2 & " ?" & vbCrLf & "(Doit être comprise entre " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt - 1), " 0.00") & " €" & " et " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt), " 0.00") & " €", "Euros")
        Else
            sstr4 = InputBox("Nouvelle valeur du tarif " & CStr(res1) & "  à " & res2 & " ?" & vbCrLf & "(Doit être comprise entre " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt - 1), " 0.00") & " €" & " et " & _
                             Format(USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt + 1), " 0.00") & " €", "Euros")
        End If
        If sstr4 <> "" Then
            c5.Activate
            Worksheets("TYP_dom").Range("D" & tt) = IIf(sstr4 <> "", sstr4, "")
            Worksheets("TYP_dom").Range("E" & tt) = sstr5
            USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & tt).Value = Worksheets("TYP_dom").Range("D" & tt)
''            USF_NewTarif_List.Frame1("Textbox" & 5 & "-" & tt).Value = Worksheets("TYP_dom").Range("E" & tt)
            Msg = "Voulez-vous mettre à jour ce tarif dans tout le fichier client ?"
            style = vbYesNo + vbCritical + vbDefaultButton2
            Title = "Changement d'un tarif existant."
            Help = ""
            Ctxt = 1000
            response = MsgBox(Msg, style, Title, Help, Ctxt)
            If response = vbYes Then
            Call maj_tarifs
            ElseIf response = vbNo Then
                mystring = "No"
            End If
            sstr5 = InputBox("Voulez-vous modifier le champ commentaire ?" & vbCrLf & "Pour mettre un champ vide, tapez entrée." _
            & vbCrLf & "Pour valider une valeur clickez sur le bouton OK.")
            For L = 1 To 5
                USF_NewTarif_List.Frame1("Textbox" & L & "-" & tt).BackColor = GetLongFromRGB(192, 224, 192)
                Worksheets("TYP_dom").Range("E" & tt) = sstr5
            Next L
        ElseIf (sstr4 = "") Then
            For L = 1 To 5
                USF_NewTarif_List.Frame1("Textbox" & L & "-" & tt).BackColor = GetLongFromRGB(192, 224, 192)
                Worksheets("TYP_dom").Range("E" & tt) = sstr5
            Next L
            sstr5 = InputBox("Voulez-vous modifier le champ commentaire ?" & vbCrLf & "Pour mettre un champ vide, tapez entrée." _
            & vbCrLf & "Pour valider une valeur clickez sur le bouton OK.")
        End If
        For L = 1 To 5
            USF_NewTarif_List.Frame1("Textbox" & L & "-" & tt).BackColor = GetLongFromRGB(255, 255, 255)
        Next L
    '                    USF_NewTarif_List.Frame1("Textbox" & 9 & "-" & tt).Value = ""
    '                    USF_NewTarif_List.Frame1("Textbox" & 10 & "-" & tt).Value = ""
    Unload USF_NewTarif_List
    USF_NewTarif_List.Show
    End If
End Sub
