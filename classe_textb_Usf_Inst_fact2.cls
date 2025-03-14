VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_textb_Usf_Inst_fact2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeTextb As MSForms.TextBox
Attribute GroupeTextb.VB_VarHelpID = -1
'Public Sub GroupeTextb_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub Groupetextb_change()
    TextMem = GroupeTextb.Name
    txtrow = Right(TextMem, Len(TextMem) - 7)
    If GroupeTextb.text = "" Then
        Exit Sub
    Else
        If no_record = True Then
        Msg = "Il n'y a aucun travaux pour ce client.." & vbCrLf & "Voulez-vous editer la facture domiciliation simple ?" & vbCrLf & "Vous pouvez choisir le mois d'echeance avec le selecteur de droite, puis validez."    ' Define message.
        style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Impression de la facture simple"
        Help = ""
        Ctxt = 1000
        response = MsgBox(Msg, style, Title, Help, Ctxt)
        If response = vbYes Then
            mystring = "Yes"
        End If
        Else
        c2.Activate
        nbrowmax = Worksheets("Travaux").Range("A65000").End(xlUp).Row
        Selection.AutoFilter
        c2.Range("$A$1:$H$" & nbrowmax).AutoFilter field:=2, Criteria1:= _
        cle_rech
        nbrowmax = Worksheets("Travaux").Range("A65000").End(xlUp).Row
        nbcol_s = 10 ' c2.Range("A1").End(xlRight).Column
        nbcol_d = 10 ' c3.Range("A65000").End(xlRight).Column
        source_range = c2.Range(Cells(1, 1), Cells(nbrowmax, nbcol_s)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Worksheets("Travaux").Range(source_range).Copy
        Set c10 = Sheets("Buff3")
        c10.Activate
        c10.Range("A1").Select
        dest_range = Range(Cells(1, 1), Cells(nbrowmax, nbcol_d)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        With dest_range
            ' copie valeurs dans range
            Worksheets("Buff3").Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        End With
        nbrowmax2 = Worksheets("Buff3").Range("A65000").End(xlUp).Row
        Set plage2 = Worksheets("Buff3").Range("A" & 1 & ":H" & nbrowmax2)
        USF_Inst_fact.ListBox2.RowSource = plage2.Address(External:=True)
        Call USF_Inst_fact.affiche3
        Range("A1").Select
        Call deactivate_all_filters
        nbrowmax2 = Worksheets("buff3").Range("A65000").End(xlUp).Row
             
        Msg = "Selectionnez les travaux avec les cases a cocher et valider." & vbCrLf & "En cas de dates correspondant a des mois differents," & vbCrLf & "L'échéance de la facture portera le mois du 1er travail selectionné"    ' Define message.
        style = vbYesNo + vbDefaultButton2
        Title = "Impression de l'échéance :" & USF_Inst_fact("Combobox3").Value
        Help = ""
        Ctxt = 1000
        response = MsgBox(Msg, style, Title, Help, Ctxt)
        If response = vbYes Then
            mystring = "Yes"
        End If
        End If
    End If
    '        If txtrow = 11 Then
    ''        USF_newjob("TextBox11").Enabled = False
    ''            Call USF_newjob.affiche1
    '            trig = False
    '        End If
End Sub







