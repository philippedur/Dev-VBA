VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_beforedoubleclick(ByVal Target As Range, Cancel As Boolean)
    Dim Chemin As String
    Dim NomRep As String
    Set Dossier = CreateObject("Scripting.FileSystemObject")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_trav")
    Set c5 = ThisWorkbook.Worksheets("TYP_dom")
    nbrowmax = c5.Range("D65000").End(xlUp).Row
    If (Target.Row = 1) Then
        Set titre = [A1:S1]
        ordretri = IIf(Target.Interior.colorIndex = 3, xlDescending, xlAscending)
        Target.CurrentRegion.Sort Key1:=Cells(1, Target.Column), Order1:=ordretri, header:=xlYes
        xx = IIf(Target.Interior.colorIndex = 2, 2, 17)
        titre.Interior.colorIndex = 2
        Target.Interior.colorIndex = 17
    ElseIf Target.Row > 1 And (Target.Column = 19) Then
        Lig = c3.Cells(Rows.Count, 1).End(xlUp).Row
        Set range1 = c5.Range("D2:D" & nbrowmax)
        Set rng = c5.Range("D2")
'        liste = tab1_liste_tarifs(list_rech)
        With Target.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="='" & c5.Name & "'!" & range1.Address ' :=list_rech
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
    '        Target.Font.FontStyle = 1
        c3.Range("S:S").NumberFormat = " 0.00"
        cle_rech = Val(c3.Cells(Target.Row, Target.Column))
    ElseIf Target.Row > 1 And (Target.Column = 26) Then
'        Path4 = "D:\Dev-VBA\Documents clients\"
'        NomRep = Dir(Chemin & NomRep, vbDirectory)
           sstr1 = c3.Range("Z" & Target.Row)
           len_sstr1 = Len(c3.Range("Z" & Target.Row))
           fich = Trim(Mid(sstr1, 6, len_sstr1))
           Call Cdk
'''        Shell "C:\windows\explorer.exe /e," & Chemin & fich, vbNormalFocus
    Else
        Exit Sub
    End If
End Sub
Public Sub Cdk()
Dim Fichier$, Niveau$, sPath$, Chemin$
With Application.FileDialog(msoFileDialogFilePicker)
    Call init_rep2
    sPath$ = Path4
    .AllowMultiSelect = False
    .InitialFileName = sPath$ & fich
    .Show
    If .SelectedItems.Count Then
        sPath = .SelectedItems(1)
        Niveau = (UBound(Split(sPath, "\")))
        Fichier = Split(sPath, "\")(Niveau)
        Chemin = Left(sPath, Len(sPath) - Len(Fichier))
    End If
'''CreateObject("Shell.Application").Open ("C:\MonDossier\MonDocument.docx")
End With
End Sub
Public Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Target.Row > 1 And (Target.Column = 19) Then
        Set c3 = Sheets("CLIENTS")
        Set c5 = Sheets("TYP_dom")
        nbrowmax = c5.Range("D65000").End(xlUp).Row
        cle_rech = Format(c3.Cells(Target.Row, Target.Column), " 0.00")
        If Find_Tarif(cle_rech, Sheets("TYP_dom"), "D2" & ":D" & nbrowmax, Armatches()) Then
            c3.Cells(Target.Row, 18) = c5.Cells(pos, 1)
            c3.Cells(Target.Row, 19).NumberFormat = "# ##0.00__€"
'            Selection.NumberFormat = "#,##0.00__€"
'            c3.Range("S:S") = Replace(Range("S:S"), ".", ",")
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
        sstr1 = CStr(CStr(Sheets("TYP_dom").Cells(i, 4)))
        sstr2 = CStr(Sheets("TYP_dom").Cells(i, 1))
        list_rech = list_rech & Format(sstr1, " 0.00") & ","
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


