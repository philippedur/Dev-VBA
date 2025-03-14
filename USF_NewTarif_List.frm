VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_NewTarif_List 
   Caption         =   "USF_NewTarif_List"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   OleObjectBlob   =   "USF_NewTarif_List.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_NewTarif_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    ' USF_NewTarif_List !!!!!!!!!!!!!!!!!!!
Private Declare PtrSafe Function FindWindowA Lib "User32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public max As Integer
Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Ht As Integer
Dim chaine As String
'Private WithEvents USF As Classe1
Private Type SYSTEMTIME
    xAnnee As Integer
    xMois As Integer
    xJourSemaine As Integer
    xJour As Integer
    xHeure As Integer
    xMinute As Integer
    xSeconde As Integer
    xMilliseconde As Integer
End Type
Dim txt() As New classe_textb_USF_NewTarif_List
Dim Cmb() As New classe_combo_USF_NewTarif_List
Dim Couleur As Long
Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        TextBox1.BackColor = choixColor(Me, TextBox1.BackColor)
    End If
End Sub
'Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
'Private Declare Function GetTickCount Lib "Kernel32" () As Long
'Private Declare Sub GetSystemTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)
Private Sub Frame1_Click()
    USF_NewTarif_List.Frame1.Visible = True
End Sub
Private Sub Supp_Inv_Click()
'    USF_NewTarif_List.Show
End Sub
'Fermeture du UserForm
Private Sub USF_Command_List_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set USF_NewTarif_List = Nothing
End Sub
Private Sub New_pharmacie_Click()
    USF_NewClient2.Show
End Sub
Private Sub NewTarif_Click()
    new_row = False
    trigged = False
    Dim Msg, style, Title, Help, Ctxt, response
    res1 = InputBox("Quelle est la valeur de ce nouveau tarif ?")
    If res1 <> "" Then
        res1 = Replace(res1, ".", ",")
        Set c5 = Sheets("TYP_dom")
        Call tri_col_generic(Sheets("TYP_dom"), 4)
        nbrowmax = c5.Range("D65000").End(xlUp).Row
        sstr1 = Worksheets("TYP_dom").Range("C" & nbrowmax)
        max = DMX(Worksheets("TYP_dom").Range("C" & nbrowmax))
        Msg = "Ajouter ce tarif ?"
        style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Ajout d'un nouveau tarif."
        Help = ""
        Ctxt = 1000
        response = MsgBox(Msg, style, Title, Help, Ctxt)
        If response = vbYes Then
            Worksheets("TYP_dom").Range("D" & nbrowmax + 1) = res1
            Worksheets("TYP_dom").Range("C" & nbrowmax + 1) = "DM-" & CStr(max + 1)
            Worksheets("TYP_dom").Range("B" & nbrowmax + 1) = "DM-" & CStr(max + 1) & "(Domiciliation Tarif " & max + 1 & ")"
            USF_NewTarif_List.Frame1("Textbox" & 4 & "-" & max).Value = Worksheets("TYP_dom").Range("D" & nbrowmax + 1)
            USF_NewTarif_List.Frame1("Textbox" & 5 & "-" & max).Value = Worksheets("TYP_dom").Range("E" & nbrowmax + 1)
            res2 = InputBox("Voulez-vous modifier le champ commentaire ?" & vbCrLf & "Pour mettre un champ vide, tapez entrée." _
            & vbCrLf & "Pour valider une valeur clickez sur le bouton OK.")
            Worksheets("TYP_dom").Range("E" & nbrowmax + 1) = res2
            Unload USF_NewTarif_List
            USF_NewTarif_List.Show
        ElseIf response = vbNo Then
        End If
    Else
        Exit Sub
    End If
End Sub
Function DMX(res As String) As String
    DMX = 0
    For i = 2 To nbrowmax
        sstr3 = Mid(Worksheets("TYP_dom").Range("C" & i), 4, 3)
        If Val(sstr3) > Val(DMX) Then
            DMX = sstr3
        End If
    Next

End Function
Private Sub Quitter_Click()
    Unload USF_NewTarif_List
    Worksheets("TYP_dom").Activate
    Range("A1").Select
    '    USF61.Show
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Cancel = 0 Then
        Unload Me
    '        End
    End If
End Sub
Private Sub Angle1_Change()
'    AppliqueArrondi
End Sub
Private Sub Angle2_Change()
'    AppliqueArrondi
End Sub
Sub AppliqueTransp()
    ActiveTransparence Me.Caption, True, True, 16706187, 255
    '    ActiveTransparence Me.Caption, True, 50, 0, Me.BackColor, ScrollBar1.Value
End Sub
Sub AppliqueArrondi()
'   Angle1.Value = 60
'   Angle2.Value = 60
    RoundCorners Me, Me.Width, Me.Height, 20, 20
End Sub
Private Sub USF_LostFocus(ByVal Txtbx As String)
    Me.Controls(Txtbx).BackColor = RGB(255, 255, 255)
End Sub
Private Sub Enregistrer_Click()
'    Set s4 = Sheets("Produits")
'    ActiveWorkbook.Save
End Sub
Public Sub Userform_initialize()    '///// PAS DE CHECK DE CELLULES VIDES
'    Call reset_buffer_commandes
    new_row = False
    trigged = False
    Set c4 = Sheets("TYP_dom")
    nbrowmax = c4.Range("B65000").End(xlUp).Row
    max = nbrowmax
    Call tab1_affichage
    Application.ScreenUpdating = True
    Call affiche_headers
    i = 1
    X = 3
    Y = 0

    For i = 2 To max
        ligne = i
        j = 1    ' Tab1(i, 2)
    '            Couleur = GetLongFromRGB(192, 224, 192)
        Couleur = GetLongFromRGB(255, 255, 255)
        Set R_t1 = Me.Frame1.Controls.Add("Forms.TextBox.1", "TextBox1-" & i, True)
        With R_t1
            Couleur = c4.Cells(ligne, 1).Interior.color
            Me.Frame1("Textbox1-" & i).BackColor = Couleur
            Me.Frame1("Textbox1-" & i).text = c4.Cells(ligne, 1)  ' TabTarifs(1, i) ' c4.Cells(ligne, 1)
            Me.Frame1("Textbox1-" & i).FontSize = 10
            Me.Frame1("Textbox1-" & i).Top = Y
            Me.Frame1("Textbox1-" & i).Left = X + 3
            Me.Frame1("Textbox1-" & i).Width = 40
            Me.Frame1("Textbox1-" & i).Visible = True
            Me.Frame1("Textbox1-" & i).SpecialEffect = 0
            Me.Frame1("Textbox1-" & i).BorderStyle = 1
        End With
        Set R_t2 = Me.Frame1.Controls.Add("Forms.TextBox.1", "TextBox2-" & i, True)
        With R_t2
            Couleur = c4.Cells(ligne, 2).Interior.color
            Me.Frame1("textbox2-" & i).BackColor = Couleur
            Me.Frame1("TextBox2-" & i).text = c4.Cells(ligne, 2)    ' TabTarifs(2, i) ' c4.Cells(ligne, 2)
            Me.Frame1("textbox2-" & i).FontSize = 10
            Me.Frame1("TextBox2-" & i).Top = Y
            Me.Frame1("TextBox2-" & i).Left = X + 43
            Me.Frame1("TextBox2-" & i).Width = 150
            Me.Frame1("TextBox2-" & i).Visible = True
            Me.Frame1("Textbox2-" & i).SpecialEffect = 0
            Me.Frame1("Textbox2-" & i).BorderStyle = 1
        End With
        Set R_t3 = Me.Frame1.Controls.Add("Forms.TextBox.1", "TextBox3-" & i, True)
        With R_t3
    '            couleur = c4.Cells(ligne, 3).Interior.Color
            Me.Frame1("textbox3-" & i).BackColor = Couleur
    '            Me.Frame1("textbox3-" & i).ForeColor = c4.Cells(ligne, 3).ForeColor
            Me.Frame1("TextBox3-" & i).text = c4.Cells(ligne, 3)
            Me.Frame1("textbox3-" & i).FontSize = 10
            Me.Frame1("TextBox3-" & i).Top = Y
            Me.Frame1("TextBox3-" & i).Left = X + 193
            Me.Frame1("TextBox3-" & i).Width = 60
            Me.Frame1("TextBox3-" & i).Visible = True
            Me.Frame1("Textbox3-" & i).SpecialEffect = 0
            Me.Frame1("Textbox3-" & i).BorderStyle = 1
        End With
        Set R_t4 = Me.Frame1.Controls.Add("Forms.TextBox.1", "TextBox4-" & i, True)
        With R_t4
            Couleur = c4.Cells(ligne, 4).Interior.color
            Me.Frame1("Textbox4-" & i).BackColor = Couleur
            sstr1 = CStr(Format(c4.Cells(ligne, 4), " 0.00"))
            sstr2 = Replace(sstr1, ".", ",") & " €"
            Me.Frame1("Textbox4-" & i).Value = IIf(sstr1 <> "", sstr2, "")
            Me.Frame1("Textbox4-" & i).FontSize = 10
            Me.Frame1("Textbox4-" & i).Top = Y
            Me.Frame1("Textbox4-" & i).Left = X + 253
            Me.Frame1("Textbox4-" & i).Width = 60
            Me.Frame1("Textbox4-" & i).Visible = True
            Me.Frame1("Textbox4-" & i).SpecialEffect = 0
            Me.Frame1("Textbox4-" & i).BorderStyle = 1
        End With
        Set R_t5 = Me.Frame1.Controls.Add("Forms.Textbox.1", "Textbox5-" & i, True)
        With R_t5
            Couleur = c4.Cells(ligne, 5).Interior.color
            Me.Frame1("Textbox5-" & i).BackColor = Couleur
            Me.Frame1("Textbox5-" & i).FontSize = 10
            Me.Frame1("Textbox5-" & i).text = c4.Cells(ligne, 5)    '  TabTarifs(5, i) ' c4.Cells(ligne, 5)
            Me.Frame1("Textbox5-" & i).Top = Y
            Me.Frame1("Textbox5-" & i).Left = X + 313
            Me.Frame1("Textbox5-" & i).Width = 270
            Me.Frame1("Textbox5-" & i).Visible = True
            Me.Frame1("Textbox5-" & i).SpecialEffect = 0
            Me.Frame1("Textbox5-" & i).BorderStyle = 1
        End With
        X = 3
        Y = Y + 15
    Next i
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            k = k + 1
            ReDim Preserve Cmb(1 To k)
            Set Cmb(k).GroupeCombo = ctrl
        ElseIf TypeName(ctrl) = "TextBox" Then
            j = j + 1
            ReDim Preserve txt(1 To j)
            Set txt(j).GroupeTextb = ctrl
        End If
    Next ctrl    '        Me.ScrollBars.min = 0
    Me.Frame1.ScrollHeight = max * 15
    '       Call SpinButton1_Change
    On Error Resume Next
End Sub
Function Nb00(r As Double, n As Byte) As String
'If Int(inp) Then
'Dim p As Byte: p = InStr(inp, ",")
    Nb00 = Format(r, "0," & String$(2, "0"))
End Function
Sub Sample()
    Dim col As Long
    col = RGB(255, 32, 32)
    '    Debug.Print col    '<~~ Gives 16674815
End Sub
Public Sub Valid_Click()
    changed = True
    '____________________________________ECRITURE DANS FICHIER STOCK______________________________________________
    s3.Cells(ligne, nbCol) = Me("textbox12-" & Trim(Str(i))).text    '  Flacon
    s2.Cells(ligne, nbCol) = Me("textbox4-" & Trim(Str(i))).text     '  Numlot
    s5.Cells(ligne, nbCol) = Me("textbox6-" & Trim(Str(i))).text     '  contenance
    ligne = Tab1(max - 1, 1)
    nbCol = Tab1(max - 1, 2) + 1
    ActiveWorkbook.Save
    Application.WindowState = xlMinimized
    affiche
End Sub
Public Sub tab1_affichage()
    Dim iarr As Integer
    Set c4 = Sheets("TYP_dom")
    c4.Activate
    '    nbrowmax = c4.Range("A65000").End(xlUp).row
    ReDim tabtarifs(10, 1 To 1) As String
    Erase tabtarifs
    max = nbrowmax
    For iarr = 1 To nbrowmax
        ReDim Preserve tabtarifs(10, 1 To iarr + 1)
        tabtarifs(1, iarr) = c4.Cells(iarr, 1).Value
        tabtarifs(2, iarr) = c4.Cells(iarr, 2).Value
        tabtarifs(3, iarr) = c4.Cells(iarr, 3).Value
        tabtarifs(4, iarr) = c4.Cells(iarr, 4).Value
        tabtarifs(5, iarr) = c4.Cells(iarr, 5).Value
        tabtarifs(6, iarr) = c4.Cells(iarr, 6).Value
        tabtarifs(7, iarr) = c4.Cells(iarr, 7).Value
        tabtarifs(8, iarr) = c4.Cells(iarr, 8).Value
        tabtarifs(9, iarr) = c4.Cells(iarr, 9).Value
        tabtarifs(10, iarr) = c4.Cells(iarr, 10).Value
    Next iarr
    max = nbrowmax
End Sub
Private Sub ScrollBar1_Change()
    Call affiche_headers
End Sub
Public Sub affiche_headers()
    X = 3
    Y = 5
    i = 1
    '        ligne = Tab1(i, 1)
    ligne = i
    j = 1    ' Tab1(i, 2)
    Couleur = GetLongFromRGB(255, 255, 255)
    '            Couleur = GetLongFromRGB(192, 224, 192)
    ''    Me.ListBox2.ColumnWidths = "50, 150, 150, 35, 30, 50, 50"
    Set R_t1 = Me.Frame2.Controls.Add("Forms.TextBox.1", "TextBox1-" & i, True)
    With R_t1
        Me.Frame2("Textbox1-" & i).BackColor = Couleur
        Me.Frame2("Textbox1-" & i).text = "CODE"    ' c4.Cells(ligne, 1)
        Me.Frame2("Textbox1-" & i).FontSize = 10
        Me.Frame2("Textbox1-" & i).Top = Y
        Me.Frame2("Textbox1-" & i).Left = X + 3
        Me.Frame2("Textbox1-" & i).Width = 40
        Me.Frame2("Textbox1-" & i).Visible = True
        Me.Frame2("Textbox1-" & i).SpecialEffect = 0
        Me.Frame2("Textbox1-" & i).BorderStyle = 1
        Me.Frame2("Textbox1-" & i).TextAlign = 2
    End With
    Set R_t2 = Me.Frame2.Controls.Add("Forms.TextBox.1", "TextBox2-" & i, True)
    With R_t2
        Me.Frame2("textbox1-" & i).BackColor = Couleur
    ''            ww = Workbooks("Calcul prix et recettes.xlsb").Worksheets("Produits").Range("B" & CStr(ligne)) & " flacon N°: " & Workbooks("Stock.xlsm").Worksheets("Flacon").Cells(ligne, j) & " " & Workbooks("Calcul prix et recettes.xlsb").Worksheets("Produits").Range("F" & CStr(ligne))
        Me.Frame2("TextBox2-" & i).text = "TARIF"    ' c4.Cells(ligne, 2)
        Me.Frame2("textbox2-" & i).FontSize = 10
        Me.Frame2("TextBox2-" & i).Top = Y
        Me.Frame2("TextBox2-" & i).Left = X + 43
        Me.Frame2("TextBox2-" & i).Width = 150
        Me.Frame2("TextBox2-" & i).Visible = True
        Me.Frame2("Textbox2-" & i).SpecialEffect = 0
        Me.Frame2("Textbox2-" & i).BorderStyle = 1
        Me.Frame2("Textbox2-" & i).TextAlign = 2
    End With
    Set R_t3 = Me.Frame2.Controls.Add("Forms.TextBox.1", "TextBox3-" & i, True)
    With R_t3
        Couleur = c4.Cells(ligne, 3).Interior.color
        Me.Frame2("textbox3-" & i).BackColor = Couleur
        Me.Frame2("TextBox3-" & i).text = "LIBELLE"    ' c4.Cells(ligne, 3)
        Me.Frame2("textbox3-" & i).FontSize = 10
        Me.Frame2("TextBox3-" & i).Top = Y
        Me.Frame2("TextBox3-" & i).Left = X + 193
        Me.Frame2("TextBox3-" & i).Width = 60
        Me.Frame2("TextBox3-" & i).Visible = True
        Me.Frame2("Textbox3-" & i).SpecialEffect = 0
        Me.Frame2("Textbox3-" & i).BorderStyle = 1
        Me.Frame2("Textbox3-" & i).TextAlign = 2
    End With
    Set R_t4 = Me.Frame2.Controls.Add("Forms.TextBox.1", "TextBox4-" & i, True)
    With R_t4
        Couleur = c4.Cells(ligne, 4).Interior.color
        Me.Frame2("Textbox4-" & i).BackColor = Couleur
        Me.Frame2("Textbox4-" & i).text = "VALEUR"    ' c4.Cells(ligne, 4)
        Me.Frame2("Textbox4-" & i).FontSize = 10
        Me.Frame2("Textbox4-" & i).Top = Y
        Me.Frame2("Textbox4-" & i).Left = X + 253
        Me.Frame2("Textbox4-" & i).Width = 60
        Me.Frame2("Textbox4-" & i).Visible = True
        Me.Frame2("Textbox4-" & i).SpecialEffect = 0
        Me.Frame2("Textbox4-" & i).BorderStyle = 1
        Me.Frame2("Textbox4-" & i).TextAlign = 2
    End With
    Set R_t5 = Me.Frame2.Controls.Add("Forms.Textbox.1", "Textbox5-" & i, True)
    With R_t5
        Couleur = c4.Cells(ligne, 5).Interior.color
        Me.Frame2("Textbox5-" & i).BackColor = Couleur
        Me.Frame2("Textbox5-" & i).FontSize = 10
        Me.Frame2("Textbox5-" & i).text = "COMMENTAIRE"    ' c4.Cells(ligne, 5)
        Me.Frame2("Textbox5-" & i).Top = Y
        Me.Frame2("Textbox5-" & i).Left = X + 313
        Me.Frame2("Textbox5-" & i).Width = 270
        Me.Frame2("Textbox5-" & i).Visible = True
        Me.Frame2("Textbox5-" & i).SpecialEffect = 0
        Me.Frame2("Textbox5-" & i).BorderStyle = 1
        Me.Frame2("Textbox5-" & i).TextAlign = 2    ' xlCenter
    End With

End Sub
