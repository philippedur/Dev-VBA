VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Eff_client 
   Caption         =   "USF_newjob"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   OleObjectBlob   =   "USF_Eff_client.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Eff_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    ' USERFORM UsF_newjob !!!!!!!!!!!!!!!!!!!
Private Declare PtrSafe Function FindWindowA Lib "User32" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Ht, Ncol As Integer
Dim chaine As String
Dim TB()
'Dim txt2() As New classe_textb_Usf_newjob
Dim cmb2() As New classe_combo_Usf_Eff_client
Public Couleur As Double
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
End Sub
Private Sub Quitter_Click()
    Unload USF_Eff_client
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
Private Sub USF_client_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set usf = Nothing
End Sub
Public Sub Userform_initialize()    '///// PAS DE CHECK DE CELLULES VIDES
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Set c6 = Sheets("EBP-Xtract-expert")
    Set c7 = Sheets("Buff2")
    Set c8 = Sheets("Gestion")
    Set c9 = Sheets("Clients resilies")
    nbrowmax = c3.Range("B65000").End(xlUp).Row
    Application.EnableEvents = True
    Eff_client_valid = False
    If new_row = True Then
        USF_Eff_client.Valid.Enabled = True
        Call affiche_raz
    Else
        USF_Eff_client.Valid.Enabled = True
    End If
    '    nbCol = c1.Cells(ligne, Columns.count).End(xlToLeft).Column + 1
    X = 3
    Y = 35
    i = 1
    Application.EnableEvents = True
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    ' Couleur = GetLongFromRGB(0, 64, 128)
    '    If trig = False Then
    '        ligne = 1
    '    Else: trig = pos
    '    End If

    Set R_C2 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox2", True)
    With R_C2
    '        ' Couleur = C1.Cells(ligne, i).Interior.Color
    '        Me("ComboBox2").BackColor = Couleur
        Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
        Me("ComboBox2").FontSize = 10
        Me("ComboBox2").Top = Y
        Me("ComboBox2").Left = X
        Me("ComboBox2").Width = 200
        Me("ComboBox2").Visible = True
    '        Me("ComboBox2").Clear
        Me("ComboBox2").SpecialEffect = 0
        Call tri_col_generic(Sheets("CLIENTS"), 14)
        For L = 2 To c3.Range("N" & Rows.Count).End(xlUp).Row
            Me("ComboBox2").AddItem c3.Range("N" & Trim(Str(L)))
        Next
    End With
    Set R_C1 = Me.Controls.Add("Forms.Textbox.1", "Textbox1", True)
    With R_C1
    ' Couleur = c1.Cells(ligne, i).Interior.Color
    '        Me("Textbox1").ForeColor = &H80000011
        Me("Textbox1").Value = ""
        Me("Textbox1").FontSize = 10
        Me("Textbox1").Top = Y
        Me("Textbox1").Left = X + 200
        Me("Textbox1").Width = 125
        Me("Textbox1").Visible = True
        Me("Textbox1").SpecialEffect = 0
    '        Call tri_col_generic(Sheets("CLIENTS"), 6)
    '        For L = 2 To c3.Range("F" & Rows.Count).End(xlUp).Row
    '            Me("Textbox1").AddItem c3.Range("F" & Trim(Str(L)))
    '        Next
    End With

    Set R_t3 = Me.Controls.Add("Forms.Textbox.1", "Textbox3", True)
    With R_t3
    '        ' Couleur = c1.Cells(ligne, i).ForeColor
    '        Me("Textbox3").ForeColor = Couleur
        Me("Textbox3").Value = ""    ' c3.Cells(ligne, 14)
        Me("Textbox3").FontSize = 15
        Me("TextBox3").Font.Bold = True
        Me("Textbox3").Top = Y    ' - 4
        Me("Textbox3").Left = X + 325
        Me("Textbox3").Width = 350
        Me("Textbox3").Visible = True
        Me("TextBox3").SpecialEffect = 0
    End With

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            i = i + 1
            ReDim Preserve cmb2(1 To i)
            Set cmb2(i).GroupeCombo2 = ctrl
        End If
    Next ctrl

    On Error Resume Next
    affiche1
    AppliqueArrondi
    ''        AppliqueTransp
End Sub
Public Sub affich_raz()
    Me("Textbox1").Value = ""    ' c3.Cells(ligne, 14)
    Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
    Me("TextBox3").Value = ""      ' c3.Cells(ligne, 2).Value
End Sub
Public Sub affiche1()
    ligne = pos
    '    Me("Textbox1").Value = c3.Cells(ligne, 6).Value       ' Societe
    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = Me("ComboBox2").Value       ' Repet Societe
    USF_Eff_client.Valid.Enabled = True
    '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
    '    Me("Textbox8").Value = c3.Cells(ligne, 20).Value        ' Ville
End Sub
Public Sub affiche2()
    ligne = pos
    Me("Textbox1").Value = c3.Cells(ligne, 6).Value       ' Societe
    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = c3.Cells(ligne, 14).Value       ' Repet Societe
    '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
    '    Me("Textbox8").Value = c3.Cells(ligne, 20).Value        ' Ville
End Sub
Public Sub Valid_Click()
    Call tri_col_generic(Sheets("CLIENTS"), 14)
    Dim Msg, style, Title, Help, Ctxt, response, mystring
    ligne = pos
    nbrowmax2 = c9.Range("A65000").End(xlUp).Row + 1
    affiche1
    Set c2 = Worksheets("Travaux")
    Worksheets("Travaux").Activate
    
    If FindAll_OneRec(cle_rech, Sheets("Travaux"), "B2:B" & nbrowmax, Armatches()) Then
        ligne = pos
        USF_Eff_client.Valid.Enabled = True
        Msg = "Il y a des travaux en cours de facturation.." & vbCrLf & "Etes vous sûr de vouloir supprimer ce client ?"    ' Define message.
        style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Effacement d'un client"
        Help = ""
        Ctxt = 1000
        response = MsgBox(Msg, style, Title, Help, Ctxt)
        If response = vbYes Then
            mystring = "Yes"
            '  COPY RECORD TO  Format(Date, "DD/MM/YYYY")
            nbrowmax = c3.Range("F65000").End(xlUp).Row
            If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            ligne = pos
            source_range = Range(c3.Cells(ligne, 1), c3.Cells(ligne, 30)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Worksheets("CLIENTS").Range(source_range).Copy
            Set c9 = Sheets("Clients resilies")
            nbrowmax2 = c9.Range("N65000").End(xlUp).Row + 1
            dest_range = Range(c9.Cells(nbrowmax2, 1), c9.Cells(nbrowmax2, 30)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            With dest_range
        Worksheets("Clients resilies").Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                                                                   xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
            nom_entreprise = Trim(c3.Range("N" & pos))
            siren = Left(c3.Range("I" & pos), 9)
            nom_entreprise = Trim(c3.Range("N" & pos))
            Call Infos_Jur(Sheets("Clients resilies"), nom_entreprise, siren, nbrowmax2)
'            c9.Hyperlinks.Add Anchor:=c9.Cells(nbrowmax2, 25), Address:="https://www.pappers.fr/nom_entreprise/" & siren, TextToDisplay:="Pappers-" & nom_entreprise               ' Creation
            End With

         End If
            c3.Rows(pos).EntireRow.Delete
        Else
            mystring = "No"
            Exit Sub
        End If
    End If
        USF_Eff_client.Valid.Enabled = True
        Eff_client_valid = False
        Msg = "Etes vous sûr de vouloir supprimer ce client ?"    ' Define message.
        style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Effacement d'un client"
        Help = ""
        Ctxt = 1000
        response = MsgBox(Msg, style, Title, Help, Ctxt)
        If response = vbYes Then
            mystring = "Yes"
            '  COPY RECORD TO  Format(Date, "DD/MM/YYYY")
            If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
            ligne = pos
            source_range = Range(c3.Cells(ligne, 1), c3.Cells(ligne, 30)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Worksheets("CLIENTS").Range(source_range).Copy
            Set c9 = Sheets("Clients resilies")
            nbrowmax2 = c9.Range("N65000").End(xlUp).Row + 1
            dest_range = Range(c9.Cells(nbrowmax2, 1), c9.Cells(nbrowmax2, 30)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            With dest_range
        Worksheets("Clients resilies").Range(dest_range).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                                                                   xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
            c3.Activate
            nom_entreprise = Trim(c3.Range("N" & pos))
            siren = Left(c3.Range("I" & pos), 9)
            nom_entreprise = Trim(c3.Range("N" & pos))
            If siren = "" Then
            Else
            c3.Range("Y" & pos) = ""
            Call Infos_Jur(Sheets("Clients resilies"), nom_entreprise, siren, nbrowmax)
'            c9.Hyperlinks.Add Anchor:=c9.Cells(nbrowmax2, 25), Address:="https://www.pappers.fr/nom_entreprise/" & siren, TextToDisplay:="Pappers-" & nom_entreprise              ' Creation
            End If
            End With
            c3.Rows(pos).EntireRow.Delete
        Else
            mystring = "No"
            Exit Sub
        End If
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
'    RoundCorners Me, Me.Width, Me.Height, 30, 30
End Sub



