VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_newjob 
   Caption         =   "USF_newjob"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   OleObjectBlob   =   "USF_newjob.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_newjob"
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
Dim txt2() As New classe_textb_Usf_newjob
Dim cmb2() As New classe_combo_Usf_newjob
Public Couleur As Double
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
End Sub

Private Sub Quitter_Click()
    Unload USF_newjob
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
    Set c6 = Sheets("TYP_trav")
    nbrowmax = c3.Range("B65000").End(xlUp).Row
    Application.EnableEvents = True
    If new_row = True Then
        USF_NewClient.Valid.Enabled = True
        Call affiche_raz
    Else
        USF_NewClient.Valid.Enabled = False
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

    Y = Y + 4
    Y = Y + 25
    X = 3

    Set R_t5 = Me.Controls.Add("Forms.Textbox.1", "Textbox5", True)
    With R_t5
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox5").BackColor = Couleur
        Me("Textbox5").Value = "Typ_TRAVAUX: "
        Me("Textbox5").FontSize = 10
        Me("Textbox5").Top = Y
        Me("Textbox5").Left = X
        Me("Textbox5").Width = 100
        Me("Textbox5").Visible = True
        Me("TextBox5").SpecialEffect = 0
    End With

    Set R_t6 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox6", True)
    With R_t6
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("ComboBox6").BackColor = Couleur
        Me("ComboBox6").Value = ""    ' c3.Cells(ligne, 3).Value
        Me("ComboBox6").FontSize = 10
        Me("ComboBox6").Top = Y
        Me("ComboBox6").Left = X + 105
        Me("ComboBox6").Width = 320
        Me("ComboBox6").Visible = True
        Me("ComboBox6").SpecialEffect = 0
        nbrowmax = c6.Range("A65000").End(xlUp).Row
        Call tri_col_generic(Sheets("TYP_trav"), 1)
        For L = 2 To nbrowmax
            Me("ComboBox6").AddItem c6.Range("A" & Trim(Str(L))) & " - " & c6.Range("F" & Trim(Str(L))) & CStr(" € ")   '  & c5.Range("D" & Trim(Str(L))) & CStr(" €")
        Next
    End With

    Y = Y + 25
    X = 3

    Set R_t5 = Me.Controls.Add("Forms.Textbox.1", "Textbox9", True)
    With R_t5
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox5").BackColor = Couleur
        Me("Textbox9").Value = "ENTREE LIBRE: "
        Me("Textbox9").FontSize = 10
        Me("Textbox9").Top = Y
        Me("Textbox9").Left = X
        Me("Textbox9").Width = 80
        Me("Textbox9").Visible = True
        Me("TextBox9").SpecialEffect = 0
    End With

    Set R_t6 = Me.Controls.Add("Forms.Textbox.1", "TextBox11", True)
    With R_t6
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("ComboBox6").BackColor = Couleur
        Me("TextBox11").Value = ""    ' c3.Cells(ligne, 3).Value
        Me("TextBox11").FontSize = 10
        Me("TextBox11").Top = Y
        Me("TextBox11").Left = X + 85
        Me("TextBox11").Width = 320
        Me("TextBox11").Visible = True
        Me("TextBox11").SpecialEffect = 0
        nbrowmax = c6.Range("A65000").End(xlUp).Row
    End With


    Set R_t6 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox10", True)
    With R_t6
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("ComboBox6").BackColor = Couleur
        Me("ComboBox10").Value = ""    ' c3.Cells(ligne, 3).Value
        Me("ComboBox10").FontSize = 10
        Me("ComboBox10").Top = Y
        Me("ComboBox10").Left = X + 410
        Me("ComboBox10").Width = 180
        Me("ComboBox10").Visible = True
        Me("ComboBox10").SpecialEffect = 0
        nbrowmax = c6.Range("A65000").End(xlUp).Row
        Call tri_col_generic(Sheets("TYP_trav"), 1)
        For L = 2 To nbrowmax
            Me("ComboBox10").AddItem c6.Range("A" & Trim(Str(L))) & " - " & c6.Range("F" & Trim(Str(L))) & CStr(" € ")   '  & c5.Range("D" & Trim(Str(L))) & CStr(" €")
        Next
    End With

    Y = Y + 25
    X = 3

    Set R_t7 = Me.Controls.Add("Forms.Textbox.1", "Textbox7", True)
    With R_t7
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox7").BackColor = Couleur
        Me("Textbox7").Value = "Nombre de travaux: "
        Me("Textbox7").FontSize = 10
        Me("Textbox7").Top = Y
        Me("Textbox7").Left = X
        Me("Textbox7").Width = 100
        Me("Textbox7").Visible = True
        Me("TextBox7").SpecialEffect = 0
    End With

    Set R_t8 = Me.Controls.Add("Forms.Textbox.1", "Textbox8", True)
    With R_t8
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox8").BackColor = Couleur
        Me("Textbox8").Value = ""    ' c3.Cells(ligne, 4).Value
        Me("Textbox8").FontSize = 10
        Me("Textbox8").Top = Y
        Me("Textbox8").Left = X + 103
        Me("Textbox8").Width = 20
        Me("Textbox8").Visible = True
        Me("TextBox8").SpecialEffect = 0
    End With
    Set R_t9 = Me.Controls.Add("Forms.Textbox.1", "TextBox12", True)
    With R_t9
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("ComboBox6").BackColor = Couleur
        Me("TextBox12").Value = "MOIS DE FACTURATION (par defaut):"    ' c3.Cells(ligne, 3).Value
        Me("TextBox12").FontSize = 10
        Me("TextBox12").Top = Y
        Me("TextBox12").Left = X + 125
        Me("TextBox12").Width = 185
        Me("TextBox12").Visible = True
        Me("TextBox12").SpecialEffect = 0
        nbrowmax = c6.Range("A65000").End(xlUp).Row
    End With


    Set R_t10 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox13", True)
    With R_t10
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("ComboBox6").BackColor = Couleur
        Me("ComboBox13").Value = ""    ' c3.Cells(ligne, 3).Value
        Me("ComboBox13").FontSize = 10
        Me("ComboBox13").Top = Y
        Me("ComboBox13").Left = X + 315
        Me("ComboBox13").Width = 120
        Me("ComboBox13").Visible = True
        Me("ComboBox13").SpecialEffect = 0
        Set c6 = Sheets("TYP_trav")
        nbrowmax = c6.Range("I65000").End(xlUp).Row
        Call tri_col_generic(Sheets("TYP_trav"), 1)
        For L = 2 To nbrowmax
            Me("ComboBox13").AddItem c6.Range("I" & Trim(Str(L)))
        Next
    End With
    k = Month(Date)
'    Me("ComboBox13").ListIndex = Fact_clients.comp_mois_rev(k, Month(Date)) + 2
    Me("ComboBox13").Value = Fact_clients.comp_mois_rev(k)
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            i = i + 1
            ReDim Preserve cmb2(1 To i)
            Set cmb2(i).GroupeCombo2 = ctrl
        ElseIf TypeName(ctrl) = "TextBox" Then
            L = L + 1
            ReDim Preserve txt2(1 To L)
            Set txt2(L).GroupeTextb2 = ctrl
        End If
    Next ctrl
    On Error Resume Next
    '    affiche1
    AppliqueArrondi
    ''        AppliqueTransp
End Sub
Public Sub affich_raz()
    Me("Textbox1").Value = ""    ' c3.Cells(ligne, 14)
    Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
    Me("TextBox3").Value = ""      ' c3.Cells(ligne, 2).Value
    Me("ComboBox6").Value = ""    ' c3.Cells(ligne, 3).Value
    Me("Textbox8").Value = ""    'c3.Cells(ligne, 4).Value
    Me("Textbox10").Value = ""    'c3.Cells(ligne, 6).Value
    Me("Textbox12").Value = ""    'c3.Cells(ligne, 6).Value
    Me("Textbox14").Value = ""    'c3.Cells(ligne, 22).Value
    Me("Textbox16").Value = ""    ' c3.Cells(ligne, 8).Value
    Me("Textbox20").Value = ""    'c3.Cells(ligne, 9).Value
    Me("Textbox22").Value = ""    ' c3.Cells(ligne, 11).Value
    Me("Textbox24").Value = ""    'c3.Cells(ligne, 12).Value
    Me("Textbox26").Value = ""    'c3.Cells(ligne, 13).Value
    Me("Textbox28").Value = ""    'c3.Cells(ligne, 14).Value
End Sub
Public Sub affiche1()
    ligne = pos2
    Me("Textbox1").Value = c3.Cells(ligne, 6).Value       ' Societe
    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = c3.Cells(ligne, 14).Value       ' Repet Societe
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
    If Me("TextBox8").Value <> "" Then
        Set c2 = Sheets("Travaux")
        nbrowmax = c2.Range("A65000").End(xlUp).Row
        ligne = nbrowmax + 1
        If selct = True Then
            c2.Cells(ligne, 1).Value = c3.Cells(pos, 7).Value                       ' Num client
            c2.Cells(ligne, 2).Value = c3.Cells(pos, 14).Value                      ' Societe
            c2.Cells(ligne, 3).Value = CStr(Sheets("Typ_trav").Cells(Lig, 1))      ' typjob
            c2.Cells(ligne, 4).Value = Me("Textbox8").Value                         ' Nb travaux
            c2.Cells(ligne, 5).Value = CStr(Sheets("Typ_trav").Cells(Lig, 3))      ' PUht trav
            c2.Cells(ligne, 6).Value = CStr(Sheets("Typ_trav").Cells(Lig, 2))       ' Ville
            c2.Cells(ligne, 7) = UCase(Format(Me("ComboBox13").Value, "mmmm"))
''            c2.Cells(ligne, 8) = Format(Date, "[$-409]dd.mm.yyyy") '  ' UCase(Format(Date, "DD/MM/YYYY"))
'            c2.Cells(ligne, 8) = Format(Date, "[$-040C]dd/mm/yyyy") '  ' UCase(Format(Date, "DD/MM/YYYY"))
             c2.Cells(ligne, 8) = ttk_Date(Date)
        Else
            c2.Cells(ligne, 1).Value = c3.Cells(pos, 7).Value                       ' Num client
            c2.Cells(ligne, 2).Value = c3.Cells(pos, 14).Value                      ' Societe
            c2.Cells(ligne, 3).Value = Me("Textbox11").Value                         ' typjob
            c2.Cells(ligne, 4).Value = Me("Textbox8").Value                         ' Nb travaux
            c2.Cells(ligne, 5).Value = CStr(Sheets("Typ_trav").Cells(Lig, 3))      ' PUht trav
            c2.Cells(ligne, 6).Value = CStr(Sheets("Typ_trav").Cells(Lig, 2))       ' Ville
            c2.Cells(ligne, 7) = UCase(Format(Me("ComboBox13").Value, "mmmm"))
'            c2.Cells(ligne, 8) = Format(Date, "[$-409]dd.mm.yyyy") '  '  UCase(Format(Date, "DD/MM/YYYY"))
            c2.Cells(ligne, 8) = Format(Date, "[$-040C]dd/mm/yyyy") '  ' UCase(Format(Date, "DD/MM/YYYY"))
            c2.Cells(ligne, 8) = ttk_Date(Date)
         End If
        affiche1
        Sheets("CLIENTS").Activate
        USF_newjob.Repaint
        USF_newjob.Valid.Enabled = True
    Else
        USF_newjob("Textbox8").SetFocus
        MsgBox ("Saisissez un nombre de travaux.")
    End If
End Sub

Private Sub UserForm2_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Cancel = 0 Then
        Unload Me
        End
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





