VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Client_edit 
   Caption         =   "USF_Client_edit"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   OleObjectBlob   =   "USF_Client_edit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Client_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    ' USERFORM USF_Client_edit !!!!!!!!!!!!!!!!!!!
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
Dim txt() As New classe_textb_Usf_client_edit
Dim Cmb() As New classe_combo_Usf_client_edit
Public Couleur As Double
Public CollectTx As Collection
Private Tx As MSForms.TextBox
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
End Sub
Private Sub Quitter_Click()
    Unload USF_Client_edit
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
Private Sub USF_Client_edit_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set usf = Nothing
End Sub
Public Sub Userform_initialize()    '///// PAS DE CHECK DE CELLULES VIDES
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    USF_Client_edit.BackColor = GetLongFromRGB(64, 192, 224)
    USF_Client_edit.Frame1.BackColor = GetLongFromRGB(64, 192, 224)
    nbrowmax = c3.Range("B65000").End(xlUp).Row
    Application.EnableEvents = True
    If new_row = True Then
        USF_Client_edit.Valid.Enabled = True
        Call affiche_raz
    Else
        USF_Client_edit.Valid.Enabled = True
    End If
    '    nbCol = c1.Cells(ligne, Columns.count).End(xlToLeft).Column + 1
    X = 3
    Y = 50
    i = 1
    Application.EnableEvents = True
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    ' Couleur = GetLongFromRGB(0, 64, 128)
    If trig = False Then
        ligne = 1
    Else: trig = pos
    End If

    Set R_C1 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox1", True)
    With R_C1
    ' Couleur = c1.Cells(ligne, i).Interior.Color
    '        Me("ComboBox1").ForeColor = &H80000011
        Me("ComboBox1").Value = ""
        Me("ComboBox1").FontSize = 10
        Me("ComboBox1").Top = Y
        Me("ComboBox1").Left = X
        Me("ComboBox1").Width = 120
        Me("ComboBox1").Visible = True
        Me("ComboBox1").SpecialEffect = 0
        Call tri_col_generic(Sheets("CLIENTS"), 6)
        For L = 2 To c3.Range("F" & Rows.Count).End(xlUp).Row
            Me("ComboBox1").AddItem c3.Range("F" & Trim(Str(L)))
        Next
    End With

    Set R_C2 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox2", True)
    With R_C2
    '        ' Couleur = C1.Cells(ligne, i).Interior.Color
    '        Me("ComboBox2").BackColor = Couleur
        Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
        Me("ComboBox2").FontSize = 10
        Me("ComboBox2").Top = Y
        Me("ComboBox2").Left = X + 125
        Me("ComboBox2").Width = 200
        Me("ComboBox2").Visible = True
    '        Me("ComboBox2").Clear
        Me("ComboBox2").SpecialEffect = 0
        Call tri_col_generic(Sheets("CLIENTS"), 14)
        For L = 2 To c3.Range("N" & Rows.Count).End(xlUp).Row
            Me("ComboBox2").AddItem c3.Range("N" & Trim(Str(L)))
        Next
    End With

    Set R_t3 = Me.Controls.Add("Forms.Textbox.1", "Textbox3", True)
    With R_t3
    '        ' Couleur = c1.Cells(ligne, i).ForeColor
    '        Me("Textbox3").ForeColor = Couleur
        Me("Textbox3").Value = ""    ' c3.Cells(ligne, 14)
        Me("Textbox3").FontSize = 15
        Me("TextBox3").Font.Bold = True
        Me("Textbox3").Top = Y - 3.6
        Me("Textbox3").Left = X + 330
        Me("Textbox3").Height = 25.2
        Me("Textbox3").Width = 350
        Me("Textbox3").Visible = True
        Me("TextBox3").SpecialEffect = 0
    End With

    '  Set R_t4 = Me.Controls.Add("Forms.TextBox.1", "TextBox4", True)
    '    With R_t4
    '        ' Couleur = c3.Cells(ligne, i).Interior.Color
    '    '        Me("TextBox4").BackColor = Couleur
    '        Me("TextBox4").Value = c3.Cells(ligne, 14)
    '        Me("TextBox4").FontSize = 15
    '        Me("TextBox4").Font.Bold = True
    '        Me("TextBox4").Top = Y - 4
    '        Me("TextBox4").Left = X + 260
    '        Me("TextBox4").Width = 400
    '        Me("TextBox4").Visible = True
    '        Me("TextBox4").SpecialEffect = 0
    '    End With

    Y = Y + 4
    Y = Y + 25
    X = 3

    Set R_t5 = Me.Controls.Add("Forms.Textbox.1", "Textbox5", True)
    With R_t5
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox5").BackColor = Couleur
        Me("Textbox5").Value = "Adresse Facturation: "
        Me("Textbox5").FontSize = 10
        Me("Textbox5").Top = Y
        Me("Textbox5").Left = X
        Me("Textbox5").Width = 100
        Me("Textbox5").Visible = True
        Me("TextBox5").SpecialEffect = 0
    End With

    Set R_t6 = Me.Controls.Add("Forms.TextBox.1", "TextBox6", True)
    With R_t6
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("TextBox6").BackColor = Couleur
        Me("TextBox6").Value = ""    ' c3.Cells(ligne, 3).Value
        Me("TextBox6").FontSize = 10
        Me("TextBox6").Top = Y
        Me("TextBox6").Left = X + 100
        Me("TextBox6").Width = 320
        Me("TextBox6").Visible = True
        Me("TextBox6").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t7 = Me.Controls.Add("Forms.Textbox.1", "Textbox7", True)
    With R_t7
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox7").BackColor = Couleur
        Me("Textbox7").Value = "Ville: "
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
        Me("Textbox8").Left = X + 100
        Me("Textbox8").Width = 320
        Me("Textbox8").Visible = True
        Me("TextBox8").SpecialEffect = 0
    End With
    Set R_t27 = Me.Controls.Add("Forms.Textbox.1", "Textbox29", True)
    With R_t27
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox29").BackColor = Couleur
        Me("Textbox29").Value = "CP: "
        Me("Textbox29").FontSize = 10
        Me("Textbox29").Top = Y
        Me("Textbox29").Left = X + 425
        Me("Textbox29").Width = 20
        Me("Textbox29").Visible = True
        Me("Textbox29").SpecialEffect = 0
    End With

    Set R_t28 = Me.Controls.Add("Forms.Textbox.1", "Textbox30", True)
    With R_t28
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox30").BackColor = Couleur
        Me("Textbox30").Value = ""    ' c3.Cells(ligne, 24).Value
        Me("Textbox30").FontSize = 10
        Me("Textbox30").Top = Y
        Me("Textbox30").Left = X + 450
        Me("Textbox30").Width = 50
        Me("Textbox30").Visible = True
        Me("Textbox30").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t9 = Me.Controls.Add("Forms.Textbox.1", "Textbox9", True)
    With R_t9
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox9").BackColor = Couleur
        Me("Textbox9").Value = "Gérant : "
        Me("Textbox9").FontSize = 10
        Me("Textbox9").Top = Y
        Me("Textbox9").Left = X
        Me("Textbox9").Width = 70
        Me("Textbox9").Visible = True
        Me("TextBox9").SpecialEffect = 0
    End With

    Set R_t10 = Me.Controls.Add("Forms.Textbox.1", "Textbox10", True)
    With R_t10
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox10").BackColor = Couleur
        Me("Textbox10").Value = ""    ' c3.Cells(ligne, 6).Value
        Me("Textbox10").FontSize = 10
        Me("Textbox10").Top = Y
        Me("Textbox10").Left = X + 70
        Me("Textbox10").Width = 100
        Me("Textbox10").Visible = True
        Me("TextBox10").SpecialEffect = 0
    End With


    Set R_t11 = Me.Controls.Add("Forms.Textbox.1", "Textbox11", True)
    With R_t11
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox11").BackColor = Couleur
        Me("Textbox11").Value = "Mail: "
        Me("Textbox11").FontSize = 10
        Me("Textbox11").Top = Y
        Me("Textbox11").Left = X + 170
        Me("Textbox11").Width = 55
        Me("Textbox11").Visible = True
        Me("TextBox11").SpecialEffect = 0
    End With

    Set R_t12 = Me.Controls.Add("Forms.Textbox.1", "Textbox12", True)
    With R_t12
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox12").BackColor = Couleur
        Me("Textbox12").Value = ""    ' c3.Cells(ligne, 22).Value
        Me("Textbox12").FontSize = 10
        Me("Textbox12").Top = Y
        Me("Textbox12").Left = X + 235
        Me("Textbox12").Width = 150
        Me("Textbox12").Visible = True
        Me("TextBox12").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t13 = Me.Controls.Add("Forms.Textbox.1", "Textbox13", True)
    With R_t13
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox13").BackColor = Couleur
        Me("Textbox13").Value = "Tel 1 : "
        Me("Textbox13").FontSize = 10
        Me("Textbox13").Top = Y
        Me("Textbox13").Left = X
        Me("Textbox13").Width = 70
        Me("Textbox13").Visible = True
        Me("TextBox13").SpecialEffect = 0
    End With

    Set R_t14 = Me.Controls.Add("Forms.Textbox.1", "Textbox14", True)
    With R_t14
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox14").BackColor = Couleur
        Me("Textbox14").Value = ""    ' c3.Cells(ligne, 22).Value
        Me("Textbox14").FontSize = 10
        Me("Textbox14").Top = Y
        Me("Textbox14").Left = X + 70
        Me("Textbox14").Width = 100
        Me("Textbox14").Visible = True
        Me("TextBox14").SpecialEffect = 0
    End With


    Set R_t15 = Me.Controls.Add("Forms.Textbox.1", "Textbox15", True)
    With R_t15
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox15").BackColor = Couleur
        Me("Textbox15").Value = "Tel 2 : "
        Me("Textbox15").FontSize = 10
        Me("Textbox15").Top = Y
        Me("Textbox15").Left = X + 180
        Me("Textbox15").Width = 70
        Me("Textbox15").Visible = True
        Me("TextBox15").SpecialEffect = 0
    End With

    Set R_t16 = Me.Controls.Add("Forms.Textbox.1", "Textbox16", True)
    With R_C16
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox16").BackColor = Couleur
        Me("Textbox16").Value = ""    ' c3.Cells(ligne, 23).Value
        Me("Textbox16").FontSize = 10
        Me("Textbox16").Top = Y
        Me("Textbox16").Left = X + 260
        Me("Textbox16").Width = 100
        Me("Textbox16").Visible = True
        Me("TextBox16").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t19 = Me.Controls.Add("Forms.Textbox.1", "Textbox17", True)
    With R_t19
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox17").BackColor = Couleur
        Me("Textbox17").Value = "TARIF : "
        Me("Textbox17").FontSize = 10
        Me("Textbox17").Top = Y
        Me("Textbox17").Left = X
        Me("Textbox17").Width = 65
        Me("Textbox17").Visible = True
        Me("Textbox17").SpecialEffect = 0
    End With

    Set R_t20 = Me.Controls.Add("Forms.textBox.1", "Textbox18", True)
    With R_t20
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox18").BackColor = Couleur
        Me("Textbox18").Value = ""    ' c3.Cells(ligne, 18).Value
        Me("Textbox18").FontSize = 10
        Me("Textbox18").Top = Y
        Me("Textbox18").Left = X + 70
        Me("Textbox18").Width = 70
        Me("Textbox18").Visible = True
        Me("Textbox18").SpecialEffect = 0
    End With



    Set R_t19 = Me.Controls.Add("Forms.Textbox.1", "Textbox19", True)
    With R_t19
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox19").BackColor = Couleur
        Me("Textbox19").Value = "Typ_dom : "
        Me("Textbox19").FontSize = 10
        Me("Textbox19").Top = Y
        Me("Textbox19").Left = X + 148
        Me("Textbox19").Width = 65
        Me("Textbox19").Visible = True
        Me("TextBox19").SpecialEffect = 0
    End With

    Set R_t20 = Me.Controls.Add("Forms.textBox.1", "Textbox20", True)
    With R_t20
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox20").BackColor = Couleur
        Me("Textbox20").Value = ""    ' c3.Cells(ligne, 18).Value
        Me("Textbox20").FontSize = 10
        Me("Textbox20").Top = Y
        Me("Textbox20").Left = X + 218
        Me("Textbox20").Width = 70
        Me("Textbox20").Visible = True
        Me("TextBox20").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t21 = Me.Controls.Add("Forms.Textbox.1", "Textbox21", True)
    With R_t21
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox21").BackColor = Couleur
        Me("Textbox21").Value = "SIREN : "
        Me("Textbox21").FontSize = 10
        Me("Textbox21").Top = Y
        Me("Textbox21").Left = X
        Me("Textbox21").Width = 55
        Me("Textbox21").Visible = True
        Me("TextBox21").SpecialEffect = 0
    End With

    Set R_t22 = Me.Controls.Add("Forms.Textbox.1", "Textbox22", True)
    With R_t22
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox22").BackColor = Couleur
        Me("Textbox22").Value = ""    ' c3.Cells(ligne, 9).Value
        Me("Textbox22").FontSize = 10
        Me("Textbox22").Top = Y
        Me("Textbox22").Left = X + 60
        Me("Textbox22").Width = 150
        Me("Textbox22").Visible = True
        Me("TextBox22").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t23 = Me.Controls.Add("Forms.Textbox.1", "Textbox23", True)
    With R_t23
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox23").BackColor = Couleur
        Me("Textbox23").Value = "TVA-DOM : "
        Me("Textbox23").FontSize = 10
        Me("Textbox23").Top = Y
        Me("Textbox23").Left = X
        Me("Textbox23").Width = 55
        Me("Textbox23").Visible = True
        Me("TextBox23").SpecialEffect = 0
    End With

    Set R_t24 = Me.Controls.Add("Forms.Textbox.1", "Textbox24", True)
    With R_t24
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox24").BackColor = Couleur
        Me("Textbox24").Value = ""    ' c3.Cells(ligne, 19).Value
        Me("Textbox24").FontSize = 10
        Me("Textbox24").Top = Y
        Me("Textbox24").Left = X + 60
        Me("Textbox24").Width = 150
        Me("Textbox24").Visible = True
        Me("TextBox24").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t25 = Me.Controls.Add("Forms.Textbox.1", "Textbox25", True)
    With R_t25
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox25").BackColor = Couleur
        Me("Textbox25").Value = "IBAN : "
        Me("Textbox25").FontSize = 10
        Me("Textbox25").Top = Y
        Me("Textbox25").Left = X
        Me("Textbox25").Width = 45
        Me("Textbox25").Visible = True
        Me("TextBox25").SpecialEffect = 0
    End With

    Set R_t26 = Me.Controls.Add("Forms.Textbox.1", "Textbox26", True)
    With R_t26
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox26").BackColor = Couleur
        Me("Textbox26").Value = c3.Cells(ligne, 13).Value
        Me("Textbox26").FontSize = 10
        Me("Textbox26").Top = Y
        Me("Textbox26").Left = X + 50
        Me("Textbox26").Width = 200
        Me("Textbox26").Visible = True
        Me("TextBox26").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3

    Set R_t27 = Me.Controls.Add("Forms.Textbox.1", "Textbox27", True)
    With R_t27
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox27").BackColor = Couleur
        Me("Textbox27").Value = "Périodicité : "
        Me("Textbox27").FontSize = 10
        Me("Textbox27").Top = Y
        Me("Textbox27").Left = X
        Me("Textbox27").Width = 55
        Me("Textbox27").Visible = True
        Me("TextBox27").SpecialEffect = 0
    End With


    Set R_C4 = Me.Controls.Add("Forms.Combobox.1", "Combobox3", True)
    With R_C4
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox28").BackColor = Couleur
        Me("Combobox3").Value = ""    ' c3.Cells(ligne, 24).Value
        Me("Combobox3").FontSize = 10
        Me("Combobox3").Top = Y
        Me("Combobox3").Left = X + 60
        Me("Combobox3").Width = 40
        Me("Combobox3").Visible = True
        Me("Combobox3").SpecialEffect = 0
        Me("Combobox3").AddItem "1"
        Me("Combobox3").AddItem "2"
        Me("Combobox3").AddItem "3"
        Me("Combobox3").AddItem "4"
        Me("Combobox3").AddItem "6"
        Me("Combobox3").AddItem "12"
    End With
    

    Set R_t28 = Me.Controls.Add("Forms.Textbox.1", "Textbox28", True)
    With R_t28
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Textbox28").BackColor = Couleur
        Me("Textbox28").Value = "Sexe"
        Me("Textbox28").FontSize = 10
        Me("Textbox28").Top = Y
        Me("Textbox28").Left = X + 105
        Me("Textbox28").Width = 50
        Me("Textbox28").Visible = True
        End With
        
        
        Set R_C21 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox21", True)
    With R_C21
    ' Couleur = c3.Cells(ligne, i).Interior.Color
    '        Me("Combobox20").BackColor = Couleur
        Me("Combobox21").Value = "H/F"    ' c3.Cells(ligne, 18).Value
        Me("Combobox21").FontSize = 10
        Me("Combobox21").Top = Y
        Me("Combobox21").Left = X + 160
        Me("Combobox21").Width = 60
        Me("Combobox21").Visible = True
        Me("Combobox21").SpecialEffect = 0
        Me("ComboBox21").AddItem "Mr"
        Me("ComboBox21").AddItem "Mme"
        Me("ComboBox21").AddItem "Melle"
    End With
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            i = i + 1
            ReDim Preserve Cmb(1 To i)
            Set Cmb(i).GroupeCombo = ctrl
        ElseIf TypeName(ctrl) = "TextBox" Then
            L = L + 1
            ReDim Preserve txt(1 To L)
            Set txt(L).GroupeTextb2 = ctrl
        End If
    Next ctrl
    trig = True
    On Error Resume Next
    AppliqueArrondi
    ''        AppliqueTransp
End Sub
Public Sub affich_raz()
    Me("ComboBox1").Value = ""    ' c3.Cells(ligne, 14)
    Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
    Me("TextBox3").Value = ""      ' c3.Cells(ligne, 2).Value
    Me("TextBox6").Value = ""    ' c3.Cells(ligne, 3).Value
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
Public Sub affiche()
    ligne = pos
    Me("ComboBox1").Value = c3.Cells(ligne, 6).Value        ' Gerant
    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = c3.Cells(ligne, 14).Value        ' Repet Societe
    Me("TextBox6").Value = c3.Cells(ligne, 1).Value         ' Adresse Facturation
    Me("TextBox30").Value = c3.Cells(ligne, 2).Value         ' CP
    Me("Textbox8").Value = c3.Cells(ligne, 3).Value         ' Ville
    Me("Textbox10").Value = c3.Cells(ligne, 6).Value        ' Gérant
    Me("Textbox12").Value = c3.Cells(ligne, 21).Value       ' Mail
    Me("Textbox14").Value = c3.Cells(ligne, 22).Value       ' Tel1
    Me("Textbox16").Value = c3.Cells(ligne, 23).Value       ' Tel2
    Me("Textbox18").Value = c3.Cells(ligne, 18).Value       ' TARIF
    Me("Textbox20").Value = c3.Cells(ligne, 19).Value       ' Typ_DOM
    Me("Textbox22").Value = c3.Cells(ligne, 9).Value        ' SIREN/SIRET
    Me("Textbox24").Value = c3.Cells(ligne, 12).Value
    Me("Textbox26").Value = c3.Cells(ligne, 13).Value
    Me("Combobox3").Value = c3.Cells(ligne, 24).Value       ' périodicité
    Me("Combobox21").Value = c3.Cells(ligne, 8).Value       ' périodicité
End Sub
Public Sub Valid_Click()
    Call tri_col_generic(Sheets("CLIENTS"), 14)
    Set c3 = Sheets("CLIENTS")
    ligne = pos
    c3.Cells(ligne, 6).Value = Me("ComboBox1").Value       ' Gerant
    c3.Cells(ligne, 14).Value = Me("ComboBox2").Value      ' Societe
'    c3.Cells(ligne, 4).Value = Format(Date, "DD/MM/YYYY")  ' Date création Societe
    c3.Cells(ligne, 4).Value = Format(ttk_Date(Date), "DD/MM/YYYY")  ' Date création Societe
'    c3.Cells(ligne, 7).Value = nbrowmax + 1               ' Numero client
    c3.Cells(ligne, 1).Value = Me("TextBox6").Value        ' Adresse Facturation
    c3.Cells(ligne, 2).Value = Me("TextBox30").Value       ' CP
    c3.Cells(ligne, 3).Value = Me("Textbox8").Value        ' Ville
    c3.Cells(ligne, 21).Value = Me("Textbox12").Value      ' Mail
    c3.Cells(ligne, 22).Value = Format(Me("Textbox14").Value, "0# ## ## ## ##")     ' Tel1
    c3.Cells(ligne, 23).Value = Format(Me("Textbox16").Value, "0# ## ## ## ##")      ' Tel2
    c3.Cells(ligne, 18).Value = Me("Textbox18").Value      ' TARIF
    c3.Cells(ligne, 19).Value = Me("Textbox20").Value      ' Typ_DOM
    c3.Cells(ligne, 9).Value = Me("Textbox22").Value       ' SIREN/SIRET
    c3.Cells(ligne, 12).Value = Me("Textbox24").Value      'TVA
    c3.Cells(ligne, 13).Value = Me("Textbox26").Value      'IBAN
    c3.Cells(ligne, 24).Value = Me("Combobox3").Value      'periodicité
    c3.Cells(ligne, 8).Value = Me("Combobox21").Value      'sexe
    USF_Client_edit.Repaint
End Sub
Public Sub centrage_ligne()
'                Centrage de lignes affichees dans Stock
    Dim NbColonnes, NbLignes, PremLigneVisible, PremColVisible, offsetligne As Integer
    Dim PlageVisible As String
    '                PlageVisible = (Range(Cells(PremLigneVisible, PremColVisible), _
                     '                       Cells(PremLigneVisible + NbLignes, PremColVisible + NbColonnes)).Address)
    nbrowmax = s1.Range("A65000").End(xlUp).Row
    Set plage_rech = s1.Range("A" & Trim(ActiveCell.Row) & ":A2000")
    Set c = plage_rech.Find(tt, , , xlPart)
    If Not c Is Nothing Then
        With c
            fnd1 = c.Row
            Lig = c.Row
            ligne = c.Row
            Set c = plage_rech.FindNext(c)
            fnd2 = c.Row
        End With
    End If
    Set plage = Range("A3", Cells(255, 1)).Find(tt, , xlValues, xlPart)
    ligne = Columns("A").Find(plage).Column
    For j = 1 To 7
        Worksheets(j).Activate
        If ligne > Round(NbLignes / 2) Then
            NbColonnes = Windows(1).VisibleRange.Columns.Count
            NbLignes = Windows(1).VisibleRange.Rows.Count
            offsetligne = PremLigneVisible - (ligne - (Round(NbLignes / 2)))
            ActiveWindow.ScrollRow = PremLigneVisible - offsetligne
            PremLigneVisible = ActiveWindow.ScrollRow
            PremColVisible = ActiveWindow.ScrollColumn
        End If
    Next
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
'____________________________________________________________________________________________________________________________________________________________________________________
'Private Sub ComboBox2_Change()
'    If Me("ComboBox2").Text <> "" Then
'        For ligne = 2 To nbrow
'            If c3.Cells(ligne, 1) Like "*" & R_t2 & "*" Then
'                c3.Cells(ligne, 1).Interior.ColorIndex = 43
'    '                ComboBox2.AddItem Cells(ligne, 1)
'            End If
'        Next
'    End If
'End Sub
'Private Sub UserForm_Activate()
'    Application.Wait Now + TimeValue("00:00:05") '5 secondes
'    Unload UserForm4
'End Sub
'____________________________________________________________________________________________________________________________________________________________________________________
'Private Sub TextBox2_Change()
'    nbrow = c3.Range("a65000").End(xlUp).row
'    Range("A2:A" & nbrow).Interior.ColorIndex = 2
'    '    Me.ComboBox2.Clear
'    If Me("ComboBox2").Text <> "" Then
'        For ligne = 2 To nbrow
'            If c3.Cells(ligne, 1) Like "*" & R_t2 & "*" Then
'                c3.Cells(ligne, 1).Interior.ColorIndex = 43
'    '                ComboBox2.AddItem Cells(ligne, 1)
'            End If
'        Next
'    End If
'End Sub
'Private Sub TextBox4_Change()
'    nbrow = c3.Range("a65000").End(xlUp).row
'    Range("C2:C" & nbrow).Interior.ColorIndex = 2
'    '    ComboBox4.Clear
'    If Me("ComboBox4").Text <> "" Then
'        For ligne = 2 To nbrow
'            If Cells(ligne, 1) Like "*" & R_t4 & "*" Then
'                Cells(ligne, 1).Interior.ColorIndex = 43
'    '                ComboBox4.AddItem Cells(ligne, 1)
'            End If
'        Next
'    End If
'End Sub
'Private Sub TextBox12_Change()
'    nbrow = c3.Range("a65000").End(xlUp).row
'    Range("A2:A" & nbrow).Interior.ColorIndex = 2
'    '    Me.ComboBox12.Clear
'    If Me("ComboBox12").Text <> "" Then
'        For ligne = 2 To nbrow
'            If c3.Cells(ligne, 1) Like "*" & R_c32 & "*" Then
'                c3.Cells(ligne, 1).Interior.ColorIndex = 43
'    '                ComboBox12.AddItem Cells(ligne, 1)
'            End If
'        Next
'    End If
'End Sub
'Private Sub TextBox14_Change()
'    nbrow = c3.Range("a65000").End(xlUp).row
'    Range("A2:A" & nbrow).Interior.ColorIndex = 2
'    '    Me.ComboBox14.Clear
'    If Me("ComboBox14").Text <> "" Then
'        For ligne = 2 To nbrow
'            If c3.Cells(ligne, 1) Like "*" & R_c34 & "*" Then
'                c3.Cells(ligne, 1).Interior.ColorIndex = 43
'    '                ComboBox14.AddItem Cells(ligne, 1)
'            End If
'        Next
'    End If
'End Sub

'
'    Set R_t13 = Me.Controls.Add("Forms.Textbox.1", "Textbox13", True)
'    With R_t13
'        ' Couleur = c3.Cells(ligne, i).Interior.Color
'        Me("Textbox13").BackColor = Couleur
'        Me("Textbox13").Value = "Bouchon: "
'        Me("Textbox13").FontSize = 10
'        Me("Textbox13").Top = Y
'        Me("Textbox13").Left = X + 215
'        Me("Textbox13").Width = 65
'        Me("Textbox13").Visible = True
'    End With
'
'    Set R_t14 = Me.Controls.Add("Forms.Textbox.1", "Textbox14", True)
'    With R_t14
'        ' Couleur = c3.Cells(ligne, i).Interior.Color
'        Me("Textbox14").BackColor = Couleur
'        Me("Textbox14").Value = c3.Cells(ligne, 7).Value
'        Me("Textbox14").FontSize = 10
'        Me("Textbox14").Top = Y
'        Me("Textbox14").Left = X
'        Me("Textbox14").Width = 65
'        Me("Textbox14").Visible = True
'    End With
'





