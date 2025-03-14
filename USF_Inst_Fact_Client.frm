VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Inst_Fact_Client 
   Caption         =   "UserForm2"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20625
   OleObjectBlob   =   "USF_Inst_Fact_Client.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Inst_Fact_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    ' USERFORM UsF_Inst_fact !!!!!!!!!!!!!!!!!!!
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
    Dim txt() As New classe_textb_Usf_Inst_fact2
    Dim Cmb() As New classe_combo_Usf_Inst_fact2
    Public Couleur As Double
    Private Sub set_headers()
    USF_Inst_Fact_Client.ListBox2.RowSource = "A2:G20"
End Sub

Private Sub Quitter_Click()
    Unload USF_Inst_Fact_Client
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
Private Sub UsF_Inst_fact_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set usf = Nothing
End Sub
Public Sub Userform_initialize()    '///// PAS DE CHECK DE CELLULES VIDES
    USF_Inst_fact.BackColor = GetLongFromRGB(192, 224, 192)
    USF_Inst_fact.BackColor = GetLongFromRGB(192, 224, 192)
    USF_Inst_Fact_Client.Print_to_DefaultPrinter.Enabled = True
    USF_Inst_Fact_Client.Print_to_DefaultPrinter.BackColor = GetLongFromRGB(192, 224, 192)
    USF_Inst_Fact_Client.Valid.Enabled = False
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c6 = Sheets("TYP_trav")
    nbrowmax = c3.Range("B65000").End(xlUp).Row
    Application.EnableEvents = True
    printer_enable = False
    no_record = False
    Call set_headers
    Call affiche_raz
    '    nbCol = c1.Cells(ligne, Columns.count).End(xlToLeft).Column + 1
    X = 3
    Y = 35
    i = 1
    Application.EnableEvents = True
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    ' Couleur = GetLongFromRGB(192, 224, 192)
    '    If trig = False Then
    '        ligne = 1
    '    Else: trig = pos
    '    End If
    c3.Activate
    '    USF_Inst_fact.Show 0
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
    Set R_C3 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox3", True)
    With R_C3
        '        ' Couleur = C1.Cells(ligne, i).Interior.Color
        '        Me("ComboBox2").BackColor = Couleur
        Me("ComboBox3").Value = ""    ' c3.Cells(ligne, 14)
        Me("ComboBox3").FontSize = 10
        Me("ComboBox3").Top = Y
        Me("ComboBox3").Left = X + 515
        Me("ComboBox3").Width = 70
        Me("ComboBox3").Visible = True
        '        Me("ComboBox2").Clear
        Me("ComboBox3").SpecialEffect = 0
        '        Call tri_col_generic(Sheets("TYP_TRAV"), 9)
        For L = 2 To c6.Range("I" & Rows.Count).End(xlUp).Row
            Me("ComboBox3").AddItem c6.Range("I" & Trim(Str(L)))
        Next
    End With


    Set R_t3 = Me.Controls.Add("Forms.Textbox.1", "Textbox3", True)
    With R_t3
        Me("Textbox3").Value = ""    ' c3.Cells(ligne, 14)
        Me("Textbox3").FontSize = 15
        Me("TextBox3").Font.Bold = True
        Me("Textbox3").Top = Y    ' - 4
        Me("Textbox3").Left = X + 210
        Me("Textbox3").Width = 300
        Me("Textbox3").Visible = True
        Me("TextBox3").SpecialEffect = 0
    End With

    Y = Y + 25
    X = 3
    Set R_L6 = Me.Controls.Add("Forms.ListBox.1", "ListBox2", False)
    With R_L6
        Me.ListBox2.RowSource = ""
        ColVisu = Array(4, 6, 7, 14, 16, 21)                     ' Adapter
        LargeurCol = Array(100, 100, 100, 100, 100, 100)         ' Adapter
        Set c1 = Worksheets("CLIENTS")
        Worksheets("CLIENTS").Select
        nbrowmax = c1.Range("A65000").End(xlUp).Row
''''            USF_Inst_fact.ScrollBar1.Min = 0
''''            Me.ScrollBar1.max = nbrowmax
        Set plage = c1.Range("A2:H" & nbrowmax)
        ' Adapter
        Me("ListBox2").ColumnCount = UBound(ColVisu) + 1
        Me("ListBox2").ColumnWidths = Join(LargeurCol, ";")
        '        Me("ListBox2").List = Application.index(plage, Evaluate("Row(2:" & c2.Range("A65000").End(xlUp).Row & ")"), ColVisu)
        Me("ListBox2").FontSize = 10
        Me("ListBox2").Top = 60
        Me("ListBox2").Width = 650
        Me("ListBox2").Left = 45
        Me("ListBox2").Height = 180
        Me("ListBox2").MultiSelect = 1
        Call EnteteListBox(ColVisu, LargeurCol())
        Me.ListBox2.RowSource = ""
        i = 0
        Me.ListBox2.RowSource = plage.Address(External:=True)
    End With
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            i = i + 1
            ReDim Preserve Cmb(1 To i)
            Set Cmb(i).GroupeCombo = ctrl
        ElseIf TypeName(ctrl) = "TextBox" Then
            L = L + 1
            ReDim Preserve txt(1 To L)
            Set txt(L).GroupeTextb = ctrl
        End If
        Next ctrl
    On Error Resume Next
    affiche1
    AppliqueArrondi
    ''        AppliqueTransp
    Call RAZ_Buff3
End Sub
Public Sub affich_raz()
    '    Me("Textbox1").Value = ""    ' c3.Cells(ligne, 14)
    Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
    Me("TextBox3").Value = ""      ' c3.Cells(ligne, 2).Value
    '    Me("ComboBox6").Value = ""    ' c3.Cells(ligne, 3).Value
    '    Me("Textbox8").Value = ""    'c3.Cells(ligne, 4).Value
End Sub
Public Sub affiche1()
    Dim ColVisu, LargeurCol()
    ligne = pos
    '    Me("Textbox1").Value = ""    ' c3.Cells(ligne, 6).Value       ' Societe
    '    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = Me("ComboBox2").Value      ' Repet Societe
    USF_Inst_Fact_Client.Valid.Enabled = True
    '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
    '    Me("Textbox8").Value = c3.Cells(ligne, 20).Value        ' Ville
End Sub
Public Sub affiche2()
    ligne = pos
    '    Me("Textbox1").Value = c3.Cells(ligne, 6).Value       ' Societe
    Me("ComboBox2").Value = c3.Cells(ligne, 14).Value       ' Societe
    Me("TextBox3").Value = c3.Cells(ligne, 14).Value       ' Repet Societe
    With R_L6
        Me.ListBox2.RowSource = ""
        ColVisu = Array(4, 6, 7, 14, 16, 21)                     ' Adapter
        LargeurCol = Array(100, 100, 100, 100, 100, 100)         ' Adapter
        Set c2 = Worksheets("Travaux")
        Worksheets("Travaux").Select
        nbrowmax = c2.Range("A65000").End(xlUp).Row
        Set plage = c2.Range("A2:G" & nbrowmax)
        ' Adapter
        Me("ListBox2").ColumnCount = UBound(ColVisu) + 1
        Me("ListBox2").ColumnWidths = Join(LargeurCol, ";")
        '        Me("ListBox2").List = Application.index(plage, Evaluate("Row(2:" & c2.Range("A65000").End(xlUp).Row & ")"), ColVisu)
        Me("ListBox2").FontSize = 10
        Me("ListBox2").Top = 60
        Me("ListBox2").Width = 550
        Me("ListBox2").Left = 45
        Me("ListBox2").Height = 180
        Me("ListBox2").MultiSelect = 1
        Call EnteteListBox(ColVisu, LargeurCol())
        Me.ListBox2.RowSource = ""
        i = 0
        Me.ListBox2.RowSource = plage2.Address(External:=True)
    End With '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
    '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
    '    Me("Textbox8").Value = c3.Cells(ligne, 20).Value        ' Ville
End Sub
Public Sub affiche3()
    Set R_L6 = Me.Controls.Add("Forms.ListBox.1", "Listbox2", False)
    '    Dim ColVisu(), LargeurCol(), Rng
    With R_L6
        Set c2 = Sheets("Travaux")
        If no_record = False Then
            With R_L6
        ColVisu = Array(4, 6, 7, 14, 16, 21)                     ' Adapter
        LargeurCol = Array(100, 100, 100, 100, 100, 100)         ' Adapter
                Set c2 = Worksheets("CLIENTS")
                Worksheets("CLIENTS").Select
                nbrowmax = c2.Range("A65000").End(xlUp).Row
                Set plage = c2.Range("A2:G" & nbrowmax)
                ' Adapter
                Me("ListBox2").ColumnCount = UBound(ColVisu) + 1
                Me("ListBox2").ColumnWidths = Join(LargeurCol, ";")
                '        Me("ListBox2").List = Application.index(plage, Evaluate("Row(2:" & c2.Range("A65000").End(xlUp).Row & ")"), ColVisu)
                Me("ListBox2").FontSize = 10
                Me("ListBox2").Top = 60
                Me("ListBox2").Width = 650
                Me("ListBox2").Left = 45
                Me("ListBox2").Height = 180
                Me("ListBox2").MultiSelect = 1
                Call EnteteListBox(ColVisu, LargeurCol())
                Me.ListBox2.RowSource = ""
                Me.ListBox2.RowSource = plage2.Address(External:=True)
            End With '    Me("ComboBox6").Value = USF_newjob("combobox6").Value        ' Adresse Facturation
 
        End If
    End With
    USF_Inst_fact_Clients.ListBox2.Top = 55
End Sub
Public Sub Valid_Click()
    If USF_Inst_fact_Clients.Valid.Enabled = True Then
        Set c2 = Sheets("Travaux")
        Set c10 = Sheets("Buff3")
        nbrowmax = c2.Range("A65000").End(xlUp).Row
        ligne = nbrowmax
        Set tblbd = c2.Range("A2" & ":G" & nbrowmax)  ' COMMENTER AVEC USAGE LISTVIEW  COMMENTER AVEC USAGE LISTVIEW  COMMENTER AVEC USAGE LISTVIEW  COMMENTER AVEC USAGE LISTVIEW
        t4 = 1
        For i = 1 To Me.ListBox2.ListCount - 1
            ReDim Preserve TabSelect(1 To 8, 1 To t4)
            If Me.ListBox2.Selected(i) = True Then
                For j = 1 To 8
                    TabSelect(j, t4) = tblbd(i, j)
                Next j
                t4 = t4 + 1
            End If
            Next i
        no_record = IIf(t4 > 0, False, True)
        If no_record = False Then
            Call Facture_clients_unitaire(Sheets("Travaux"), k)
        Else
            USF_Inst_fact_Clients.Enabled = True
            Call Facture_clients_unitaire(Sheets("Travaux"), k)
        End If
    End If
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
Public Sub Print_to_DefaultPrinter_click()
    printer_enable = True
End Sub
Public Sub EnteteListBox(ColVisu, LargeurCol())
    i = 0
    X = Me("ListBox2").Left + 8
    Y = Me("ListBox2").Top - 12
    For Each c In ColVisu
        i = i + 1
        '        Me("label" & i).Caption = rng.offset(-1).Item(1, c)
        '        Me("label" & i).Top = Y
        '        Me("label" & i).Left = X
        '        Me("label" & i).Height = 24
        '        Me("label" & i).Width = largeurcol(i - 1)
        '        X = X + largeurcol(i - 1)
    Next
End Sub
Public Sub CreateListBoxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
    ' make column count match
    header.ColumnCount = body.ColumnCount
    header.ColumnWidths = body.ColumnWidths

    ' add header elements
    '        header.Clear
    '        header.AddItem
    '        Dim i As Integer
    '        For i = 0 To UBound(arrHeaders)
    '            header.List(0, i) = arrHeaders(i)
    '        Next i

    ' make it pretty
    '        body.ZOrder (1)
    '        header.ZOrder (0)
    '        header.SpecialEffect = fmSpecialEffectFlat
    '        header.BackColor = RGB(200, 200, 200)
    '        header.Height = 10

    ' align header to body (should be done last!)
    header.Width = body.Width
    header.Left = body.Left
    '        header.Top = body.Top - (header.Height - 1)
End Sub
Private Sub UserForm_Activate()
'    Call CreateListBoxHeader(USF_Inst_fact.ListBox2, USF_Inst_fact.ListBox2, Array("Header 1", "Header 2"))
End Sub
Public Sub RAZ_Buff3()
    Worksheets("buff3").Activate
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
        
End Sub
Sub clistbox()
        USF_Inst_fact_Clients.ListBox2.Clear
End Sub







