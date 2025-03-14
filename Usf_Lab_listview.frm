VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Lab_listview 
   Caption         =   "UserForm1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "Usf_Lab_listview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Lab_listview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit    ' USERFORM UsF_Lab_listview !!!!!!!!!!!!!!!!!!!
''''Private Declare PtrSafe Function FindWindowA Lib "User32" _
''''        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
''''Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" _
''''        (ByVal hWnd As Long, ByVal wMsg As Long, _
''''         ByVal wParam As Long, lParam As Any) As Long
''''Private Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
''''Private Const WM_NCLBUTTONDOWN = &HA1
''''Private Const HTCAPTION = 2
''''Private Ht, Ncol As Integer
''''Dim chaine As String
''''Dim TB()
''''Dim Cmb() As New classe_combo_Usf_Lab_listview
''''Dim lsv() As New classe_listv_Usf_Lab_listview
''''Public Couleur As Double
''''Private Sub Quitter_Click()
''''    Unload Usf_Lab_listview
''''    Set Usf_Lab_listview = Nothing
''''    USF61.Repaint
''''End Sub
''''Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
''''    Unload Usf_Lab_listview
''''    Set Usf_Lab_listview = Nothing
''''    USF61.Repaint
''''End Sub
''''
''''Public Sub UserForm_Initialize()
''''    Usf_Lab_listview.BackColor = GetLongFromRGB(192, 224, 192)
''''    Usf_Lab_listview.Frame1.BackColor = GetLongFromRGB(192, 224, 192)
''''    Usf_Lab_listview.Print_to_DefaultPrinter.Enabled = True
''''    Usf_Lab_listview.Print_to_DefaultPrinter.BackColor = GetLongFromRGB(192, 224, 192)
''''    Set c2 = Sheets("Travaux")
''''    Set c3 = Sheets("CLIENTS")
''''    Dim ColVisu, LargeurCol()
''''    nbrowmax = c3.Range("B65000").End(xlUp).Row
''''    Application.EnableEvents = True
''''    printer_enable = False
''''    Call Affiche_RAZ
''''    x = 3
''''    Y = 35
''''    i = 1
''''    Application.EnableEvents = True
''''    nbrowmax = c3.Range("N65000").End(xlUp).Row
''''    ' Couleur = GetLongFromRGB(192, 224, 192)
''''    c3.Activate
''''    Set R_C2 = Me.Controls.Add("Forms.ComboBox.1", "ComboBox2", True)
''''    With R_C2
''''    '        ' Couleur = C1.Cells(ligne, i).Interior.Color
''''    '        Me("ComboBox2").BackColor = Couleur
''''        Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
''''        Me("ComboBox2").FontSize = 10
''''        Me("ComboBox2").Top = Y
''''        Me("ComboBox2").Left = x
''''        Me("ComboBox2").Width = 200
''''        Me("ComboBox2").Visible = True
''''    '        Me("ComboBox2").Clear
''''        Me("ComboBox2").SpecialEffect = 0
''''        Call tri_col_generic(Sheets("CLIENTS"), 14)
''''        For L = 2 To c3.Range("N" & Rows.Count).End(xlUp).Row
''''            Me("ComboBox2").AddItem c3.Range("N" & Trim(Str(L)))
''''        Next
''''    End With
''''
''''    Set R_t3 = Me.Controls.Add("Forms.Textbox.1", "Textbox3", True)
''''    With R_t3
''''        Me("Textbox3").Value = ""    ' c3.Cells(ligne, 14)
''''        Me("Textbox3").FontSize = 15
''''        Me("TextBox3").Font.Bold = True
''''        Me("Textbox3").Top = Y    ' - 4
''''        Me("Textbox3").Left = x + 210
''''        Me("Textbox3").Width = 350
''''        Me("Textbox3").Visible = True
''''        Me("TextBox3").SpecialEffect = 0
''''    End With
''''
''''    Y = Y + 25
''''    x = 3
''''    Set R_L6 = Usf_Lab_listview.ListView1
''''    With R_L6
''''    '        Me.Listview1.RowSource = rng
''''        ColVisu = Array(1, 2, 3, 4, 5, 6, 7)                      ' Adapter
''''        LargeurCol = Array(50, 150, 150, 30, 30, 50, 50)          ' Adapter
''''        Set c2 = Worksheets("Travaux")
''''        Worksheets("Travaux").Select
''''        Set Rng = c2.Range("A2:G" & c2.Range("A65000").End(xlUp).Row)
''''        Me.ListView1.MultiSelect = 1
''''        Me.ListView1.View = lvwReport
''''        Me.ListView1.FullRowSelect = True
''''    'ListView1.ListItems(1).ListSubItems(2).ForeColor = RGB(100, 0, 100)
''''        Me.ListView1.HideColumnHeaders = False
''''    'ListView1.ListItems(1).Selected = False
''''        Set ListView1.SelectedItem = Nothing
''''        liste = ListView1.ListItems.Count
''''        ListView1.FullRowSelect = True
''''        For i = 1 To liste
''''            If ListView1.ListItems.Item(i).Selected Then
''''                MsgBox ListView1.ListItems.Item(i).Text
''''            End If
''''        Next i
''''        With Me.ListView1
''''            .CheckBoxes = True
''''            With .ColumnHeaders
''''                .Clear
''''                .Add , , "Num client", 50
''''                .Add , , "Societe", 150
''''                .Add , , "Nature travaux", 80
''''                .Add , , "Nb travaux", 50
''''                .Add , , "Prix HT", 50
''''                .Add , , "Code Travaux", 80
''''                .Add , , "Echeance", 92
''''            End With
''''            nbrowmax = c2.Range("A65000").End(xlUp).Row
''''            For i = 2 To nbrowmax
''''                With .ListItems
''''                    .Add , , c2.Cells(i, 1)
''''                End With
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 2)
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 3)
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 4)
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 5)
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 6)
''''                .ListItems(i - 1).ListSubItems.Add , , c2.Cells(i, 7)
''''            Next i
''''        End With
''''        ListView1.Gridlines = True
''''    End With
''''    For Each ctrl In Me.Controls
''''        If TypeName(ctrl) = "ComboBox" Then
''''            i = i + 1
''''            ReDim Preserve Cmb(1 To i)
''''            Set Cmb(i).GroupeCombo = ctrl
''''        End If
''''    Next ctrl
''''    On Error Resume Next
''''End Sub
''''Public Sub affich_raz()
''''    Me("ComboBox2").Value = ""    ' c3.Cells(ligne, 14)
''''    Me("TextBox3").Value = ""      ' c3.Cells(ligne, 2).Value
''''End Sub
''''Public Sub affiche1()
''''    ligne = pos
''''    Me("TextBox3").Value = Me("ComboBox2").Value      ' Repet Societe
''''End Sub
''''Public Sub affiche3()
''''    Set R_L6 = Usf_Lab_listview.ListView1
''''    Dim ColVisu(), LargeurCol(), Rng
''''    With R_L6
''''        ColVisu = Array(1, 2, 3, 4, 5, 6, 7)                      ' Adapter
''''        LargeurCol = Array(50, 150, 150, 30, 30, 50, 50)          ' Adapter
''''        Set c2 = Worksheets("Travaux")
''''        Worksheets("Travaux").Select
''''        Me.ListView1.MultiSelect = 1
''''        Me.ListView1.View = lvwReport
''''        Me.ListView1.FullRowSelect = True
''''    'ListView1.ListItems(1).ListSubItems(2).ForeColor = RGB(100, 0, 100)
''''        Me.ListView1.HideColumnHeaders = False
''''        Me.ListView1.ListItems.Clear
''''        Set ListView1.SelectedItem = Nothing
''''        liste = ListView1.ListItems.Count
''''        ListView1.FullRowSelect = True
''''        With Me.ListView1
''''            .CheckBoxes = True
''''            With .ColumnHeaders
''''                .Clear
''''                .Add , , "Num client", 50
''''                .Add , , "Societe", 150
''''                .Add , , "Nature travaux", 80
''''                .Add , , "Nb travaux", 50
''''                .Add , , "Prix HT", 50
''''                .Add , , "Code Travaux", 80
''''                .Add , , "Echeance", 92
''''            End With
''''            nbrowmax = c2.Range("A65000").End(xlUp).Row
''''            For i = 1 To UBound(Armatches())
''''                With .ListItems
''''                    .Add , , c2.Cells(Armatches(i), 1) '''''''''''
''''                End With
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 2)
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 3)
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 4)
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 5)
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 6)
''''                .ListItems(i).ListSubItems.Add , , c2.Cells(Armatches(i), 7)
''''            Next i
''''        End With
''''        ListView1.Gridlines = True
''''    End With
''''
''''End Sub
''''Public Sub Valid_Click()
''''    If Usf_Lab_listview.Valid.Enabled = True Then
''''        Set c2 = Worksheets("Travaux")
''''        With Me.ListView1
'''''            Erase TabSelect
''''            t4 = 1
''''            For i = 1 To ListView1.ListItems.Count
''''                ReDim Preserve TabSelect(1 To 7, 1 To t4)
''''                If ListView1.ListItems(i).Checked = True Then
''''                    For j = 1 To 7
''''                        TabSelect(j, t4) = c2.Cells(Armatches(i), j)
''''                        Debug.Print i, j, t4, TabSelect(j, t4)
''''                    Next j
''''                    t4 = t4 + 1
''''                End If
''''            Next i
''''        cle_rech = Me("ComboBox2").Value
''''        nbrowmax = c3.Range("N65000").End(xlUp).Row
''''        If FindAll_OneRec(cle_rech, Sheets("CLIENTS"), "N2:N" & nbrowmax, Armatches()) Then
''''            t1 = pos
''''            Call Facture_clients_unitaire(Sheets("Travaux"), t1)
''''            Else
''''            MsgBox ("Pas trouvé d'entreprise de ce nom" & vbCrLf & "Vérifiez la feuille 'CLIENTS'.")
''''        End If
''''        End With
''''    End If
''''End Sub
''''
''''Private Sub Angle1_Change()
'''''    AppliqueArrondi
''''End Sub
''''Private Sub Angle2_Change()
'''''    AppliqueArrondi
''''End Sub
''''Sub AppliqueTransp()
''''    ActiveTransparence Me.Caption, True, True, 16706187, 255
''''    '    ActiveTransparence Me.Caption, True, 50, 0, Me.BackColor, ScrollBar1.Value
''''End Sub
''''Sub AppliqueArrondi()
'''''   Angle1.Value = 60
'''''   Angle2.Value = 60
'''''    RoundCorners Me, Me.Width, Me.Height, 30, 30
''''End Sub
''''Public Sub Print_to_DefaultPrinter_click()
''''    printer_enable = True
''''End Sub
''''
''''Public Sub CreateListBoxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
''''    header.ColumnCount = body.ColumnCount
''''    header.Width = body.Width
''''    header.Left = body.Left
''''    '        header.Top = body.Top - (header.Height - 1)
''''End Sub
''''Private Sub UserForm_Activate()
'''''    Call CreateListBoxHeader(USF_Lab_listview.ListBox2, USF_Lab_listview.ListBox2, Array("Header 1", "Header 2"))
''''End Sub
'''''Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''''    Dim j As Integer
'''''    If Item.Checked = True Then
'''''            Item.ForeColor = RGB(0, 0, 255) 'Changement couleur
'''''            Item.Bold = True 'Gras
'''''            For j = 1 To Item.ListSubItems.Count
'''''                Item.ListSubItems(j).ForeColor = RGB(0, 0, 255)
'''''                Item.ListSubItems(j).Bold = True
'''''            Next j
'''''        Else
'''''            Item.ForeColor = RGB(1, 0, 0) 'Changement couleur
'''''            Item.Bold = False
'''''
'''''            For j = 1 To Item.ListSubItems.Count
'''''                Item.ListSubItems(j).ForeColor = RGB(1, 0, 0)
'''''                Item.ListSubItems(j).Bold = False
'''''            Next j
'''''    End If
'''''End Sub
''''
''''
''''
