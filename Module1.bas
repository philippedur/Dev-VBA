Attribute VB_Name = "Module1"
'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" ( _
'                                      ByVal lpszSoundName As String, _
'                                      ByVal uFlags As Long) As Long
'Private Declare Function sndPlaySoundmem Lib "winmm.dll" Alias "PlaySoundA" ( _
'                                         ByVal lpszName As Long, _
'                                         ByVal hModule As Long, _
'                                         ByVal dwFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NOWAIT = &H2000
Function Sound() As String
    stFichier = Path2 & "\Windows User Account Control.wav"
    Call sndPlaySound(stFichier, SND_SYNC Or SND_NOWAIT)
End Function
Sub Menu_principal()
    Application.WindowState = xlMinimized
    USF61.Show
End Sub
Private Sub Worksheet_beforedoubleclick(ByVal Target As Range, Cancel As Boolean)
    Set titre = [A1:H1]
    ordretri = IIf(Target.Interior.colorIndex = 3, xlDescending, xlAscending)
    Target.CurrentRegion.Sort Key1:=Cells(2, Target.Column), Order1:=ordretri, header:=xlYes
    c2 = IIf(Target.Interior.colorIndex = 2, 2, 17)
    titre.Interior.colorIndex = 2
    Target.Interior.colorIndex = 17
End Sub
Public Sub maj_tarifs()
    Call tri_col_generic(Sheets("CLIENTS"), 14)
    Set c3 = Sheets("CLIENTS")
    Set c5 = Sheets("TYP_dom")
    Lig = c3.Range("S65000").End(xlUp).Row
    Call tab_affichage
    With c3
        For i = 2 To Lig
            cle_rech = Trim(Mid(c3.Cells(i, 19), 1, 2))
            If Find_Tarif(cle_rech, Sheets("TYP_dom"), "D2:D200", Armatches()) Then
                c3.Cells(i, 18) = c5.Cells(pos, 1)
                c3.Cells(i, 19) = c5.Cells(pos, 4)
                c3.Cells(i, 19) = Format(c3.Cells(i, 19), "###0.00") & " €"
                c3.Cells(i, 18).HorizontalAlignment = xlCenter
                c3.Cells(i, 19).HorizontalAlignment = xlCenter
            Else
                Msg = "Le tarif à  " & cle_rech & " €  ne vient pas de la table Tarifs" & vbCrLf _
                & "Il a été entré manuellement. Voulez-vous le corriger ?"
                style = vbYesNo + vbDefaultButton2
                Title = "Changement d'un tarif existant."
                Help = ""
                Ctxt = 1000
                response = MsgBox(Msg, style, Title, Help, Ctxt)
                If response = vbYes Then
                    USF_NewTarif_List.Show
                    ElseIf response = vbNo Then
                End If
            End If
        Next i
    End With
End Sub
Public Sub Affiche_pct(k, r1, Mess As String)
    USF61b.L1.Width = nbfiles / total_files * 330
    USF61b.Label10.Caption = Mess & " " & Format(Round(k / r1 * 100)) & CStr(" %")
    If Round(k / r1 * 100) < 40 Then
        USF61b.Label10.ForeColor = vbBlack
    Else
        USF61b.Label10.ForeColor = vbWhite
    End If
End Sub
Public Sub Affiche_pct_send(ByVal k As Integer, ByVal d1 As String, ByVal d2 As String, ByVal total_files As Integer, ByVal Mess As String)
    nbfiles = k
    USF61b.L11.Width = nbfiles / total_files * 330
    USF61b.Label5.Caption = "Adresse mail client :   " & d1
    USF61b.Label6.Caption = "Entreprise :   " & d2
    If Round(nbfiles / total_files * 100) < 50 Then
    USF61b.Label10.Caption = Mess & " " & Format(Round(nbfiles / total_files * 100)) & CStr(" %")
        USF61b.Label10.ForeColor = vbBlack
    Else
    USF61b.Label10.Caption = Mess & " " & Format(Round(nbfiles / total_files * 100)) & CStr(" %")
        USF61b.Label10.ForeColor = vbWhite
    End If
End Sub
Public Sub Affiche_trait1(Mess As String)
    USF61b.L1.Width = i / nbrowmax * 330
    USF61b.Label10.Caption = Mess & " " & Format(Round(i / nbrowmax * 100)) & CStr(" %")
    If Round(i / nbrowmax * 100) < 40 Then
        USF61b.Label10.ForeColor = vbBlack
    Else
        USF61b.Label10.ForeColor = vbWhite
    End If
End Sub
Public Sub affiche_raz()
    USF61b.L1.Width = 0
    USF61b.Label10.Caption = ""
End Sub
Sub Filt_BD()
    Range("A3:J25000").Select
    Selection.AutoFilter field:=2, Criteria1:="=" & Range("B2").text & "411AFASSI"    ' , _
    '    Criteria2:="=" & Range("B2").text & "706003", _
    '    Operator:=xlOr
End Sub
Public Sub comp_fichiers()
    Set c3 = Workbooks("Facturation-auto-mail-MIDI-SERVICES-01(Rians)12042021.xlsm").Worksheets("CLIENTS")
    Set c4 = Workbooks("Facturation-auto-mail-MIDI-SERVICES-01.xlsm").Worksheets("CLIENTS")
    Lig = c4.Range("S65000").End(xlUp).Row
    With c4
        For i = 2 To Lig
            cle_rech = Left(c3.Cells(i, 14), 10)
            Set c4 = Workbooks("Facturation-auto-mail-MIDI-SERVICES-01.xlsm").Worksheets("CLIENTS")
            If FindAll_OneRec(cle_rech, c4, "N2:N" & Lig, Armatches()) Then
                ''            Debug.Print "C3 " & c3.Cells(i, 14) & "  C4  " & c4.Cells(pos, 14)
            Else
                Debug.Print i, "C3 " & c3.Cells(i, 14) & "  C4.. No record !"
            End If
        Next i
    End With
End Sub
Public Sub deactivate_all_filters()
    Dim Sht As Worksheet
    For Each Sht In ActiveWorkbook.Sheets
        Sht.Activate
        Sht.AutoFilterMode = False
    Next
    Worksheets("CLIENTS").Activate
End Sub
Public Sub tab_affichage()
    Dim iarr As Integer
    Set c5 = Sheets("TYP_dom")
    c5.Activate
    nbrowmax = c5.Range("A65000").End(xlUp).Row
    ReDim tabtarifs(5, 1 To 1) As String
    Erase tabtarifs
    For iarr = 2 To nbrowmax
        ReDim Preserve tabtarifs(5, 1 To iarr + 1)
        tabtarifs(1, iarr) = c5.Cells(iarr, 1).Value
        tabtarifs(2, iarr) = c5.Cells(iarr, 2).Value
        tabtarifs(3, iarr) = c5.Cells(iarr, 3).Value
        tabtarifs(4, iarr) = c5.Cells(iarr, 4).Value
        tabtarifs(5, iarr) = c5.Cells(iarr, 5).Value
        '    Debug.Print tabtarifs(1, iArr), tabtarifs(2, iArr), tabtarifs(3, iArr), tabtarifs(4, iArr), tabtarifs(5, iArr)
    Next iarr
End Sub
Public Function ttk_Date(Sdt As Variant) As Double
    Dim T As Variant
    If Sdt = "" Then
        ttk_Date = CDbl(Date)
    Else
        If IsNumeric(Sdt) Then
            ttk_Date = CDbl(Sdt)
        Else
            T = Split(Sdt, Application.International(xlDateSeparator))
            Select Case Application.International(xlDateOrder)
                Case 0: ttk_Date = DateSerial(CInt(T(2)), CInt(T(0)), CInt(T(1))) ' 0 = month-day-year
                Case 1: ttk_Date = DateSerial(CInt(T(2)), CInt(T(1)), CInt(T(0))) ' 1 = day-month-year
                Case 2: ttk_Date = DateSerial(CInt(T(0)), CInt(T(1)), CInt(T(2))) ' 2 = year-month-day
            End Select
        End If
    End If
End Function





