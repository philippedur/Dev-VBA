Attribute VB_Name = "print_http_Xract"
Option Explicit    ' print_http_Xtract!!!!!!!!!!!!!!!!!!!
Const SW_SHOWNORMAL = 1
Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
                 ByVal lpFile As String, ByVal lpParameters As String, _
                 ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Public Sub OpenUrl()
    Dim lSuccess As LongPtr
    Dim zlparent As LongPtr
'    lSuccess = ShellExecute(0, "Open", "file:///E:\Dev-VBA\Midi-services\Send_Facturation/Listing_Travaux_Log%20-%20v2.html", path3, 1)
    lSuccess = ShellExecute(zlparent, "open", "file:///M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\Etat_Clients - v1.html", "", "", SW_SHOWNORMAL)
End Sub
Public Sub Listing_stock_Alarmes_Click()
    Dim strA(1) As String * 50
    Dim strB(1) As String * 50
    Dim strC(1) As String * 50
    Dim strD(1) As String * 50
    Dim strE(1) As String * 50
    Dim stylesheet As String
    Dim text, css As String
    Dim t0, t1, t2 As Long
    Call init_rep2
    Worksheets("CLIENTS").Activate
    Set s1 = Sheets("CLIENTS")
    nbrowmax = s1.Range("N65000").End(xlUp).Row
    ficopen = Path2 & "Etat_Clients - v1.html"
    Open ficopen For Output As #1
    Dim irow As Long
    Dim tRow As Long
    Dim iStage As Integer
    Dim iCounter As Integer
    Dim iPage As Integer
    Dim lastCol As Integer
    Dim LastRow As Integer
    With ActiveSheet
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "N").End(xlUp).Row
    End With
    Call cal_CAT_CAR
    Dim Tab4(6) As Integer
    Tab4(1) = 7     ' Num client
    Tab4(2) = 14    ' Entreprise
    Tab4(3) = 10    ' Theorique
    Tab4(4) = 11    ' Reel
    Tab4(5) = 18    ' Typ tarif
    Tab4(6) = 19    ' tarif
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<style type=""text/css"">"
    Print #1, "table {font-size: 15px;font-family: Optimum, Helvetica, sans-serif; border-collapse: collapse}"
    Print #1, "tr {border-bottom: thin solid #A9A9A9;}"
    Print #1, "tr:hover {background-color: CCFFFF;}"
    Print #1, "td {padding: 4px; margin: 0px; padding-left: 1px; padding-right: 0px; width: 15%; text-align: center;border-right: thin solid #A9A9A9}"
    Print #1, "th { background-color: #33FF66; color: #FFF; font-weight: bold; font-size: 28px; text-align: center;}"
    Print #1, "td:first-child { font-weight: bold; width: 5%;}"
    Print #1, "</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<table class=""table""><thead><tr class=""firstrow""><th colspan=""7"">EXTRACTION STATUS CLIENTS      </th></tr></thead><tbody>"
        Print #1, "<tr><td bgcolor="; "#33FF66"; ">" & "Num" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Entreprise" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Theorique" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Réel" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Typ_Tarif" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Tarif" & "</td>"
        Print #1, "<td bgcolor="; "#33FF66"; ">" & "Status" & "</td></tr>"
    For irow = 2 To LastRow
        t0 = Cells(irow, Tab4(3))  ' Theorique
        t1 = Cells(irow, Tab4(4))  ' Reel
        t2 = Cells(irow, Tab4(6))  ' tarif
        Print #1, "<td>" & c3.Cells(irow, Tab4(1)).Value & "</td>"
        Print #1, "<td>" & c3.Cells(irow, Tab4(2)).Value & "</td>"
        Print #1, "<td>" & c3.Cells(irow, Tab4(3)).Value & "</td>"
        Print #1, "<td>" & c3.Cells(irow, Tab4(4)).Value & "</td>"
        Print #1, "<td>" & c3.Cells(irow, Tab4(5)).Value& & "</td>"
        Print #1, "<td>" & Format(c3.Cells(irow, Tab4(6)).Value, "###0.00") & " €" & "</td>"
        Print #1, "<td bgcolor="; CStr(oColor); ">" & cmpdt(t0, t1, t2, irow) & "</td></tr>"
    Next irow
    Print #1, "table {font-size: 11px;font-family: Optimum, Helvetica, sans-serif; border-collapse: collapse}"
        Print #1, "<td>"; "</td>"
        Print #1, "<td>" & "C.A Théorique:" & "</td>"
        Print #1, "<td>" & t4 & "</td>"
        Print #1, "<td>" & "C.A Réel:" & "</td>"
        Print #1, "<td>" & t5 & "</td>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close
    Call OpenUrl
End Sub
Public Function cmpdt(ByVal t0 As String, ByVal t1 As String, ByVal t2 As String, irow) As String
    Dim Societe As String
    date_creation = c3.Cells(irow, 4)
    Societe = c3.Cells(irow, 14)
    t10 = c3.Cells(irow, 11)
    periodicite = c3.Range("X" & irow)
    res = calc_period2(irow, Date, periodicite, date_creation)
'    If (t0 = "") And (t1 = "") Or (t0 = 0) And (t1 = 0) Then Exit Function
    t10 = t1
' case study:
'    -1- Client avec périodicité <12 devant payer bientot
'    -2- Client avec périodicité <12 payant par prelevt.
'    -3- Client avec périodicité =12 en retard > 1 mois
'    -4- Client avec périodicité =12
' exprimer en 3 couleurs
    If t10 = "" Then
    Exit Function
    Else
    If t1 >= 0 Then
        cmpdt = Format(t1, "###0.00") & " €"
        oColor = "#66FF00"
    ElseIf t1 < 0 Then
        cmpdt = Format(t1, "###0.00") & " €"
        oColor = "#FF6600"
    End If
    End If
End Function
Public Function gm(texte As String) As String
    Dim G As String
    G = """"
    gm = G & texte & G
End Function
Function MSA$(ByVal chaine$)
    Const VAccent = "àáâãäåéêëèìíîïðòóôõöùúûüç-° ", VSsAccent = "aaaaaaeeeeiiiioooooouuuuc . "
    Dim Bcle&
    For Bcle = 1 To Len(VAccent)
        chaine = Replace(chaine, Mid(VAccent, Bcle, 1), Mid(VSsAccent, Bcle, 1))
    Next Bcle
    MSA = LCase(chaine)
End Function
Public Sub cal_CAT_CAR()
    Worksheets("CLIENTS").Activate
    t4 = 0
    t5 = 0
    Set c3 = Sheets("CLIENTS")
    nbrowmax = c3.Range("J65000").End(xlUp).Row
        For i = 2 To nbrowmax
        t4 = t4 + c3.Range("J" & i)
        t5 = t5 + c3.Range("K" & i)
        Next i
End Sub




