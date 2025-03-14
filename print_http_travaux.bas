Attribute VB_Name = "print_http_travaux"
Option Explicit
'Public Declare Function ShellExecute _
'                         Lib "shell32.dll" Alias "ShellExecuteA" ( _
'                             ByVal hWnd As Long, _
'                             ByVal Operation As String, _
'                             ByVal Filename As String, _
'                             Optional ByVal Parameters As String, _
'                             Optional ByVal Directory As String, _
'                             Optional ByVal WindowStyle As Long = vbMinimizedFocus _
'                             ) As Long
'Public Sub OpenUrl()
'    Dim lSuccess As Long
'    lSuccess = ShellExecute(0, "Open", "file:///E:\Dev-VBA\Midi-services\Send_Facturation/Listing_Travaux_Log%20-%20v2.html", path3, 1)
'End Sub
Private Sub Listing_Travaux_Click()
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Dim strA(1) As String * 50
    Dim strB(1) As String * 50
    Dim strC(1) As String * 50
    Dim strD(1) As String * 50
    Dim strE(1) As String * 50
    Dim stylesheet As String
    Dim text, css As String
    Call init_rep2
    Call filrage_1
    nbrowmax = c2.Range("H65000").End(xlUp).Row
    ficopen = Path2 & "Listing_Travaux_Log - v2.html"
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
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Dim Tab4(9) As Integer
    Tab4(1) = 1
    Tab4(2) = 2 ' 8
    Tab4(3) = 3 ' 21
    Tab4(4) = 4 ' 22
    Tab4(5) = 5 ' 23
    Tab4(6) = 6 ' 23
    Tab4(7) = 7 ' 23
    Tab4(8) = 8 ' 23
    Tab4(9) = 9 ' 23
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<style type=""text/css"">"
    Print #1, "table {font-size: 15px;font-family: Optimum, Helvetica, sans-serif; border-collapse: collapse}"
    Print #1, "tr {border-bottom: thin solid #A9A9A9;}"
    Print #1, "td {padding: 4px; margin: 3px; padding-left: 20px; width: 25%; text-align: justify;border-right: thin solid #A9A9A9}"
    Print #1, "th { background-color: #A9A9A9; color: #FFF; font-weight: bold; font-size: 28px; text-align: center;}"
    Print #1, "td:first-child { font-weight: bold; width: 10%;}"
    Print #1, "</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<table class=""table""><thead><tr class=""firstrow""><th colspan=""2"">EXTRACTION TRAVAUX ANNUELLE      </th></tr></thead><tbody>"
    For irow = 2 To LastRow
        Print #1, "<tr><td>"; Cells(irow, Tab4(1)).Value; "</td><td>"; Cells(irow, Tab4(2)).Value; "</td><td>"; Cells(irow, Tab4(3)).Value; "</td><td>"; Cells(irow, Tab4(4)).Value; "</td><td>"; Cells(irow, Tab4(5)).Value; "</td><td>"; Cells(irow, Tab4(6)).Value; "</td><td>"; Cells(irow, Tab4(7)).Value; "</td>"; "</tr>"
    Next irow
    Print #1, "</body>"
    Print #1, "</html>"
    Close
    Call OpenUrl
End Sub
Public Function cmpdt(dlu As String) As String
    If dlu = "" Then Exit Function
    If CDate(dlu) < Date + 30 Then
        cmpdt = "<font color=" & gm("RGB(255,33,170)") & "> " & "<th>" & MSA(dlu) & "</th>"
    Else
        cmpdt = "<td>" & MSA(dlu) & "</td>"
    End If
End Function
Public Function gm(texte As String) As String
    Dim G As String
    G = """"
    gm = G & texte & G
End Function
Function MSA$(ByVal chaine$)
    Const VAccent = "‡·‚„‰ÂÈÍÎËÏÌÓÔÚÛÙıˆ˘˙˚¸Á-∞ ", VSsAccent = "aaaaaaeeeeiiiioooooouuuuc . "
    Dim Bcle&
    For Bcle = 1 To Len(VAccent)
        chaine = Replace(chaine, Mid(VAccent, Bcle, 1), Mid(VSsAccent, Bcle, 1))
    Next Bcle
    MSA = LCase(chaine)
End Function
Sub filrage_1()
    Range("A3:J502").Select
    Selection.AutoFilter field:=2, Criteria1:="=" & Range("B2").text & "706001", _
    Criteria2:="=" & Range("B2").text & "706003", _
    Operator:=xlOr
    sstr2 = "SMA"
    Selection.AutoFilter field:=7, Criteria1:="=" & sstr2 & "*", Operator:=xlAnd
End Sub
Private Sub Listing_EBP_1_Click()
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Dim strA(1) As String * 50
    Dim strB(1) As String * 50
    Dim strC(1) As String * 50
    Dim strD(1) As String * 50
    Dim strE(1) As String * 50
    Dim stylesheet As String
    Dim text, css As String
    Call init_rep2
    Call filrage_1
    nbrowmax = c2.Range("H65000").End(xlUp).Row
    ficopen = Path2 & "Listing_Travaux_Log - v2.html"
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
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Dim Tab4(9) As Integer
    Tab4(1) = 1
    Tab4(2) = 2 ' 8
    Tab4(3) = 3 ' 21
    Tab4(4) = 4 ' 22
    Tab4(5) = 5 ' 23
    Tab4(6) = 6 ' 23
    Tab4(7) = 7 ' 23
    Tab4(8) = 8 ' 23
    Tab4(9) = 9 ' 23
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<style type=""text/css"">"
    Print #1, "table {font-size: 15px;font-family: Optimum, Helvetica, sans-serif; border-collapse: collapse}"
    Print #1, "tr {border-bottom: thin solid #A9A9A9;}"
    Print #1, "td {padding: 4px; margin: 3px; padding-left: 20px; width: 25%; text-align: justify;border-right: thin solid #A9A9A9}"
    Print #1, "th { background-color: #A9A9A9; color: #FFF; font-weight: bold; font-size: 28px; text-align: center;}"
    Print #1, "td:first-child { font-weight: bold; width: 10%;}"
    Print #1, "</style>"
    Print #1, "</head>"
    Print #1, "<body>"
    Print #1, "<table class=""table""><thead><tr class=""firstrow""><th colspan=""2"">EXTRACTION TRAVAUX ANNUELLE      </th></tr></thead><tbody>"
    For irow = 2 To LastRow
        Print #1, "<tr><td>"; Cells(irow, Tab4(1)).Value; "</td><td>"; Cells(irow, Tab4(2)).Value; "</td><td>"; Cells(irow, Tab4(3)).Value; "</td><td>"; Cells(irow, Tab4(4)).Value; "</td><td>"; Cells(irow, Tab4(5)).Value; "</td><td>"; Cells(irow, Tab4(6)).Value; "</td><td>"; Cells(irow, Tab4(7)).Value; "</td>"; "</tr>"
    Next irow
    Print #1, "</body>"
    Print #1, "</html>"
    Close
    Call OpenUrl
End Sub

