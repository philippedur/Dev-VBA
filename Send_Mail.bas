Attribute VB_Name = "Send_Mail"
Option Base 1
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public Sub send_sheet(tab_enum As Variant)
    '_______________Version balayant les fichiers pdf avec appel Send_Ionos_Mail_______________
    Set F = Sheets("expe")
    ficopen = path3 & "Transmissions_Log.txt"
    With F
        start_row = .Cells(.Rows.Count, 4).End(xlUp).Row + 2
        lrow = start_row
'        nbfiles = 1
        k = 1
        max = total_files ' UBound(StrFileList, 2) - 1
        For k = 1 To max - 1
            d1 = tab_enum(3, k)     '" Entreprise
            d2 = tab_enum(2, k)     '" Email
            d3 = tab_enum(1, k)     '" emplacement dossier /facture .pdf
            d4 = tab_enum(4, k)     '" nom facture .pdf
            Call Affiche_pct_send(k, d2, d1, total_files, Mess)
            USF61b.Label7.Visible = False
            USF61b.Label5.Caption = "Adresse mail client :   " & d2
            USF61b.Label5.Visible = True
            USF61b.Label5.ForeColor = RGB(20, 148, 20)
            USF61b.Label5.Font.Size = 11
            USF61b.Label5.Font.Name = "Calibri"
            USF61b.Label6.Visible = True
            USF61b.Label6.Caption = "Entreprise :   " & d1
            USF61b.Label6.ForeColor = RGB(20, 148, 20)
            USF61b.Label6.Font.Size = 11
            USF61b.Label6.Font.Name = "Calibri"
            USF61b.Label8.Visible = False
            USF61b.Label8.ForeColor = RGB(20, 148, 20)
            USF61b.Label8.Font.Size = 11
            USF61b.Label8.Font.Name = "Calibri"
            Call Send_Ionos_Mail(k, d1, d2, d3, d4)
            log_message2 = strf25(Date) & strf25(d1) & strf30(d2) & strf60(d3)
            Debug.Print strf60(tab_enum(1, k)) & "  " & strf30(tab_enum(2, k)) & "  " & strf25(tab_enum(3, k))
            Call log_txt(log_message2)
            lrow = lrow + 1
'            nbfiles = nbfiles + 1
            DoEvents
            Sleep 200
        Next k
        USF61b.Label8.Visible = False
   End With
     USF61b.Label8.Caption = "Expédition terminée."
    USF61b.Label8.Visible = True
    ActiveWorkbook.Save
End Sub
Public Sub populatesheet(k)
    Call init_rep2
    Call set_rep
        Call EnumerateFiles2(rep_pdf, "*.pdf", colfiles)
End Sub
Sub EnumerateFiles2(ByVal sDirectory As String, ByVal sFileSpec As String, ByRef cCollection As Collection)
    sDirectory = rep_pdf
    Dim sTemp As String
    Dim arr() As String
    sTemp = Dir(sDirectory & sFileSpec)
    total_files = 0
Dim FNames As String
    Erase arr()
        FNames = Dir$(sDirectory & sFileSpec)
        i = 1
        total_files = 0
        Do Until FNames = ""
        ReDim Preserve arr(i)
        arr(i) = FNames
        i = i + 1
        FNames = Dir
        Loop
        j = 1
        
            total_files = 1
            For j = 1 To UBound(arr)
            k = j
            If decod5(j, arr(j)) = True Then
                ReDim Preserve tab_enum(4, 1 To total_files)
    ''            TriRapid arr(), LBound(arr), UBound(arr)
    '            If arrTemp(2) <> "" Then
                tab_enum(1, total_files) = sDirectory & arr(total_files)
                Call decod4(k, arr(total_files))
                tab_enum(2, total_files) = arrTemp(2)
                tab_enum(3, total_files) = cle_rech
                tab_enum(4, total_files) = arr(j)
                total_files = total_files + 1
'            Debug.Print strf60(tab_enum(1, j)) & "  " & strf30(tab_enum(2, j)) & "  " & strf25(tab_enum(3, j)) & strf25(tab_enum(4, j))
            End If
            Next j

End Sub
Public Sub EnumerateFiles(ByVal targetFolder, ByVal sFileSpec As String, ByRef cCollection As Collection)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(rep_pdf)
    Dim foundFile As Variant
    total_files = 0
    With cCollection
    For Each foundFile In targetFolder.Files
        cCollection.Add targetFolder & foundFile ' targetFolder.Files
        total_files = total_files + 1
        Debug.Print total_files, foundFile.Name
    Next
    End With
End Sub

Function FileExist(FullName As String) As Boolean
    Dim piece_jointe As String
    FileExist = Dir(FullName) <> vbNullString
End Function
Public Function decod3(ByVal k As Integer, arr As Variant) As Variant
    If tab_enum(1, k) = "" Then
        Exit Function
    Else
    Erase arrTemp
        sstr4 = InStr(1, tab_enum(1, k), "___") + 3
        sstr1 = InStrRev(tab_enum(1, k), "__F")
        If sstr4 < sstr1 Then
        Set c3 = Sheets("CLIENTS")
'        InStrRev(vfile, "\") + 1
        cle_rech = Trim(Mid(tab_enum(1, k), sstr4, sstr1 - sstr4))
        Set rng = c3.Range("N" & k & ":N2000")
'''(rng As Range, ByVal What As Variant, Optional LookIn As XlFindLookIn = xlFormulas, Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False, Optional iDoEvents As Boolean = False) As Range
        If FindAll_ByArea(rng, cle_rech) Then
            arrTemp(2) = c3.Range("N" & pos).Value
            arrTemp(3) = c3.Range("U" & pos).Value   '  RECH ADRSS MAIL
        End If
        End If
    End If
'        decod3(k) = arrTemp
End Function
Public Sub decod4(ByVal k As Integer, arr As Variant)
    If arr = "" Then
        Exit Sub
    Else
    Erase arrTemp
        sstr4 = InStr(1, arr, "___") + 3
        sstr1 = InStrRev(arr, "__F")
        If sstr4 < sstr1 Then
        Set c3 = Sheets("CLIENTS")
'        InStrRev(vfile, "\") + 1
        sstr3 = Trim(Mid(arr, sstr4, sstr1 - sstr4))
        cle_rech = sstr3
        Set rng = c3.Range("N" & k & ":N2000")
        If FindAll_ByArea(rng, cle_rech) Then
            arrTemp(1) = c3.Range("N" & pos).Value
            arrTemp(2) = c3.Range("U" & pos).Value   '  RECH ADRSS MAIL
            arrTemp(3) = Mid(arr, InStrRev(arr, "\") + 1)
'            Else
         End If
        End If
    End If
'        decod3(k) = arrTemp
End Sub
Public Function decod5(ByVal k As Integer, arr As Variant) As Boolean
    If arr = "" Then
    decod5 = False
        Exit Function
    Else
    decod5 = False
    Erase arrTemp
        sstr4 = InStr(1, arr, "___") + 3
        sstr1 = InStrRev(arr, "__F")
        If sstr4 < sstr1 Then
        Set c3 = Sheets("CLIENTS")
'        InStrRev(vfile, "\") + 1
        sstr3 = Trim(Mid(arr, sstr4, sstr1 - sstr4))
        cle_rech = sstr3
        Set rng = c3.Range("N" & k & ":N2000")
        If FindAll_ByArea(rng, cle_rech) Then
            arrTemp(1) = c3.Range("N" & pos).Value
            arrTemp(2) = c3.Range("U" & pos).Value   '  RECH ADRSS MAIL
        If arrTemp(2) = "" Then
            decod5 = False
            ElseIf arrTemp(1) = "" Then
            decod5 = False
            ElseIf (arrTemp(1) = "") And (arrTemp(1) = "") Then
            decod5 = False
            Else
            decod5 = True
         End If
        End If
        End If
    End If
End Function
Public Function filter_clients(ByVal Societe As String) As Boolean
    If colfiles(k) = "" Then
    filter_clients = False
        Exit Function
    Else
        Set c3 = Sheets("CLIENTS")
        sstr4 = InStr(1, colfiles(k), "___") + 3
        sstr1 = InStr(sstr4, colfiles(k), "__F")
        filter_clients = True ' Trim(Mid(colfiles(k), sstr4, sstr1 - sstr4))
    With c3
    c3.Range("$N$2:$N$3000").AutoFilter field:=1, Criteria1:= _
    Societe
    nbrowmax = c3.Range("N65535").End(xlUp).Row
'Call test4
    nbrowmin = NbLignesFiltrées(Worksheets("Travaux").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
'    nbrowmin = NbLignesFiltrées(c3.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
            nbrowmin = c3.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Row
            d2 = Worksheets("CLIENTS").Range("N" & nbrowmin)
            d1 = Worksheets("CLIENTS").Range("U" & nbrowmin)    '  RECH ADRSS MAIL
            If d1 = "" Or d2 = "" Then
            filter_clients = False
            Else
            filter_clients = True
            End If

''    c2.Range("$A$2:$H" & nbrowmax).AutoFilter field:=7, Criteria1:=Smois '
''    nbrowmax = c2.Range("A65535").End(xlUp).Row
''    nbrowmin = NbLignesFiltrées(Worksheets("Travaux").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
''    DateSup = DateSerial(Year(Date), Month(Date) - 1, Day(Date))
''    dateInf = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
'''''     c2.[A1].AutoFilter field:=8, Criteria1:=">" & CDbl(dateInf) _
'''''     , Operator:=xlAnd, Criteria2:="<=" & CDbl(DateSup)             '  ##### 2 criteres
''
''     c2.[A1].AutoFilter field:=8, Criteria1:=">" & CDbl(dateInf)    '  ##### 2 criteres
''
''    nbrowmax = c2.Range("A65535").End(xlUp).Row
''    If nbrowmax > 2 Then
''    nbrowmin = NbLignesFiltrées(Worksheets("Travaux").AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address)
''    Call copy_range_to_buff(Filter_Start, Filter_End)
'    End If
    End With
    Call deactivate_C3_filters
    End If
End Function
Public Sub deactivate_C3_filters()
        c3.Activate
        c3.AutoFilterMode = False
' Worksheets("CLIENTS").Activate
End Sub
Sub test4()
valVisi = 0
Set plage = c3.Range("N2:N1000")
 valVisi = plage.SpecialCells(xlCellTypeVisible).Count

total = plage.Count
    For Each ro In Range(plage.SpecialCells(xlCellTypeVisible).Address).Rows
        Debug.Print ro.Address(0, 0)
    Next
End Sub
Sub GetFileName()
    Dim xlRow As Long
    Dim sDir As String
    Dim fileName As String
    Dim sFolder As String

    sFolder = "C:\Temp\"

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Please select a folder"
        .InitialFileName = sFolder
        .Show
        If .SelectedItems.Count <> 0 Then
            sDir = .SelectedItems(1) & "\"
            fileName = Dir(sDir, 7)

            Do While fileName <> ""
'                Range("A1").offset(xlRow) = FileName
'                xlRow = xlRow + 1
                fileName = Dir
            Loop
        End If
    End With
End Sub
Private Sub TriRapid(strArray(), intBottom As Integer, intTop As Integer)
Dim strPivot As String, strTemp As String
Dim intBottomTemp As Integer, intTopTemp As Integer
intBottomTemp = intBottom
intTopTemp = intTop
strPivot = strArray((intBottom + intTop) \ 2)
While (intBottomTemp <= intTopTemp)
While (strArray(intBottomTemp) < strPivot And intBottomTemp < intTop)
intBottomTemp = intBottomTemp + 1
Wend
While (strPivot < strArray(intTopTemp) And intTopTemp > intBottom)
intTopTemp = intTopTemp - 1
Wend
If intBottomTemp < intTopTemp Then
strTemp = strArray(intBottomTemp)
strArray(intBottomTemp) = strArray(intTopTemp)
strArray(intTopTemp) = strTemp
End If
If intBottomTemp <= intTopTemp Then
intBottomTemp = intBottomTemp + 1
intTopTemp = intTopTemp - 1
End If
Wend
If (intBottom < intTopTemp) Then TriRapid strArray, intBottom, intTopTemp
If (intBottomTemp < intTop) Then TriRapid strArray, intBottomTemp, intTop
End Sub
Sub transferts_sendMail()
    Dim fileName As String
    Dim SourcePath As String
    Dim DestinationPath As String
    Call init_rep2
    Call set_rep
    DestinationPath = "C:\Users\Pierre\Transferts_fichiers_NE_PAS_EFFACER"
    SourcePath = rep_pdf
    fileName = Dir(SourcePath & "*.pdf", vbNormal)
 
    Do While fileName <> ""
        Name SourcePath & fileName As DestinationPath & "\" & fileName
        fileName = Dir()
    Loop
End Sub



