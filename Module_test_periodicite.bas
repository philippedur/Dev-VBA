Attribute VB_Name = "Module_test_periodicite"
Option Explicit
Public Sub test1_Check_des_abonnements_du_mois()
    Erase tabmin2, Tabmin1
    Set c10 = Worksheets("ref1")
    If c10.Cells(Val(Month(Date)), 2) = "F" Then
        Dim inp
        Call init_rep2
        ficopen = path3 & "Clients_A_Facturer_Du_Mois_Log.txt"
        log_message2 = "Liste des Entreprises non mensualisees pour le mois de : " & UCase(Format(Date, "mmmm")) & vbCrLf
        Call log_txt(log_message2)
        Call tri_col_generic(Sheets("CLIENTS"), 14)
        Set c2 = Sheets("Travaux")
        Set c3 = Sheets("CLIENTS")
        Set c4 = Sheets("TYP_trav")
        Set c5 = Sheets("TYP_dom")
        Lig = c3.Range("N65000").End(xlUp).Row
        With c3
            Dim calc2 As Integer, date_
            date_ = Array("01/01/2022", "01/02/2022", "01/03/2022", "01/04/2022", "01/05/2022", "01/06/2022", "01/07/2022", "01/08/2022", "01/09/2022", "01/10/2022", "01/11/2022", "01/12/2022")
            o = 1
            Societe = ""
            M = 1
            For j = 2 To Lig
                periodicite = c3.Range("X" & j)
                If periodicite <> 12 Then
                    date_creation = c3.Range("D" & j)
                    Societe = c3.Range("N" & j)
                    res = calc_period2(Societe, j, date_(k), periodicite, date_creation)
                    If trig_fact Then
                        log_message2 = list_m     ' date_creation & " " & date_(k) & " " & d0 & " " & " " & IIf(trig_fact = True, "Facture", "")
                        Call log_txt(log_message2)
                        o = o + 1
                    End If
                End If
            Next j
            c10.Cells(Val(Month(Date)), 2) = "T"
            Call Send_Service_Message(Tabmin1)
        End With
    End If
End Sub
Public Function calc_period2(ByVal Societe As String, ByVal k As Integer, ByVal date_n As Date, ByVal periodicite As Integer, ByVal date_creation As Date) As Integer
    Dim calc2, res, date_(12) As Date, ech As Integer, dom_type As String
    trig_fact = False
    d0 = Month(Date)
    d1 = Month(date_creation)
    offset = Abs(d0 - d1)
    Select Case periodicite
    Case Is = 12
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 1
        ech = 12 / periodicite
        Call test1(Societe, ech)
        dom_type = "DOMICILIATION MENSUELLE"
    Case Is = 6
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 2
        ech = 12 / periodicite
        Call test1(Societe, ech)
        dom_type = "DOMICILIATION SEMESTRIELLE"
    Case Is = 4
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 3
        ech = 12 / periodicite
        Call test1(Societe, ech)
        dom_type = "DOMICILIATION TRIMESTRIELLE "
    Case Is = 3
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 4
        ech = 12 / periodicite
        Call test1(Societe, ech)
        dom_type = "DOMICILIATION QUADRIMESTRIELLE"
    Case Is = 2
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 6
        ech = 12 / periodicite
        Call test1(Societe, ech)
        dom_type = "DOMICILIATION BIMENSUELLE"
    Case Is = 1
        calc_period2 = offset Mod (12 / periodicite)
        trig_fact = IIf(calc_period2 = 0, True, False)
        ind_fact = 12
        ech = 12 / periodicite
        Call test1(Societe, ech)
    Case Else
        trig_fact = False
        dom_type = "DOMICILIATION ANNUELLE"
    End Select
End Function
Public Sub log_txt(inp)
    Open ficopen For Append As #1
    '    Print #1, log_message1
    Print #1, log_message2
    Close #1
End Sub
Public Sub calc_cumul_by_client()  '  DERIVEE
    Dim inp
    Call init_rep2
    ficopen = Path2 & "Facturations_Log.txt"
    Dim Societe As String, t2 As Integer, res As Integer
    Call tri_col_generic(Sheets("CLIENTS"), 14)
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_trav")
    Set c5 = Sheets("TYP_dom")
    Lig = c3.Range("S65000").End(xlUp).Row
    With c3
        Dim calc2 As Integer, date_
        date_ = Array("05/01/CLIENTS", "05/02/CLIENTS", "05/03/CLIENTS", "05/04/CLIENTS", "05/05/CLIENTS", "05/06/CLIENTS", "05/07/CLIENTS", "05/08/CLIENTS", "05/09/CLIENTS", "05/10/CLIENTS", "05/11/CLIENTS", "05/12/CLIENTS")
        For j = 2 To 10    ' Lig
            For i = 0 To Month(Date)
                periodicite = c3.Cells(j, 24)
                date_creation = c3.Cells(j, 4)
                Societe = c3.Cells(j, 14)
                res = calc_period2(j, date_(i), periodicite, date_creation)
    '                Debug.Print societe; " "; date_creation & " " & date_(i) & " " & d0 & " " & res & " " & IIf(trig_fact = True, "Facture", "")
                log_message2 = Societe & " " & list_m
                Call log_txt(log_message2)
            Next i
        Next j
    End With
End Sub
Public Function IsInArray(ByVal Societe As String, ByVal ech As String, valToBeFound As Variant, arr1(), arr2()) As Boolean
    Dim element, ofs, i_fact
    Erase tabmin2
    list_m = ""
    i_fact = 0
    L = 0
    ofs = offset
    On Error GoTo IsInArrayError:    ' err array vide
    element = 1
    L = 1
    For i = Month(date_creation) To (Month(date_creation) + (ech * periodicite)) Step (12 / periodicite)
        ReDim Preserve Tabmin1(1 To 10, 1 To 12)  ' TAB VALEURS MIDI/EBP
        ReDim Preserve tabmin2(1 To 10, 1 To 12)   ' TAB VALEURS MIDI/EBP
        tabmin2(L, 1) = Societe
        tabmin2(L, 2) = arr2(IIf(i > 12, i - 12, i))
        tabmin2(L, 3) = arr1(IIf(i > 12, i - 12, i))
        tabmin2(L, 4) = L
        tabmin2(L, 5) = ech
        If Val((tabmin2(L, 3))) = Val(Month(Date)) Then
            Tabmin1(M, 1) = tabmin2(L, 1)
            Tabmin1(M, 2) = tabmin2(L, 2)
            Tabmin1(M, 3) = tabmin2(L, 3)
            Tabmin1(M, 4) = tabmin2(L, 4)
            Tabmin1(M, 5) = tabmin2(L, 5)
            Tabmin1(M, 6) = periodicite
             list_m = list_m & " Societe " & tabmin2(L, 1) & " ----  Echeance  " & tabmin2(L, 2) & "      (" & tabmin2(L, 4) & " / " & periodicite & " )"
            M = M + 1
        End If
       L = L + 1
        If i > 12 + ofs Then i = i - 12
        If L > periodicite Then Exit Function
    Next i

    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
Public Sub test1(ByVal Societe As String, ByVal ech As Integer)
    Dim arr1(), arr2(), element
    element = "0" & CStr(ech)
    arr1 = Array("", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    arr2 = Array("", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")
    If IsInArray(Societe, ech, element, arr1(), arr2()) Then Debug.Print list_m
End Sub
