Attribute VB_Name = "check_iban"
Function Mod97(numero As String) As Integer
    Dim Nro As String
    Dim a As Variant, b As Variant, c As Variant, d As Variant, e As Variant, div97 As Variant
    Nro = numero
    a = 0
    b = 0
    c = 0
    e = Right(Nro, 6)
    d = Mid(Nro, Len(Nro) - 11, 6)
    Select Case Len(Nro)
    Case 13 To 20
        c = CDbl(Mid(Nro, 1, Len(Nro) - 12))
    Case 21 To 28
        c = CDbl(Mid(Nro, Len(Nro) - 19, 8))
        If Len(Nro) <> 20 Then b = CDbl(Mid(Nro, 1, Len(Nro) - 20))
    Case 29 To 38
        c = CDbl(Mid(Nro, Len(Nro) - 19, 8))
        b = CDbl(Mid(Nro, Len(Nro) - 27, 8))
        a = CDbl(Mid(Nro, 1, Len(Nro) - 28))
    Case Else
        Mod97 = 0
        Exit Function
    End Select
    div97 = Int((a * 93 + b * 73 + c * 50 + d * 27 + e Mod 97) / 97)
    Mod97 = (a * 93 + b * 73 + c * 50 + d * 27 + e Mod 97) - div97 * 97
End Function
Function convIBAN(lettre As String)
    convIBAN = (Asc(lettre) - 55)
End Function
Public Function ControleIban(ByVal LeNumIban As String) As Boolean
    If LeNumIban <> "" Then
        Dim X As String
        LeNumIban = Replace(LeNumIban, " ", "")
        LeNumIban = Right(LeNumIban, Len(LeNumIban) - 4) & Left(LeNumIban, 4)
        n = 1
        While n <= Len(LeNumIban)
            X = Mid(LeNumIban, n, 1)
            If Not IsNumeric(X) Then
                LeNumIban = Replace(LeNumIban, X, convIBAN(X), 1, 1)
            End If
            n = n + 1
        Wend
        n_iban = Mod97(LeNumIban)
        If n_iban = 1 Then
            ControleIban = True
        Else
            ControleIban = False
        End If
    End If
End Function
