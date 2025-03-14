Attribute VB_Name = "API_Pappers"
Option Explicit
Const Pappers_Token = "api_token=358489c8c617e827a75050424929d6284061473d0d03a5c8"
Const Pappers_URL = "https://api.pappers.fr/v1/entreprise"
Const Pappers_jetons = "https://api.pappers.fr/v2/suivi-jetons?api_token=358489c8c617e827a75050424929d6284061473d0d03a5c8"
Function Extraction_champ(Cha�ne As Variant, nomchamp)
    Dim PositionD�but As Long
    nomchamp = nomchamp & """:"""
    PositionD�but = InStr(Cha�ne, nomchamp) + Len(nomchamp)
    Extraction_champ = Mid(Cha�ne, PositionD�but, InStr(Right(Cha�ne, Len(Cha�ne) - PositionD�but + 1), """") - 1)
End Function
Public Sub Infos_Jur(ByVal osht As Worksheet, ByVal nom_entreprise As String, ByVal siren As String, ByVal j As Integer)
    Dim url1, url2 As String
    Dim R�ponse As String
    Dim Req As Object
    Dim i As Integer
    Dim MsgErreur As String
    Dim LibChamp As String
    Dim Champ As String
    Dim ligne As Integer
    Dim D�nomSoc As String
    Set c3 = Sheets("CLIENTS")
    Set c9 = Sheets("Clients resilies")
            nom_entreprise = Trim(c3.Range("N" & j))
            siren = Left(c3.Range("I" & j), 9)
    c3.Activate
    Lig = c3.Range("N65000").End(xlUp).Row
    j = c9.Range("N65000").End(xlUp).Row + 1
    If siren = "" Then Exit Sub
    'Constitution de l'URL � requ�ter
    url1 = Pappers_URL & "?" & Pappers_Token & "&siren=" & siren
    url2 = Pappers_jetons
    'Envoi de la requ�te au site Pappers
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "GET", url1, False
    Req.Send
    'Lecture de la r�ponse � la requ�te restitu�e par Pappers
    R�ponse = Req.responseText
    'Si la r�ponse correspond � un message d'erreur la variable MsgErreur renvoie un message
    MsgErreur = ""
    If InStr(R�ponse, "statusCode"":400,""error") > 1 Or InStr(R�ponse, "statusCode"":404,""error") > 1 Then
        MsgErreur = "SIREN inconnu !"
    ElseIf InStr(R�ponse, "statusCode"":401,""error") > 1 Then
        MsgErreur = "Token non reconnu ! Pour obtenir un Token valide, s'inscrire sur https://www.pappers.fr/api/register"
    End If
    
        'Restitution des donn�es juridiques
        For i = 1 To 10    '10 champs de donn�es restitu�s
        'Extraction des champs de donn�es
            ligne = j
            Select Case i
            Case 1:
                LibChamp = "nom_entreprise"
                Champ = Extraction_champ(R�ponse, "nom_entreprise")
                nom_entreprise = Champ
            Case 2:
                LibChamp = "siren"
                Champ = Extraction_champ(R�ponse, "siren")
                D�nomSoc = Champ
'                c3.Cells(ligne, 25).Value = Champ
            Case 3:
                LibChamp = "Capital social"
                Champ = Extraction_champ(R�ponse, "Capital social")
'                c3.Cells(ligne, 27).Value = Champ
            Case 4:
                LibChamp = "Code postal + ville"
                Champ = Extraction_champ(R�ponse, "code_postal") & " " & Extraction_champ(R�ponse, "ville")
'                c3.Cells(ligne, 28).Value = Champ
            Case 5:
                LibChamp = "code_naf"""
                Champ = Extraction_champ(R�ponse, "code_naf")
'                c3.Cells(ligne, 30).Value = Champ
            Case 6:
                LibChamp = "libelle_code_naf"
                Champ = Extraction_champ(R�ponse, "libelle_code_naf")
'                c3.Cells(ligne, 29).Value = Champ
            Case 7:
                LibChamp = "Objet social"
                Champ = Extraction_champ(R�ponse, "objet_social")
'                c3.Cells(ligne, 31).Value = Champ
            Case 8:
                LibChamp = "Date de cr�ation"
                Champ = Extraction_champ(R�ponse, "date_creation_formate")
'                c3.Cells(ligne, 32).Value = Champ
            Case 9:
                LibChamp = "G�rant"
                Champ = Extraction_champ(R�ponse, "entreprise_cessee")
'                c3.Cells(ligne, 33).Value = Champ
            Case 10:
                LibChamp = "Num�ro de TVA intracommunautaire"
                Champ = Extraction_champ(R�ponse, "numero_tva_intracommunautaire")
'                c3.Cells(ligne, 34).Value = Champ
            End Select
    
            If MsgErreur <> "" Then Champ = MsgErreur
        Next i
    ligne = j
    With osht
        If MsgErreur <> "" Then
            .Cells(ligne, 34).Value = ""
            Else
            .Hyperlinks.Add Anchor:=.Cells(ligne, 25), Address:=URLEncode("https://www.pappers.fr/entreprise/" & siren), TextToDisplay:="Pappers"
        Debug.Print nom_entreprise
        End If
    '        .Hyperlinks.Add Anchor:=c3.Cells(ligne, 34).Value, Address:="https://www.auditsi.eu/?p=9377", TextToDisplay:="Plus d'informations sur www.auditsi.eu"
    End With
End Sub
