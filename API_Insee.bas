Attribute VB_Name = "API_Insee"
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Dim nomchamp As String
Const url2 As String = "https://recherche-entreprises.api.gouv.fr/search?q=la%20poste&page=1&per_page=1"
Const url1 As String = "https://recherche-entreprises.api.gouv.fr/search?q=&page=20&per_page=20"
Dim PositionDébut As Long

Function Extraction_champ(Chaîne As Variant, nomchamp As String) As String
    nomchamp = nomchamp & """:"""
    PositionDébut = InStr(Chaîne, nomchamp) + Len(nomchamp)
    Extraction_champ = Mid(Chaîne, PositionDébut, InStr(Right(Chaîne, Len(Chaîne) - PositionDébut + 1), """") - 1)
End Function
Sub test5()
    Call init_rep2
    Call Infos_Jur2(Sheets("CLIENTS"), "ABC")
End Sub
Public Sub Infos_Jur2(ByVal osht As Worksheet, ByVal nom_entreprise As String, Optional ByVal siren As String, Optional ByVal k As Integer)
    Dim url1 As String, url2 As String, url3 As String, q As String
    Dim Réponse As String
    Dim Req As Object
    Dim i As Integer
    Dim MsgErreur As String
    Dim LibChamp As String
    Dim Champ As String
    Dim ligne As Integer
    Dim DénomSoc As String

    Set c3 = Sheets("CLIENTS")
    Set c9 = Sheets("Clients resilies")
    nbrowmax = c3.Range("N" & c3.Rows.Count).End(xlUp).Row

    c3.Activate
    Set rng = c3.Range("N2:N2000")

    For k = 2 To nbrowmax
            nom_entreprise = Trim(c3.Range("N" & k))
            q = c3.Range("I" & k)
            ligne = k
            url2 = "https://recherche-entreprises.api.gouv.fr/search?q=" & q & "&per_page=5"
            url1 = "https://recherche-entreprises.api.gouv.fr/search?q=" & q & "&per_page=10"
            siren = q

            If siren <> "" Then
                ' Envoi de la requête au site
                Set Req = CreateObject("MSXML2.ServerXMLHTTP")
                Req.Open "GET", url1, False
                Req.Send
                DoEvents
                Sleep 200
                ' Lecture de la réponse
                Réponse = Req.responseText

                ' Gestion des erreurs
                MsgErreur = ""
                If InStr(Réponse, "title"":400,""Bad Request") > 1 Or InStr(Réponse, "statusCode"":404,""error") > 1 Then
                    MsgErreur = "SIREN inconnu !"
                ElseIf InStr(Réponse, "statusCode"":401,""error") > 1 Then
                    MsgErreur = "Token non reconnu ! Pour obtenir un Token valide, s'inscrire sur https://www.pappers.fr/api/register"
                End If

                With osht
                    If MsgErreur <> "" Then
                        .Cells(ligne, 34).Value = ""
                    Else
                        .Hyperlinks.Add Anchor:=c3.Range("AA" & k), Address:=URLEncode(url1), TextToDisplay:="Info-Insee--" & " " & nom_entreprise
                    End If
                End With
            End If
    Next k
    MsgBox ("operation terminée ! ")
End Sub

