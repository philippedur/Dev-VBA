VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_textb_Usf_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeTextb2 As MSForms.TextBox
Attribute GroupeTextb2.VB_VarHelpID = -1
'Public Sub GroupeTextb_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub GroupeTextb2_change()
'    If (GroupeTextb.Name = "Textbox28") Then
'        Application.EnableEvents = False
'        For i = 1 To Len(GroupeTextb.Text)
'            sstr1 = Mid(GroupeTextb.Text, i, 1)
'            If Not sstr1 Like "#" Then Exit Sub
'        Next i
'        Application.EnableEvents = True
'    End If
End Sub
Public Function start2(inp As String) As String    ' detection Flacon dans affichage
    sstr1 = InStr(1, inp, "Flacon", vbTextCompare)
    If sstr1 > 0 Then
        start2 = Trim(Left(inp, sstr1 - 2))
    End If
End Function
Public Function num_flacon(chaine) As String
    On Error GoTo Catch
    Dim iloop As Integer
    Dim sRtn As String
    Dim sTmp, htmp As String
    sRtn = ""
    sTmp = ""
    '    Const sValidChars = "1234567890,"
    If Len(chaine) > 0 Then
        iloop = InStr(1, chaine, ":", vbTextCompare)
        For i = iloop To iloop + 5
            sTmp = Mid(chaine, iloop, 1)
            If sTmp Like "#" Then
                sRtn = sRtn & sTmp
                If Mid(chaine, iloop + 1, 1) = " " Then GoTo Catch
            End If
            iloop = iloop + 1
        Next
        num_flacon = Trim(Mid(chaine, iloop - 2, 1)) & " "
    End If
Finally:
    Exit Function
Catch:
    num_flacon = Trim(Mid(chaine, 1, iloop - 2)) & " "
    Resume Finally
End Function
Function lect_titre()
    titre = c3.Cells(fnd1 + 1, 1).Value
    sstr1 = InStr(1, titre, "ml")
    sstr2 = InStr(sstr1 - 6, titre, " ")
    lect_titre = Trim(Mid(titre, sstr2, sstr1 - sstr2))
End Function

