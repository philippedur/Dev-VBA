VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Worksheet_Change(ByVal Target As Range)
    Set c5 = Sheets("TYP_dom")
    If Target.Row > 1 And (Target.Column = 4) Then
''''        Call maj_tarifs
    Else
        Exit Sub
    End If
End Sub
