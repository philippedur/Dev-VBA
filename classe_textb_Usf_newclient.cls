VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_textb_Usf_newclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeTextb As MSForms.TextBox
Attribute GroupeTextb.VB_VarHelpID = -1
'Public Sub GroupeTextb_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub Groupetextb_change()
'    If GroupeTextb.Name = "" Then
'    Exit Sub
    If GroupeTextb.Name = "TextBox2" Then
        USF_NewClient("Textbox3").Value = USF_NewClient("Textbox2").Value
    End If
End Sub

