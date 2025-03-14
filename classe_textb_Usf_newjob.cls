VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classe_textb_Usf_newjob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents GroupeTextb2 As MSForms.TextBox
Attribute GroupeTextb2.VB_VarHelpID = -1
'Public Sub GroupeTextb_Mousedown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Private Sub GroupeTextb2_change()
    TextMem = GroupeTextb2.Name
    txtrow = Right(TextMem, Len(TextMem) - 7)
    If GroupeTextb2.text = "" Then Exit Sub
    If txtrow = 8 Then
        cle_rech = USF_newjob("TextBox8").Value
            Call USF_newjob.affiche1
            trig = False
        End If
        If txtrow = 11 Then
'        USF_newjob("TextBox11").Enabled = False
'            Call USF_newjob.affiche1
            trig = False
        End If
        
End Sub


