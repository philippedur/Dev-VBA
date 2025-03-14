VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF61 
   Caption         =   "Menu principal"
   ClientHeight    =   930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15165
   OleObjectBlob   =   "USF61.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cmd1_Click()
    Cmd1.ForeColor = &HC000&      '  Jaune
    Cmd2.ForeColor = &H4000&      '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &H4000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    new_row = True
    trig = False
    USF_NewClient.Show
End Sub
Private Sub Cmd2_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &HC000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &H4000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    new_row = False
    trig = False
    Set c3 = Sheets("CLIENTS")
    USF_newjob.Show
End Sub
Private Sub Cmd3_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &HC000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &H4000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    new_row = False
    trig = False
    USF_Client.Show
End Sub
Private Sub Cmd4_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &HC000&          '  Violet
    Cmd5.ForeColor = &H4000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    new_row = True
    trig = False
    USF_Client_edit.Show
End Sub
Private Sub Cmd5_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &HC000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    USF61b.Show
End Sub
Private Sub Cmd6_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &H4000&            '  Violet
    Cmd6.ForeColor = &HC000&          '  Violet
    USF_Eff_client.Show
End Sub
Private Sub Inst_Fact_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &HC000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    USF_Inst_fact.Show
End Sub
Private Sub Dashboard_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &HC000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    USF_simul_gestion.Show
End Sub
Private Sub Gestion_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &HC000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    Call Listing_stock_Alarmes_Click
End Sub
Private Sub Tarifs_Click()
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &H4000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &HC000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet
    USF_NewTarif_List.Show
End Sub

Public Sub Userform_initialize()
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Cmd1.ForeColor = &H4000&      '  Jaune
    Cmd2.ForeColor = &H4000&          '  Violet
    Cmd3.ForeColor = &HC000&          '  Violet
    Cmd4.ForeColor = &H4000&          '  Violet
    Cmd5.ForeColor = &H4000&          '  Violet
    Cmd6.ForeColor = &H4000&          '  Violet²²
    trig = False
    Application.WindowState = xlMinimized
End Sub

