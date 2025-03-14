VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF61b 
   Caption         =   "Menu principal"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   OleObjectBlob   =   "USF61b.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF61b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit    ' USERFORM USF61 !!!!!!!!!!!!!!!!!!!
Private Sub Envoi_définitif_Click()
    send_from_server = True
    Call affiche_raz
    Call populatesheet(k)
    Call affiche_raz
    '    Call Send_Service_Message(tabmin2)
    send_from_server = True
    Call send_sheet(tab_enum)
End Sub
Public Sub Simul_Facturation_Click()
    Dim sh As Worksheet
    USF61.Show
    '    Application.WindowState = xlMinimized
    Call init_rep2
    Set sh = Sheets("modele1")
    Call affiche_raz
    Call Facture_clients(sh)
    Call affiche_raz
'    Call populatesheet(k)
'    Call affiche_raz
'    Call send_sheet
End Sub
Public Sub Userform_initialize()
    Application.WindowState = xlMinimized
    send_from_server = False
End Sub
