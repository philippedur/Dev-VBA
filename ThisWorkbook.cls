VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
    Dim yrCode As String, mnthCode As String, dyCode As String
    Dim dtcode As String
    Dim nM As Name
'    SendKeys "{NUMLOCK}"
'    USF61.Show
    With Application
        .DecimalSeparator = ","
        .ThousandsSeparator = " "
        .UseSystemSeparators = False
    End With
    ' INITIALISATION PROJET / DECLENCHEMENT TACHES LIEES AU TEMPS.
    Call init_rep2
'    Call test1_Check_des_abonnements_du_mois
    Call deactivate_all_filters
    Application.OnKey "%j", "Menu_principal"    ' Alt+j
    '    Application.WindowState = xlMinimized
    MsgBox ("Pour demarrer l'application tapez sur Alt+j ")
' DOCUMENTATION PROJET / ARCHITECTURE
' Public Sub Calc_Reel_cumul_clients()              ' MISE A JOUR COL K DANS CLIENTS A PARTIR DE Xtract-expert (Mensuel)
' Public Sub Calc_pmt_clients_graph_charge_table()  ' CHARGEMENT DE LA TABLE GESTION
' Public Sub Calc_pmt_clients_graph_cmpt_table()    ' TRAITEMENT TABLE GESTION DANS Tabmin1(), Tabmin2 POUR AFFICHAGE GRAPH
' Public Sub Graph_Tab_init(Tabmin1(), Tabmin2())   ' INIT TABLEAUX POUR GRAPHIQUES
 Call deactivate_all_filters                  'VERIF DES FILTRES ET SUPPRESSION
'change text function date code
    
With Application
    yrCode = WorksheetFunction.Rept(.International(xlYearCode), 2)
    mnthCode = WorksheetFunction.Rept(.International(xlMonthCode), 2)
    dyCode = WorksheetFunction.Rept(.International(xlDayCode), 2)
End With

'Can only add a name if it is absent
For Each nM In ThisWorkbook.Names
    If nM.Name = "dtFormat" Then
        nM.Delete
        Exit For
    End If
Next nM

dtcode = mnthCode & " " & dyCode & " " & yrCode
ThisWorkbook.Names.Add Name:="dtFormat", RefersTo:="=""" & dtcode & """", Visible:=False
End Sub


