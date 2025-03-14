Attribute VB_Name = "init1"
Option Explicit
Public nbCol, nbrow, nbfiles, total_files, nbrowmax, browmin, nbrowmax2, pointeur, fMax, max, ctr1, ctr2, ctr3, X, Y, i, j, k, L, M, o, n, li, ligne, col, Lig, pos, pos2, len_sstr1, len_sstr2, len_sstr3, Filter_Start, Filter_End, periodicite As Integer
Public row_offset, col_offset, top_A, top_B, top_C, top_D, height_A, height_B, height_C, height_D, r1, r2, r3, r4, r5, Top, bot, TopRow, fnd1, fnd2, fnd3, fnd4, fnd5, fnd6, fnd7, ind_fact, offset, max_affich, LastFilterRow, cptr As Integer
Public bleu, jaune, violet, combomem, TextMem, sstr1, sstr2, sstr3, sstr4, sstr5, TxtBRrow, TxtBCol, qsp, pctr, mlgts, Pml, total_R, codi, d0, d1, d2, d3, d4, buff, tmp1, rtcmbl, name_pdf, DefaultPrinter, fich, reponse, Smois, HostName As String
Public s1, s2, s3, s4, s5, s6, s7, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, p, W, F As Worksheet
Public date_creation, dtcode, d_date As Date
Public pdfjob As Object
Public ctrl As Control
Public c, d, plage, plage2, plage_rech, rfnd, rng As Range
Public list_rech, liste, premier, MotCle, ww, log_message1, log_message2, ficopen, source_range, dest_range, oSUserName, rep_pdf, oColor As String
Public t0, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, grR, total, dens, ml, tmp2, res, res1, res2, res3 As Double
Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long
Public L_t1, L_t2, L_t3, L_t4 As MSForms.Label
Public R_L1, R_L2, R_L3, R_L4, R_L5, R_L6, R_L12, R_L13, R_L14, R_L15, R_L16, R_L17 As MSForms.ListBox
Public R_C1, R_C2, R_C3, R_C4, R_C5, R_C6, R_C12, R_C13, R_C14, R_C15, R_C16, R_C17 As MSForms.ComboBox
Public R_C20, R_C21, R_C22, R_C23, R_C24, R_C25, R_C26, R_C27, R_C28 As MSForms.ComboBox
Public R_t1, R_t2, R_t3, R_t4, R_t5, R_t6, R_t7, R_t8, R_t9, R_t10, R_t11, R_t12 As MSForms.TextBox
Public R_t13, R_t14, R_t15, R_t16, R_t17, R_t18, R_t19, R_t20 As MSForms.TextBox
Public R_t21, R_t22, R_t23, R_t24, R_t25, R_t26, R_t27, R_t28 As MSForms.TextBox
Public R_t29, R_t30, R_t31, R_t32, R_t33, R_t34, R_t35, R_t36 As MSForms.TextBox
Public new_row, trig, trigged, trig_fact, send_from_server, selct, printer_enable, no_record, Eff_client_valid As Boolean
Public Tab1(1 To 20, 1 To 20)
Public ColVisu, LargeurCol(), Tdate()
Public tblbd As Variant
Public colfiles As Collection
Public tblres
Public buffer1 As String * 256
Public Armatches() As String
Public Cmb(), txt(), cmb2(), txt2(), lsv(), TabSelect(), Tabmin1(), tabmin2(), Tab_Expe(), tab_enum(), strfilelist(), arrTemp(3), ResultList() As String
Public BuffLen, StockByProd, Couleur, rt1, rt2, rt3, rt4, iLigDep, iLigArr As Long
Public Stock, rt5 As Double
Public Coo_1, Coo_2, Coo_3, Coo_4, Coo_5, Coo_6, Coo_7, Coo_8 As Double
Public cle_rech, cle_rech2, cle_rech3 As Variant
Public Path1, Path2, path3, Path4, Path5, init_path, buffer2, tt, prev_tt, uu, yy, siren, list_m, Societe, nom_entreprise, strAddress As String
Public Row_IPub, Row_Start, topA, topB, topC, tRow, tcol, nbcol_s, nbrow_s, nbcol_d, nbrow_d, e_date As Integer
Public maCollection As Collection
Public local_path, file1, lnk As String
Public lrow, start_row As Long
Public vfile As Variant
Public Sub init_rep2()
    For i = 1 To 5
        If GetUserName(buffer1, BuffLen) Then _
        oSUserName = Left(buffer1, BuffLen - 1)
        HostName = Environ("computername")
    Next
    If oSUserName = "phili" And HostName = "ALPHA" Then
        Path2 = "G:\Dev-VBA\SynologyDrive\Midi-services\Send_mail_Facturation\"
        path3 = "G:\Dev-VBA\SynologyDrive\Midi-services\Send_mail_Facturation\Factures_pdf\"
'        Path4 = "M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\Factures_pdf\MIDI-SERVICES\Midi Services\Domiciliation\Documents clients\"
        ElseIf oSUserName = "phili" And HostName = "BETA" Then
        Path2 = "M:\Dev-VBA\SynologyDrive\Midi-services\Send_mail_Facturation\"
        path3 = "M:\Dev-VBA\SynologyDrive\Midi-services\Send_mail_Facturation\Factures_pdf\"
        Path4 = "M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\Factures_pdf\MIDI-SERVICES\Midi Services\Domiciliation\Documents clients\"
        ElseIf oSUserName = "Philippe Durieux" Then
        Path2 = "M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\"
'        Path2 = "E:\Dev-VBA\SynologyDrive\Midi-services\Send_mail_Facturation\"
        path3 = Path2 & "Factures_pdf\"
'        Path2 = "M:\\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\"
        Path2 = "M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\Factures_pdf"
        ElseIf oSUserName = "Pierre" Then
        Path2 = "M:\MIDI-SERVICES\philippe\MAINTENANCE-PHILIPPE\Softwares\apps\Facturation\"
        path3 = Path2 & "Factures_pdf\"
        Path4 = "M:\MIDI-SERVICES\Midi Services\Domiciliation\Documents clients\"
    End If
    '    file1 = Path2 & "Fiche Modele Facture MIDI-SERVICES-01.xlsm"
    local_path = path3    '   "\\PIERRE-HP\Users\Public\Scan" & "\"
    Set c1 = Sheets("modele1")
    Set c2 = Sheets("Travaux")
    Set c3 = Sheets("CLIENTS")
    Set c4 = Sheets("TYP_dom")
    Set c5 = Sheets("expe")
    Set c6 = Sheets("EBP-Xtract-expert")
    Set c7 = Sheets("Buff2")
    Set c8 = Sheets("Gestion")
    Set c9 = Sheets("Clients resilies")
    Set c10 = Sheets("Buff3")
    Call Deactivation_PDFCreator
End Sub




