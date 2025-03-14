Attribute VB_Name = "Module4"
Option Compare Text
Sub peupler_Dossier()
    Dim NomFich As String
    Set Dossier = CreateObject("Scripting.FileSystemObject")
    Set c3 = Sheets("CLIENTS")
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    '    Path4 = "D:\Dev-VBA\Documents clients\"
    Path4 = "M:\MIDI-SERVICES\Midi Services\Domiciliation\Documents clients\"
    i = 1
    For ligne = 2 To nbrowmax
        NomFich = URLEncode(c3.Range("N" & ligne).Value)
        Debug.Print c3.Range("N" & ligne).Value, NomFich
        Do While NomFich <> ""
            If NomFich = Dir(Path4 & NomFich, vbDirectory) Then
                If (NomFich <> "") Or (NomFich <> ".") Or (NomFich <> "..") Then
                    c3.Cells(ligne, 26) = "Infos " & NomFich
                    c3.Cells(ligne, 26).Font.color = vbBlue
                    Exit Do
                End If
            Else
                MkDir Path4 & NomFich
            End If
        Loop
    Next ligne
End Sub
Public Sub Addrep_Onerec(ByRef chaine As String, ByVal ligne)
    Dim Path4 As String
    Dim NomFich As String
    NomFich = chaine
    Set Dossier = CreateObject("Scripting.FileSystemObject")
    Set c3 = Sheets("CLIENTS")
    nbrowmax = c3.Range("N65000").End(xlUp).Row
    Path4 = "M:\MIDI-SERVICES\Midi Services\Domiciliation\Documents clients\"
    '    Path4 = "D:\Dev-VBA\Documents clients\"
    NomFich = URLEncode(c3.Range("N" & ligne).Value)
    Debug.Print c3.Range("N" & ligne).Value, NomFich
    Do While NomFich <> ""
        If NomFich = Dir(Path4 & NomFich, vbDirectory) Then
            If (NomFich <> "") Or (NomFich <> ".") Or (NomFich <> "..") Then
                c3.Cells(ligne, 26) = "Infos " & NomFich
                c3.Cells(ligne, 26).Font.color = vbBlue
                Exit Do
            End If
        Else
            MkDir Path4 & NomFich
            Call tri_col_generic(Sheets("CLIENTS"), 14)
        End If
    Loop
End Sub
Public Function URLEncode(chaine) As String
    On Error GoTo Catch
    Dim sRtn As String
    Dim sTmp As String
    Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdeéèêfghijklmnopqrstuvwxyz:/',.?=_-$()~&"
    If Len(chaine) > 0 Then
        For iloop = 1 To Len(chaine)
            sTmp = Mid(chaine, iloop, 1)
            If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                sTmp = Hex(Asc(sTmp))
                If sTmp = "20" Then
                    sTmp = " "
                    '                    sTmp = "+"
                ElseIf Len(sTmp) = 1 Then
                    sTmp = " "
                    '                    sTmp = "+"
                Else
                    sTmp = " "
                    '                    sTmp = "+"
                End If
            End If
            sRtn = sRtn & sTmp

        Next iloop
        URLEncode = Replace(sRtn, "  ", " ")
        URLEncode = Replace(sRtn, "é", "e")
        URLEncode = Replace(sRtn, "è", "e")
        URLEncode = Replace(sRtn, "â", "a")
        URLEncode = Trim(sRtn)
    End If
Finally:
    Exit Function
Catch:
    URLEncode = ""
    Resume Finally
End Function
Public Sub test3_Check_info_juridiques()
    Call init_rep2
    Call tri_col_generic(Sheets("CLIENTS"), 14)
    Set c3 = Sheets("CLIENTS")
    Lig = c3.Range("N65000").End(xlUp).Row
    c3.Activate
    k = 0
    With c3
        For j = 2 To Lig
            nom_entreprise = Trim(c3.Range("N" & j))
            siren = Trim(c3.Range("I" & j))
            If c3.Range("Y" & j) = "" Then
                Call Infos_Jur2(Sheets("CLIENTS"), nom_entreprise, siren, j)
            End If
        Next j
    End With
End Sub
Public Sub SendSimpleCDOMail()
    Dim mail As Object        ' CDO.MESSAGE
    Dim config As Object          ' CDO.Configuration
    Set mail = CreateObject("CDO.Message")
    Set config = CreateObject("CDO.Configuration")
    
    config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
    config.Fields(cdoSMTPServer).Value = "auth.smtp.1and1.fr"
    config.Fields(cdoSMTPServerPort).Value = 25
    config.Fields.Update
    
    Set mail.Configuration = config
    
    With mail
        .To = "philippe.durieux@tpe-connect.com"
        .From = "philippe.durieux@tpe-connect.com"
        .Subject = "First email with CDO"
        .TextBody = "This is the body of the first plain text email with CDO."
        
        .AddAttachment local_path & vfile
        
        .Send
        Beep
    End With
    
    Set config = Nothing
    Set mail = Nothing
    
End Sub
Public Sub SendSimpleCDOMailWithBasicAuthentication(ByVal k As Integer, ByVal d1 As String, ByVal d2 As String, ByVal vfile As String)   ' Actuelle version du robot d'expédition facturation
                '''            d1 = strfilelist(1, k)   #####    Adresse Mail   #####
                '''            d2 = strfilelist(2, k)   #####    Fichier pdf Facture   #####
                '''            d3 = strfilelist(3, k)   #####    Nom Entreprise   #####
    F.Activate
    If d1 <> Empty Then
        Dim mail As Object        ' CDO.MESSAGE
        Dim config As Object          ' CDO.Configuration
        Set mail = CreateObject("CDO.Message")
        Set config = CreateObject("CDO.Configuration")
        config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
        config.Fields(cdoSMTPServer).Value = "smtp.ionos."
        config.Fields(cdoSMTPServerPort).Value = 465
        config.Fields(cdoSMTPUseSSL).Value = "True"
        config.Fields(cdoSMTPAuthenticate).Value = cdoBasic
        config.Fields(cdoSendUserName).Value = "philippe.durieux@tpe-connect.com"
        config.Fields(cdoSendPassword).Value = "Mpe##2017Mpe@@2023"
'''        config.Fields(cdoSendUserName).Value = "pierre.durieux@midi-services.fr"
'''        config.Fields(cdoSendPassword).Value = "Valdeblore06!stdalmas##"
        config.Fields.Update
        Set mail.Configuration = config
        On Error Resume Next ' GoTo Err_Trap
        With mail
          .To = "philippe.durieux@tpe-connect.com" ' d1  "pierre.durieux@midi-services.fr"  '
            .From = "pierre.durieux@midi-services.fr"
            .Subject = "facture : " & d3 & "  "  '  "Votre nouvelle facture MIDI-SERVICES est arrivée."
            .TextBody = "Bonjour," & vbCrLf & _
            "Veuillez trouver ci-jointe votre nouvelle facture" & vbCrLf & vbCrLf & _
            "En cas d'erreur, n’hésitez pas à nous en faire part par retour de mail." & vbCrLf & vbCrLf & _
            "Restant à votre disposition" & vbCrLf & _
            "L'équipe Midi Services vous souhaite une très bonne année 2025 ! " & vbCrLf
            local_path = rep_pdf
            Debug.Print d1, d2, d3
            .AddAttachment local_path & d2
            If send_from_server = True Then
                .Send  '   COMMENT /DISABLE ENVOI
            lrow = lrow + 1
            End If
        End With
        Worksheets("expe").Range("A" & lrow) = strfilelist(2, k)
        Worksheets("expe").Range("D" & lrow) = "Sent"
        Worksheets("expe").Range("E" & lrow) = Date
        Worksheets("expe").Range("F" & lrow) = Time()
        Worksheets("expe").Range("G" & lrow) = d1
        Set mail = Nothing
        Set config = Nothing
Exit_Err:
        Set mail = Nothing
        Set config = Nothing
        Exit Sub
Err_Trap:
        If Err <> 0 Then
            Select Case Err.Number
                Case -2147220973  'Could be because of Internet Connection
                log_message1 = CStr(Worksheets("expe").Range("A" & lrow) & Worksheets("expe").Range("B" & lrow) & Worksheets("expe").Range("C" & lrow) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & lrow) = "Err."
                Worksheets("expe").Range("E" & lrow) = Date
                Worksheets("expe").Range("F" & lrow) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                '                            MsgBox "Pas de liaison InInternet Connectionternet !!  -- " & Err.Description
                Case -2147220975  'Incorrect credentials User ID or password
                log_message1 = CStr(Worksheets("expe").Range("A" & lrow) & Worksheets("expe").Range("B" & lrow) & Worksheets("expe").Range("C" & lrow) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & lrow) = "Err."
                Worksheets("expe").Range("E" & lrow) = Date
                Worksheets("expe").Range("F" & lrow) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                '                            MsgBox "Mdp serveur erroné !!  -- " & Err.Description
                Case Else   'Rest other errors
                log_message1 = CStr(Worksheets("expe").Range("A" & lrow) & Worksheets("expe").Range("B" & lrow) & Worksheets("expe").Range("C" & lrow) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & lrow) = "Err."
                Worksheets("expe").Range("E" & lrow) = Date
                Worksheets("expe").Range("F" & lrow) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                '                            MsgBox "Erreur pendant l'envoi de la facture !!  -- " & Err.Description
            End Select
            Call log_txt(log_message2)
            Resume Exit_Err
        End If
    Else
        Set config = Nothing
        Set mail = Nothing
    End If
    Set config = Nothing
    Set mail = Nothing
End Sub

Public Sub SendSimpleCDOMailWithAuthenticationAndEncryption()
    
    Dim mail As Object        ' CDO.MESSAGE
    Dim config As Object          ' CDO.Configuration
    
    Set mail = CreateObject("CDO.Message")
    Set config = CreateObject("CDO.Configuration")
    
    config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
    config.Fields(cdoSMTPServer).Value = "philippe.durieux@tpe-connect.com"
    
    config.Fields(cdoSMTPServerPort).Value = 465  ' implicit SSL - Does not work with Explicit SSL (StartTLS) usually on Port 587
    config.Fields(cdoSMTPUseSSL).Value = "true"
    
    config.Fields(cdoSMTPAuthenticate).Value = cdoBasic
    config.Fields(cdoSendUserName).Value = "philippe.durieux@tpe-connect.com"
    config.Fields(cdoSendPassword).Value = "Mpe##2017"
    
    config.Fields.Update
    
    Set mail.Configuration = config
    
    With mail
        .To = "someone@somewhere.invalid"
        .From = "me@mycompany.invalid"
        .Subject = "First email with CDO"
        .TextBody = "This is the body of the first plain text email with CDO."
        
        .AddAttachment local_path & vfile
        
        .Send
    End With
    
    Set config = Nothing
    Set mail = Nothing
    
End Sub

Public Sub SendSimpleCDOMailWithWindowsAuthentication()
    
    Dim mail As Object        ' CDO.MESSAGE
    Dim config As Object          ' CDO.Configuration
    
    Set mail = CreateObject("CDO.Message")
    Set config = CreateObject("CDO.Configuration")
    
    config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
    config.Fields(cdoSMTPServer).Value = "mail.mycompany.invalid"
    config.Fields(cdoSMTPServerPort).Value = 25
    
    ' You can use integrated Windows Authentication within a Windows Domain/Active Directory,
    ' if the mailserver supports it
    ' Set cdoSMTPAuthenticate to cdoNTLM. You don't need to supply username/password then
    config.Fields(cdoSMTPAuthenticate).Value = cdoNTLM
    
    config.Fields.Update
    
    Set mail.Configuration = config
    
    With mail
        .To = "someone@somewhere.invalid"
        .From = "me@mycompany.invalid"
        .Subject = "First email with CDO"
        .TextBody = "This is the body of the first plain text email with CDO."
        
        .AddAttachment local_path & vfile
        
        .Send
    End With
    
    Set config = Nothing
    Set mail = Nothing
    
End Sub

