Attribute VB_Name = "Send_Mail2"
Option Base 1
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Option Explicit
Public Sub Send_Ionos_Mail(ByVal k As Integer, ByVal d1 As String, ByVal d2 As String, ByVal d3 As String, ByVal d4 As String)   ' DERIVEE DE SendSimpleCDOMailWithBasicAuthentication(ByVal d1 As String)
    If d3 <> "" Then
        Dim mail As Object        ' CDO.MESSAGE
        Dim config As Object          ' CDO.Configuration
        Set mail = CreateObject("CDO.Message")
        Set config = CreateObject("CDO.Configuration")
        config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
        config.Fields(cdoSMTPServer).Value = "smtp.ionos.fr"
        config.Fields(cdoSMTPServerPort).Value = 465
        config.Fields(cdoSMTPUseSSL).Value = "true"
        config.Fields(cdoSMTPAuthenticate).Value = cdoBasic
'''        config.Fields(cdoSendUserName).Value = "philippe.durieux@tpe-connect.com"
'''        config.Fields(cdoSendPassword).Value = "Mpe##2017Mpe@@2023"
        config.Fields(cdoSendUserName).Value = "pierre.durieux@midi-services.fr"
        config.Fields(cdoSendPassword).Value = "Valdeblore06!stdalmas##"
        config.Fields.Update
        Set mail.Configuration = config
        On Error GoTo Err_Trap
        With mail
            .To = d2 ' "philippe.durieux@tpe-connect.com"
            .From = "pierre.durieux@midi-services.fr"  ' "philippe.durieux@tpe-connect.com" '
            .Subject = "Facture Sté: " & d1 & "  "  '  "Votre nouvelle facture MIDI-SERVICES est arrivée."
            .TextBody = "Bonjour," & vbCrLf & _
            "Veuillez trouver ci-jointe votre nouvelle facture" & vbCrLf & _
            "En cas d'erreur, n’hésitez pas à nous en faire part par retour de mail." & vbCrLf & _
            "Restant à votre disposition" & vbCrLf & _
            "Si vous avez déjà reçu cette facture, veuillez ne pas en tenir compte" & vbCrLf
            local_path = rep_pdf
            .AddAttachment d3
            send_from_server = True
            .Send  '   COMMENT /EVITER ENVOI
        End With
            F.Range("A" & lrow) = d4
            F.Range("C" & lrow) = d1
            F.Range("D" & lrow) = "Sent"  ' CHECK
            F.Range("E" & lrow) = Date
            F.Range("F" & lrow) = Time()   '  RECH ADRSS MAIL
            F.Range("G" & lrow) = d2  '  RECH ADRSS MAIL
            cptr = cptr + 1
Exit_Err:
        Set mail = Nothing
        Set config = Nothing
        Exit Sub
Err_Trap:
        If Err <> 0 Then
            Select Case Err.Number
                Case -2147220973  'Could be because of Internet Connection
                log_message1 = CStr(Worksheets("expe").Range("A" & k) & Worksheets("expe").Range("B" & k) & Worksheets("expe").Range("C" & k) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & k) = "Err."
                Worksheets("expe").Range("E" & k) = Date
                Worksheets("expe").Range("F" & k) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                   F.Range("D" & k) = "Err."  ' CHECK
                '                            MsgBox "Pas de liaison InInternet Connectionternet !!  -- " & Err.Description
                Case -2147220975  'Incorrect credentials User ID or password
                log_message1 = CStr(Worksheets("expe").Range("A" & k) & Worksheets("expe").Range("B" & k) & Worksheets("expe").Range("C" & k) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & k) = "Err."
                Worksheets("expe").Range("E" & k) = Date
                Worksheets("expe").Range("F" & k) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                  F.Range("D" & k) = "Err."  ' CHECK
                '                            MsgBox "Mdp serveur erroné !!  -- " & Err.Description
                Case Else   'Rest other errors
                log_message1 = CStr(Worksheets("expe").Range("A" & k) & Worksheets("expe").Range("B" & k) & Worksheets("expe").Range("C" & k) & Err.Description & Date & Time)
                Worksheets("expe").Range("D" & k) = "Err."
                Worksheets("expe").Range("E" & k) = Date
                Worksheets("expe").Range("F" & k) = Time()
                log_message2 = d1 & " " & Err.Description & " " & Date & " " & Time()
                 F.Range("D" & k) = "Err."  ' CHECK
                '                            MsgBox "Erreur pendant l'envoi de la facture !!  -- " & Err.Description
            End Select
        Call log_txt(log_message2)
        Resume Exit_Err
    End If
    Set config = Nothing
    Set mail = Nothing
    End If
    End Sub
Public Sub Send_Service_Message(tabmin2)   ' DERIVEE DE SendSimpleCDOMailWithBasicAuthentication(ByVal d1 As String)
    If True Then
        Dim mail As Object        ' CDO.MESSAGE
        Dim config As Object          ' CDO.Configuration
        Set mail = CreateObject("CDO.Message")
        Set config = CreateObject("CDO.Configuration")
        config.Fields(cdoSendUsingMethod).Value = cdoSendUsingPort
        config.Fields(cdoSMTPServer).Value = "smtp.ionos.fr"
        config.Fields(cdoSMTPServerPort).Value = 465
        config.Fields(cdoSMTPUseSSL).Value = "true"
        config.Fields(cdoSMTPAuthenticate).Value = cdoBasic
        config.Fields(cdoSendUserName).Value = "philippe.durieux@tpe-connect.com"
        config.Fields(cdoSendPassword).Value = "Mpe##2017Mpe@@2023"
'''        config.Fields(cdoSendUserName).Value = "pierre.durieux@midi-services.fr"
'''        config.Fields(cdoSendPassword).Value = "Valdeblore06!stdalmas##"
        config.Fields.Update
        Set mail.Configuration = config
        On Error Resume Next
        With mail
            .To = "philippe.durieux@tpe-connect.com"
            .From = "philippe.durieux@tpe-connect.com" ' "pierre.durieux@midi-services.fr"
            .Subject = "MESSAGE DE SERVICE FACTURATION-MS "
            .TextBody = "CLIENTS FACTURES NON-MENSUELS AU MOIS DE :     " & comp_mois_rev(Month(Date)) & " " & vbCrLf
            For i = 1 To UBound(tabmin2)
                .TextBody = .TextBody & vbCrLf & tabmin2(i) & vbCrLf
            Next i
            send_from_server = True
            .Send  '   COMMENT /EVITER ENVOI
        End With
    Else
        Set config = Nothing
        Set mail = Nothing
    End If
    Set config = Nothing
    Set mail = Nothing
End Sub
