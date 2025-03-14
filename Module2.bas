Attribute VB_Name = "Module2"
Sub testl()
'Call Facture_clients(Sheets("modele1"))
'Call Send_Service_Message(tabmin2)
'Call test11
'Call populatesheet(k)
'Call EnumerateFiles2(rep_pdf, "*.pdf", colfiles)
End Sub
Sub ShowDriveList()
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d In dc
        s = s & d.DriveLetter & " - "
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
        s = s & n & vbCrLf
    Next
    MsgBox s
End Sub
Private Sub test35()
Set c3 = Sheets("CLIENTS")
rng = c3.Range("N:N")
nbrowmax = c3.Range("N65000").End(xlUp).Row
    nbrowmax = c4.Range("B65000").End(xlUp).Row
For i = 2 To nbrowmax
cle_rech = wsht.Range("N" & i)
If FoundCell(cle_rech, wsht, rng) Then
res = pos
Debug.Print i, cle_rech
Else:
End If
Next i
End Sub
