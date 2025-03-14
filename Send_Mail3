Sub EnumerateFiles2(ByVal sDirectory As String, ByVal sFileSpec As String, ByRef cCollection As Collection)
    Dim sTemp As String
    Dim FNames As String
    Dim i As Long

    ' Initialize the collection
    Set cCollection = New Collection

    ' Get the list of PDF files in the directory
    FNames = Dir(sDirectory & sFileSpec)
    Do Until FNames = ""
        ' Add each file to the collection
        cCollection.Add sDirectory & FNames
        FNames = Dir
    Loop
End Sub

Sub TransferCollectionToArray(ByRef cCollection As Collection, ByRef tab_enum() As Variant, ByVal clientsFilePath As String)
    Dim i As Long
    Dim filePath As String
    Dim fileName As String
    Dim clientEmail As String
    Dim clientName As String
    Dim validEntryCount As Long
    Dim sstr4 As Long, sstr1 As Long
    Dim ws As Worksheet
    Dim foundCell As Range

    ' Set the worksheet
    Set ws = Workbooks.Open(path2).Sheets("clients")

    ' Initialize the valid entry count
    validEntryCount = 0

    ' Redimension the array to match the collection size
    ReDim tab_enum(1 To 4, 1 To cCollection.Count)

    ' Transfer elements from the collection to the array
    For i = 1 To cCollection.Count
        filePath = cCollection(i)
        fileName = Dir(filePath)

        ' Extract the company name from the file name
        sstr4 = InStr(1, fileName, "___") + 3
        sstr1 = InStrRev(fileName, "__F")

        If sstr4 < sstr1 Then
            clientName = Mid(fileName, sstr4, sstr1 - sstr4)

            ' Find the client email based on the company name
            Set foundCell = ws.Range("N:N").Find(What:=clientName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                clientEmail = ws.Cells(foundCell.Row, "U").Value
            Else
                clientEmail = ""
            End If
        Else
            clientName = ""
            clientEmail = ""
        End If

        ' Check if all fields are non-empty
        If filePath <> empty And clientEmail <> empty And clientName <> empty And fileName <> empty Then
            ' Increment the valid entry count
            validEntryCount = validEntryCount + 1

            ' Store file details in the tab_enum array
            tab_enum(1, validEntryCount) = filePath
            tab_enum(2, validEntryCount) = clientEmail
            tab_enum(3, validEntryCount) = clientName
            tab_enum(4, validEntryCount) = fileName
        End If
    Next i

    ' Resize the array to remove any empty slots
    If validEntryCount < cCollection.Count Then
        ReDim Preserve tab_enum(1 To 4, 1 To validEntryCount)
    End If

    ' Close the workbook
    ws.Parent.Close SaveChanges:=False
End Sub

Sub MainProcedure()
    Dim sDirectory As String
    Dim sFileSpec As String
    Dim cCollection As Collection
    Dim tab_enum() As Variant
    Dim i As Long
    Dim clientsFilePath As String
    Call init_rep2
    Call set_rep
        FNames = Dir$(sDirectory & sFileSpec)

    ' Define the directory and file specification
'    sDirectory = "C:\Chemin\Vers\Votre\Repertoire\"
 '   sFileSpec = "*.pdf"
    clientsFilePath = "Bureau\Facturation-auto-mail-MIDI-SERVICES-10032025.xlsx" ' Adjust the path as needed

    ' Initialize the collection
    Set cCollection = New Collection

    ' Enumerate PDF files into the collection
    Call EnumerateFiles2(sDirectory, sFileSpec, cCollection)

    ' Transfer the collection to an array with client data
    Call TransferCollectionToArray(cCollection, tab_enum, clientsFilePath)

    ' Output the array contents (for verification)
    For i = LBound(tab_enum, 2) To UBound(tab_enum, 2)
        Debug.Print "File Path: " & tab_enum(1, i)
        Debug.Print "Client Email: " & tab_enum(2, i)
        Debug.Print "Client Name: " & tab_enum(3, i)
        Debug.Print "File Name: " & tab_enum(4, i)
    Next i
End Sub
