VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cellWidth_Change()
    If cellWidth.Value = True Then table100pct.Value = False
    
End Sub


Private Sub findFile_Click()
 
   ' Requires reference to Microsoft Office 11.0 Object Library.

   Dim fDialog As Office.FileDialog
   Dim varFile As Variant

   ' Clear listbox contents.
   'Me.FileList.RowSource = ""

   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog

      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
            
      ' Set the title of the dialog box.
      .Title = "Please select one or more files"

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"
      .Filters.Add "ASP files", "*.asp"
      .Filters.Add ".Net files", "*.aspx"
      .Filters.Add "Html files", "*.htm, *.html"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then

         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            filePath.text = varFile
         Next
         'MsgBox .SelectedItems.Item(0)

      
      End If
   End With
End Sub

Private Sub makeHTML_Click()
    Dim DestFile As String
    Dim htmlOut As String
    Dim FileNum As Integer
    Dim ColumnCount As Integer
    Dim RowCount As Integer
    Dim vbTableWith As String
    Dim vbTableFStyle As String
    Dim vbCellWith As String
    Dim vbCellBGColor As String
    Dim vbCellFStyle As String
    Dim vbFontColor As String
    Dim vbBold As String
    Dim vbItalic As String
    
    Dim outputObj As New DataObject
    
    If Trim(tableStyle.text) <> "" Then vbTableStyle = " style='" & tableStyle.text & "' " Else vbTableStyle = ""
    If Trim(tableClass.text) <> "" Then vbTableClass = " class='" & tableClass.text & "' " Else vbTableClass = ""
    If Trim(tableId.text) <> "" Then vbTableId = " id='" & tableId.text & "' " Else vbTableId = ""
    
    If Trim(rowStyle.text) <> "" Then vbRowStyle = " style='" & rowStyle.text & "' " Else vbRowStyle = ""
    If Trim(rowClass.text) <> "" Then vbRowClass = " class='" & rowClass.text & "' " Else vbRowClass = ""
    
    If Trim(cellStyle.text) <> "" Then vbCellStyle = " style='" & cellStyle.text & "' " Else vbCellStyle = ""
    If Trim(cellClass.text) <> "" Then vbCellClass = " class='" & cellClass.text & "' " Else vbCellClass = ""
    
            If cellWidth = True Then
                vbTableWidth = " width:" & Selection.Columns.Width & "; "
            End If
            
            If table100pct = True Then
                vbTableWidth = "  width:100%; "
            End If
            
            
            vbTableFStyle = " style='" & vbTableWidth & "' "
            
    htmlOut = "<table cellpadding=0 cellspacing=0 border=0 " & vbTableId & vbTableStyle & vbTableClass & vbTableFStyle & ">" & vbCrLf
    
    ' Loop for each row in selection.
    For RowCount = 1 To Selection.Rows.Count
      ' Loop for each column in selection.
      htmlOut = htmlOut & "<tr" & vbRowClass & vbRowStyle & ">" & vbCrLf
      For ColumnCount = 1 To Selection.Columns.Count
        
            If cellWidth = True Then
                vbCellWidth = " width:" & Selection.Cells(RowCount, ColumnCount).Width & "; "
            Else
                vbCellWith = ""
            End If
            
            
            If useFontColor = True Then
                vbFontColor = " color: " & index2Hex(Selection.Cells(RowCount, ColumnCount).Font.colorIndex) & "; "
            Else
                vbFontColor = ""
            End If
            
            If useBGColor = True Then
                vbCellBGColor = " background: " & index2Hex(Selection.Cells(RowCount, ColumnCount).Interior.colorIndex) & "; "
            Else
                vbCellBGColor = ""
            End If
            
            If useBold = True Then
                If Selection.Cells(RowCount, ColumnCount).Font.Bold = True Then
                    vbBold = " font-weight: bold; "
                End If
            Else
                vbBold = ""
            End If
            
            If useItalic = True Then
                If Selection.Cells(RowCount, ColumnCount).Font.Italic = True Then
                    vbItalic = " font-style: italic; "
                End If
            Else
                vbItalic = ""
            End If
            
                vbCellFStyle = " style='" & vbFontColor & vbCellWidth & vbCellBGColor & vbBold & vbItalic & "' "
                     
         ' Write current cell's text to file with quotation marks.
         htmlOut = htmlOut & "<td" & vbCellClass & vbCellStyle & vbCellFStyle & ">" & Selection.Cells(RowCount, _
            ColumnCount).text & "</td>"
         ' Check if cell is in last column.
         If ColumnCount = Selection.Columns.Count Then
            ' If so, then write a blank line.
            htmlOut = htmlOut & vbCrLf
         End If
      ' Start next iteration of ColumnCount loop.
      Next ColumnCount
    ' Start next iteration of RowCount loop.
    htmlOut = htmlOut & "</tr>" & vbCrLf
    Next RowCount
    htmlOut = htmlOut & "</table>" & vbCrLf
    
    'force rendering of empty cells
    If emptyCell = True Then htmlOut = Replace(htmlOut, "></td>", ">&nbsp;</td>")

    
    'Writing HTML to file
    If Trim(filePath.text) <> "" Then
    
        DestFile = filePath.text
        
        ' Obtain next free file handle number.
        FileNum = FreeFile()
    
        ' Turn error checking off.
        On Error Resume Next
    
        ' Attempt to open destination file for output.
        Open DestFile For Output As #FileNum
        ' If an error occurs report it and end.
        If Err <> 0 Then
          MsgBox Err.Description
            
          MsgBox "Cannot open filename " & DestFile
          End
        Else
            Print #FileNum, htmlOut;
            ' Close destination file.
            Close #FileNum
        End If
    End If
       
    ' Turn error checking on.
    On Error GoTo 0
    
    If copyClipboard.Value = True Then
        outputObj.SetText (htmlOut)
        outputObj.PutInClipboard
    End If
    
    End
    
End Sub

Private Sub table100pct_Change()
    If table100pct.Value = True Then cellWidth.Value = False
End Sub

Private Function GetRGB(colorIndex As Long)
    Red = colorIndex And 255
    Green = colorIndex \ 256 And 255
    Blue = colorIndex \ 256 ^ 2 And 255
    
    GetRGB = Red & ", " & Green & ", " & Blue
End Function

Private Function getHexColor(colorIndex As Long)
    Dim r As Integer
    Dim G As Integer
    Dim b As Integer
    Dim hexR As Variant
    Dim hexG As Variant
    Dim hexB As Variant
    
    'first convert to RGB value
    r = colorIndex And 255
    G = colorIndex \ 256 And 255
    b = colorIndex \ 256 ^ 2 And 255
    
    hexR = Hex(r)
        If Len(hexR) < 2 Then hexR = "0" & hexR
    hexG = Hex(G)
        If Len(hexG) < 2 Then hexG = "0" & hexG
    hexB = Hex(b)
        If Len(hexB) < 2 Then hexB = "0" & hexB
    
    getHexColor = "#" & hexR & hexG & hexB
End Function

Private Function index2Hex(index)

    Dim hexColor As String
    Dim colorTable(56) As String
    
    colorTable(1) = "#000000"
    colorTable(2) = "#FFFFFF"
    colorTable(3) = "#FF0000"
    colorTable(4) = "#00FF00"
    colorTable(5) = "#0000FF"
    colorTable(6) = "#FFFF00"
    colorTable(7) = "#FF00FF"
    colorTable(8) = "#00FFFF"
    colorTable(9) = "#800000"
    colorTable(10) = "#008000"
    colorTable(11) = "#000080"
    colorTable(12) = "#808000"
    colorTable(13) = "#800080"
    colorTable(14) = "#008080"
    colorTable(15) = "#C0C0C0"
    colorTable(16) = "#808080"
    colorTable(17) = "#9999FF"
    colorTable(18) = "#993366"
    colorTable(19) = "#FFFFCC"
    colorTable(20) = "#CCFFFF"
    colorTable(21) = "#660066"
    colorTable(22) = "#FF8080"
    colorTable(23) = "#0066CC"
    colorTable(24) = "#CCCCFF"
    colorTable(25) = "#000080"
    colorTable(26) = "#FF00FF"
    colorTable(27) = "#FFFF00"
    colorTable(28) = "#00FFFF"
    colorTable(29) = "#800080"
    colorTable(30) = "#800000"
    colorTable(31) = "#008080"
    colorTable(32) = "#0000FF"
    colorTable(33) = "#00CCFF"
    colorTable(34) = "#CCFFFF"
    colorTable(35) = "#CCFFCC"
    colorTable(36) = "#FFFF99"
    colorTable(37) = "#99CCFF"
    colorTable(38) = "#FF99CC"
    colorTable(39) = "#CC99FF"
    colorTable(40) = "#FFCC99"
    colorTable(41) = "#3366FF"
    colorTable(42) = "#33CCCC"
    colorTable(43) = "#99CC00"
    colorTable(44) = "#FFCC00"
    colorTable(45) = "#FF9900"
    colorTable(46) = "#FF6600"
    colorTable(47) = "#666699"
    colorTable(48) = "#969696"
    colorTable(49) = "#003366"
    colorTable(50) = "#339966"
    colorTable(51) = "#003300"
    colorTable(52) = "#333300"
    colorTable(53) = "#993300"
    colorTable(54) = "#993366"
    colorTable(55) = "#333399"
    colorTable(56) = "#333333"
    
    If index = xlColorIndexNone Then index = 2
    If index = xlColorIndexAutomatic Then index = 1
    hexColor = colorTable(index)
    
    index2Hex = hexColor
End Function


Private Sub vbUseFontColor_Click()

End Sub
