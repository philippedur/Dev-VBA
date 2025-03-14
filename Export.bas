Attribute VB_Name = "Export"
Option Explicit
'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
Public Sub ExportAllComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, FName As String, ext As String
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If ActiveWorkbook.Path = "" Then
        MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        Exit Sub
    End If
    Call init_rep2
    '    destDir = Path2
    destDir = Path2 & "export_01012025"
'      destDir = ActiveWorkbook.path & "\" & ActiveWorkbook.Name & "\Modules"
'      destDir = ActiveWorkbook.path & "\" & ActiveWorkbook.Name & " Modules"
    If Len(Dir(destDir, vbDirectory)) > 0 Then
    End If
    '    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    '        path3 = Path2 & Date & "\"
    '        MkDir (destDir)
    '    End If
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
    'Determine the standard extention of the exported file.
    'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_Document: ext = ".cls"
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_MSForm: ext = ".frm"
            Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                FName = destDir & "\" & VBComp.Name & ext
    'Overwrite the existing file
    'Alternatively, you can prompt the user before killing the file.
                If Dir(FName, vbNormal) <> vbNullString Then Kill (FName)
                VBComp.Export (FName)
            End If
        End If
    Next VBComp
    Call Compt_lignes
End Sub
Sub Compt_lignes()
    Dim v As Object, i As Long
    For Each v In ActiveWorkbook.VBProject.VBComponents
        i = i + v.CodeModule.CountOfLines
    Next v
    MsgBox i & " lignes de code dans le classeur " & ActiveWorkbook.Name
End Sub



