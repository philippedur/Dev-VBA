Attribute VB_Name = "Module3"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWindow.SmallScroll Down:=-54
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$3012").AutoFilter field:=2, Criteria1:= _
        "ALPHA CONTRUCTION"
End Sub
