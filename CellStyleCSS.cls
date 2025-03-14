VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellStyleCSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Style Border Order: 1=xlEdgeBottom;2=xlEdgeTop;3=xlEdgeLeft;4=xlEdgeRight;5=xlDiagonalDown;6=xlDiagonalUp
'Css Border Order: top, right, bottom,left

Private Sub appendCSS_Border(sb As Object, Style1 As style)
    Dim Item As Variant
    Dim borderStyles(1 To 4) As String, borderWidths(1 To 4) As String, borderColors(1 To 4) As String

    For Each Item In Array(2, 4, 1, 3)
        borderStyles(Item) = getBorderStyleCSSName(Style1.Borders(Item).LineStyle)
        borderWidths(Item) = getBorderWeightCSS(Style1.Borders(Item).Weight)
        borderColors(Item) = GetRGB(Style1.Borders(Item).color)
    Next

    If Len(Join(borderStyles, "")) > 0 Then sb.Append_3 "border-style:" & Join(borderStyles, " ") & ";"
    If Len(Join(borderWidths, "")) > 0 Then sb.Append_3 "border-width:" & Join(borderWidths, " ") & ";"
    If Len(Join(borderColors, "")) > 0 Then sb.Append_3 "border-color:" & Join(borderColors, " ") & ";"
End Sub

Private Sub appendCSS_Color(sb As Object, Style1 As style)
    sb.Append_3 "color:" & GetRGB(Style1.Font.color) & ";"
End Sub

'font    Sets all the font properties in one declaration
'font: font-style font-variant font-weight font-size/line-height font-family|caption|icon|menu|message-box|small-caption|status-bar|initial|inherit;
'font: italic bold 12px/30px Georgia, serif;
'font-family: "Times New Roman", Times, serif;
'font-family Specifies the font family for text
'font-size   Specifies the font size of text
'font-style  Specifies the font style for text
'font-variant    Specifies whether or not a text should be displayed in a small-caps font
'font-weight Specifies the weight of a font

Private Sub appendCSS_Font(sb As Object, Style1 As style)
    With Style1
        sb.Append_3 "Font:"
        With .Font
            If .Bold Then sb.Append_3 "Bold "
            If .Italic Then sb.Append_3 "Italic "
            sb.Append_3 .Size & "pt "
            sb.Append_3 .Name
            sb.Append_3 ";"
        End With
    End With
End Sub

'text-decoration: none|underline|overline|line-through|initial|inherit;
'none    Defines a normal text. This is default
'underline   Defines a line below the text
'overline    Defines a line above the text
'line-through    Defines a line through the text
'initial Sets this property to its default value. Read about initial
'inherit Inherits this property from its parent element. Read about inherit

Private Sub appendCSS_TextDecoration(sb As Object, Style1 As style)
    With Style1
        With .Font
            If .Strikethrough Then
                sb.Append_3 "text-decoration:line-through;"
            ElseIf .Underline Then
                sb.Append_3 "text-decoration:underline;"
            End If
        End With
    End With
End Sub

Private Sub appendCSS_TextIndent(sb As Object, Style1 As style)
    sb.Append_3 "text-indent:" & GetRGB(Style1.IndentLevel) * 9 & "px;"
End Sub

'word-wrap: normal|break-word|initial|inherit;

Private Sub appendCSS_WordWrap(sb As Object, Style1 As style)
    sb.Append_3 "word-wrap:" & IIf(Style1.WrapText, "break-word", "normal") & ";"
End Sub

Private Function getBorderStyleCSSName(BorderLineStyle As Variant) As String
'Border CSS Line Styles: dotted, dashed, solid, double, groove, ridge, inset, outset
    Select Case BorderLineStyle
        Case XlLineStyle.xlContinuous
            getBorderStyleCSSName = "solid"
        Case XlLineStyle.xlDash
            getBorderStyleCSSName = "dashed"
        Case XlLineStyle.xlDashDot
            getBorderStyleCSSName = "dashed"
        Case XlLineStyle.xlDashDotDot
            getBorderStyleCSSName = "dotted"
        Case XlLineStyle.xlDot
            getBorderStyleCSSName = "dotted"
        Case XlLineStyle.xlDouble
            getBorderStyleCSSName = "double"
        Case XlLineStyle.xlLineStyleNone
            getBorderStyleCSSName = ""
        Case XlLineStyle.xlSlantDashDot
            getBorderStyleCSSName = "ridge"
        Case Else
            getBorderStyleCSSName = "solid"
    End Select
End Function

Private Function getBorderWeightCSS(Weight As XlBorderWeight) As String
    getBorderWeightCSS = Switch(Weight = xlHairline, "0.4pt", Weight = xlMedium, "1.0pt", Weight = xlThick, "1.5pt", Weight = xlThin, "0.5pt", True, "")
End Function

' https://stackoverflow.com/questions/6003324/how-do-i-get-the-corresponding-hex-value-of-an-rgb-color-in-excel-vba

Function GetRGB(ByVal color As Long) As String
    Dim hexCode As String, hexColor As String, r As String, b As String, G As String

    If color = vbGreen Then
        GetRGB = "rgb(0, 255, 0)"
    ElseIf color = vbRed Then
        GetRGB = "rgb(255, 0, 0)"
    Else
        hexCode = Hex(color)
        If Len(hexCode) <> 0 Then
            'Note the order excel uses for hex is BGR.
            r = Val("&H" & Mid(hexCode, 5, 2))
            b = Val("&H" & Mid(hexCode, 1, 2))
            G = Val("&H" & Mid(hexCode, 3, 2))
            GetRGB = "rgb(" & r & ", " & G & ", " & b & ")"
        End If
    End If
End Function

Public Function ToString(Item As Variant) As String
    Dim sb As Object: Set sb = CreateObject("System.Text.StringBuilder")
    Dim Style1 As style
    If IsNumeric(Item) Then
        Set Style1 = ThisWorkbook.Styles(Item)
    Else
        Set Style1 = ThisWorkbook.Styles(CStr(Item))
    End If

    With Style1
        If .IncludeFont Then
            appendCSS_Color sb, Style1
            appendCSS_Font sb, Style1
            appendCSS_TextDecoration sb, Style1
        End If
        appendCSS_WordWrap sb, Style1
        On Error Resume Next
        Call IsNull(.IndentLevel)
        If Err.Number = 0 Then appendCSS_TextIndent sb, Style1
        On Error GoTo 0

        If .IncludeBorder Then appendCSS_Border sb, Style1

        sb.Append_3 "Background:" & GetRGB(Style1.Interior.color) & ";"
    End With
    ToString = sb.ToString
End Function

