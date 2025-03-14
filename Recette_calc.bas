Attribute VB_Name = "Recette_calc"
Public Function Recette_Calcul(ByRef sh As Worksheet, Target As Range) As Double
    If Target.Row < 3 Or Target.Column <> 6 Then Exit Function
    '    Call init_rep
    off_stock = RGB(192, 0, 64)
    Workbooks("Calcul prix et recettes.xlsb").Activate
    Set c1 = Sheets("Produits")
    Set c3 = Sheets("recettes en atelier")
    Set c4 = Sheets("Catégories")
    Set plage = c1.Range("A3:I2000")
    tt = yy
    Workbooks("Calcul prix et recettes.xlsb").Activate
    Cells(Target.Row, Target.Column).Select
    With Target.Validation    ' ATTENTION validation= méthode Excel elle attend des arguments de feuille !!!
        .Delete
        nbrowmax = Workbooks("Calcul prix et recettes.xlsb").Sheets(5).Range("A65000").End(xlUp).Row
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:="=Produits!$H$2:$H" & nbrowmax  '  LISTE PRODUITS
        .IgnoreBlank = True
        .InputMessage = "Stock: " & Str(Stock)
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
    Application.EnableEvents = True
    cle_rech = Cells(Target.Row, Target.Column)
    Nbre = Application.CountIf(plage, cle_rech)
    If Nbre > 0 Then
        nbrow = Target.Row
        pos = nbrow
        nbCol = Target.Column
        nbrowmax = Workbooks("Stock.xlsm").Sheets(1).Range("A65000").End(xlUp).Row
    '        Stock = affich_info_stock()
        Set plage_rech = Workbooks("Stock.xlsm").Sheets(1).Range("A1" & ":A" & nbrowmax)
    '        Set plage_rech = c1.Range("A" & Trim(ActiveCell.row - 1) & ":A2000")
        Set c = plage_rech.Find(cle_rech, , , xlPart)
        If Not c Is Nothing Then
            With c
                fnd1 = c.Row
                Lig = c.Row
                ligne = c.Row
                pos = c.Row
    '        Set c = plage_rech.FindNext(c)
    '        fnd2 = c.row
                c3.Cells(Target.Row, 12) = c1.Cells(ligne, 16).Value
                c3.Cells(Target.Row, 13) = (c3.Cells(Target.Row, 12) * c3.Cells(Target.Row, 8))
                c1.Cells(Target.Row, 8) = (c1.Cells(Target.Row, 2) & c1.Cells(Target.Row, 4) & c1.Cells(Target.Row, 5) & c1.Cells(Target.Row, 6))
                c3.Cells(Target.Row, 9) = (c3.Cells(Target.Row, 12) * c3.Cells(Target.Row, 8))

    '        c1.Cells(ligne, 13).Value = Me("Textbox10").Value  '  densite
    '        c1.Cells(ligne, 14).Value = c1.Cells(ligne, 4).Value * c1.Cells(ligne, 13).Value '  conversion en g
    '        c1.Cells(ligne, 15).Value = c1.Cells(ligne, 10).Value / c1.Cells(ligne, 4).Value '  prix au ml
    '        c1.Cells(ligne, 16).Value = c1.Cells(ligne, 10).Value / c1.Cells(ligne, 14).Value '  Prix au g
    '        c1.Cells(ligne, 16).Value = c1.Cells(ligne, 10).Value / c1.Cells(ligne, 14).Value '  Prix au g

            End With
    '            UserForm3.Show
        End If
    Else
        MsgBox "Cette denomination de produit n'existe pas"
    End If
End Function
