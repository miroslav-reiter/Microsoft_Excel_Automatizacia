REM Podmienene Formatovanie a VBA

' =IF(F9>0;REPT("●";INT(F9*100));"")
' =IF(F10<0;REPT("●";INT(-F10*100));"")
' =IF(F11>0;REPT("●";INT(F11*100));"")
' =IF(F12>0;REPT("●";INT(F12*100));"")

' Farba Vyplne/Bunky
' Rozsah Stlpec, Oblast, Tabulka
Function sum_farba(Farba As Range, Rozsah As Range)
    Dim X As Double
    Dim Y As Double
    Dim i As Variant 'Object
    
    Y = Farba.Interior.ColorIndex
    
    For Each i In Rozsah
        If i.Interior.ColorIndex = Y Then
            X = WorksheetFunction.Sum(i, X)
        End If
    Next i
    
    sum_farba = X
End Function

' Farba Vyplne/Bunky
' Rozsah Stlpec, Oblast, Tabulka
Function min_farba(Farba As Range, Rozsah As Range)
    Dim X As Double
    Dim Y As Double
    Dim i As Variant 'Object
    
    Y = Farba.Interior.ColorIndex
    
    For Each i In Rozsah
        If i.Interior.ColorIndex = Y Then
            X = WorksheetFunction.Min(i, X)
        End If
    Next i
    
    min_farba = X
End Function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' ColorIndex property (Excel Graph)
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
    ' 1 - cierna, 2 - biela, 3 - cervena, 4 - Zelena,
    ' 5 - Modra, 6 - zlta, 7 - magenta, 8 - cyan, 9 - bordova
    Cells.Interior.ColorIndex = xlColorIndexNone
    Target.EntireColumn.Interior.ColorIndex = 6
    Target.EntireRow.Interior.ColorIndex = 6
    Target.Interior.ColorIndex = xlColorIndexNone
End Sub

