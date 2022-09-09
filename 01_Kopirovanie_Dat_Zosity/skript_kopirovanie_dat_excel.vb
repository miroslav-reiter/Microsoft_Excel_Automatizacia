```vba
Option Explicit

Private Sub cmdbtn_odoslat_vop_Click()

Dim polozkaNazov As String
Dim polozkaCena As Single
Dim dataVOP As Workbook

Worksheets("VOP").Select
polozkaNazov = Range("B1")

Worksheets("VOP").Select
polozkaCena = Range("B2")

Set dataVOP = Workbooks.Open("C:\VOP\Vystupy.xlsx")
Worksheets("Importy").Select
Worksheets("Importy").Range("A1").Select

Dim pocetRiadkov As Integer
pocetRiadkov = Worksheets("Importy").Range("A1").CurrentRegion.Rows.Count

With Worksheets("Importy").Range("A1")
.Offset(pocetRiadkov, 0) = polozkaNazov
.Offset(pocetRiadkov, 1) = polozkaCena
End With

dataVOP.Save

End Sub
```
