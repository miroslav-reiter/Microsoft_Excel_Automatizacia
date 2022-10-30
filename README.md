# 📋 Microsoft Excel Automatizácia pomocou makier a jazyka VBA
Zdrojové kódy a skripty v jazyku VBA pre automatizáciu úloh v Microsoft Excel

## Trieda Application a jej Procedúry 
### 📂 Otvorenie Súborov (GetOpenFilename)
```vb
Option Explicit

Sub otvorit_subor()

Dim subor_na_otvorenie As Variant
subor_na_otvorenie = Application.GetOpenFilename("Text Files (*.txt), *.txt")

If subor_na_otvorenie <> False Then
 MsgBox "Je otvoreny subor: " & subor_na_otvorenie
End If

End Sub
```
### ⌚ Spustenie procedúry v danom čase (Wait)
```vb
Sub spusti_v_case()

Dim dtCas As Date: dtCas = "22:32:00"
Dim cakanie As Boolean

cakanie = Application.Wait(Time:=dtCas)
MsgBox "Nastal cas... " & cakanie

End Sub

```
### 🍒 Množinové hromadné operácie nad rozsahmi (Range) a to zjednotenie (Union)
```vba
Sub vypocitaj_hromadne()

Application.Worksheets("hárok1").Activate
Dim velkyRozsah As Variant
Set velkyRozsah = Application.Union(Range("B1:C100000"), Range("F5:J100000"))
velkyRozsah.Formula = "=randbetween(1,6)"

End Sub

```
### 🍒 Množinové hromadné operácie nad rozsahmi (Range) a to prienik (Intersect)
```vba

Sub over_prienik_rozsahov()

Application.Worksheets("hárok1").Activate
Dim velkyRozsah As Variant
Set velkyRozsah = Application.Intersect(Range("B1:F100000"), Range("B5:J100000"))


If velkyRozsah Is Nothing Then
    MsgBox "Rozsahy nemaju prienik"
Else
    MsgBox "Rozsahy maju prienik"
    velkyRozsah.Select
End If

End Sub
```

### 💀 Konvertovanie Štýlu funkcie (A1 <--> R1C1, RELATIVE <--> ABSOLUTE) (ConvertFormula)
```vba
Sub konvertuj_funkcie()

Dim vstupna_funkcia As Variant
vstupna_funkcia = "=sum(R2C1:R5C2)"
MsgBox Application.ConvertFormula(Formula:=vstupna_funkcia, _
fromReferenceStyle:=xlR1C1, toReferenceStyle:=xlA1)

End Sub
```

### 🖨️ Tlač Dokumentov a staré makrá (ExecuteExcel4Macro)
```vba
Sub tlac_dokument()

Dim pocetStran As Long
pocetStran = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
MsgBox "Celkovy pocet stran na tlac: " & pocetStran

With ActiveSheet.PageSetup
    .CenterHeader = "Testovaci text"
    ActiveSheet.PrintOut From:=1, To:=1, copies:=1, preview:=True
    .CenterHeader = "Projekt ABC"
    ActiveSheet.PrintOut From:=2, To:=pocetStran, copies:=1, preview:=True
End With
End Sub
```
### 🟨 Zvýraznenie celého riadku a stĺpca podľa aktuálne vybranej bunky (Do )
![2022-10-24 20_52_07-SelectionChange xlsm - Excel](https://user-images.githubusercontent.com/24510943/197603184-d853ae6d-6c29-4cb2-be0e-9357537ac5b6.png)

```vba
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
```

![colors](https://user-images.githubusercontent.com/24510943/197624888-29fd0036-5363-4d99-8c71-341f4d90145f.png)


```vba
Sub vytlacFarby()
    Dim riadok As Integer: riadok = 2
    Dim stlpec As Integer: stlpec = 2
    Dim i As Integer

    For i = 1 To 56
        Cells(riadok, stlpec).Interior.ColorIndex = i
        Cells(riadok, stlpec).Value = i

        If i > 1 And i Mod 14 = 0 Then
            stlpec = stlpec + 1
            riadok = 2
        Else
            riadok = riadok + 1
        End If
    Next i

    'Range("B2:E15").Interior.ColorIndex = -4142
    Range("B2:E15").Borders.ColorIndex = -4142
    Range("B2:E15").Font.ColorIndex = -4105

End Sub
```

![color-vba](https://user-images.githubusercontent.com/24510943/197604877-5859216e-352d-494f-af55-dc7c29e747c8.gif)

## 🧱 Príklad na Const (Konštantu)
```vba
Const odpovedOtazkaZivotaSmrti As Integer = 42
Const bulharskaKonstanta = 8
Const pocetBodov = 4
Const PI As Double = 3.14
Const E = 2.78
Const DPH = 1.2
Public Const SPRAVA As String = "Zapis sa do prezencky"
Const konstanta1 = "Ahoj", konstanta2 As String = "Hello"
```
## ✨ Príklady na Enum (Enumerácia)

```vba
Enum ZnackyAut
    Porsche = 100
    Audi
    Skoda
    Opel
    Seat
End Enum
```

```vba
Enum OddeleniaRozpocty
    IT = 10000
    HR = 9000
    SALES = 8000
    MARKETING = 20000
    OPERATION = 5000

End Enum
```

```vba
Enum OddeleniaRozpocty
    IT = 10000
    HR = 9000
    SALES = 8000
    MARKETING = 20000
    OPERATION = 5000

End Enum
```

```vba
Public Enum InterfaceColors 

icDeepSkyBlue = &HFFBF00& 
icSpringGreen = &H7FFF00& 
icForestGreen = &H228B22& 
icGoldenrod = &H20A5DA& 

End Enum
```
