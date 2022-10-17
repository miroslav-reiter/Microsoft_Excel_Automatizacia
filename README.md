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





```
