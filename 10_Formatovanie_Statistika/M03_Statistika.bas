Attribute VB_Name = "M03_Statistika"
Option Explicit

Sub vypis_statistiku_tab_znecistenie()


Dim tabZnecistenie As ListObject
Set tabZnecistenie = ActiveSheet.ListObjects("country_level_data_0")

MsgBox "Tabulka Znecistenie ma celkovy pocet riadkov: " & tabZnecistenie.Range.Rows.Count
MsgBox "Tabulka Znecistenie ma celkovy pocet riadkov v hlavicke: " & tabZnecistenie.HeaderRowRange.Rows.Count
MsgBox "Tabulka Znecistenie ma celkovy pocet riadkov v hlavicke: " & tabZnecistenie.DataBodyRange.Rows.Count

MsgBox "Tabulka Znecistenie ma celkovy pocet stlpcov: " & tabZnecistenie.Range.Columns.Count


Set tabZnecistenie = Nothing


End Sub
