'Option Explicit

Public Sub spracuj_vo()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim precinok_cesta As String
    Dim subor_cesta As String
    Dim subor_nazov As String

    Dim riadok_posledny As Long, stlpec_posledny As Long

    precinok_cesta = "C:\VO\"
    subor_cesta = precinok_cesta & "*.xls*"
    subor_nazov = Dir(subor_cesta)

    Debug.Print subor_nazov

    ' Len(subor_nazov) > 0
    Do While subor_nazov <> ""
        If subor_nazov = "Vystup.xlsm" Then
            GoTo Koniec
        End If

        Workbooks.Open (precinok_cesta & subor_nazov)
        
        riadok_posledny = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        stlpec_posledny = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

        'Range("A2:F2").Copy
        Range(Cells(2, 1), Cells(riadok_posledny, stlpec_posledny)).Copy
        ActiveWorkbook.Close

        riadok_posledny = Hárok1.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        stlpec_posledny = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        ActiveSheet.Paste Destination:=Worksheets("hárok1").Range(Cells(riadok_posledny, 1), Cells(riadok_posledny, 1))

        subor_nazov = Dir()

    Loop
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

Koniec:
    Worksheets("Hárok1").Range("A:F").Columns.AutoFit
    
    MsgBox "Si nakopiroval data z inych Excel suborov."
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub

Rem -----------------------------------
Rem Automaticke kopirovanie po otvoreni
Private Sub Workbook_Open()
    Call spracuj_vo
End Sub

Rem -----------------------------------