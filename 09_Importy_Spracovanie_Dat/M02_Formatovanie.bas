Attribute VB_Name = "M02_Formatovanie"
Option Explicit

Sub formatovanie_dat_znecistenie()
Attribute formatovanie_dat_znecistenie.VB_Description = "Format cisel, bez desatin, podmienene formatovanie udajove pruhy GDP,\npercent preè, zalomit text, nahradit _ za nic -> odstranit\nsklo/glass podmienen formatovanie"
Attribute formatovanie_dat_znecistenie.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' formatovanie_dat_znecistenie Makro
' Format cisel, bez desatin, podmienene formatovanie udajove pruhy GDP, percent preè, zalomit text, nahradit _ za nic -> odstranit sklo/glass podmienen formatovanie
'
' Klávesová skratka: Ctrl+Shift+F
'
    Range("country_level_data_0[gdp]").Select
    Selection.Style = "Comma [0]"
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 8061142
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
    Range( _
        "country_level_data_0[[#Headers],[composition_food_organic_waste_percent]]"). _
        Select
    Cells.Replace What:="percent", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="_", Replacement:=" ", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
    Range("country_level_data_0[#Headers]").Select
    Range("country_level_data_0[[#Headers],[composition food organic waste ]]"). _
        Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.RowHeight = 35
    Range("country_level_data_0[#Headers]").Select
    Range("country_level_data_0[[#Headers],[country name]]").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D6").Select
    Columns("D:D").EntireColumn.AutoFit
    Columns("D:D").ColumnWidth = 17.92
    Columns("E:E").ColumnWidth = 15.83
    Columns("F:F").ColumnWidth = 15.5
    Columns("G:G").ColumnWidth = 16
    Columns("G:G").ColumnWidth = 11.33
    Columns("H:H").ColumnWidth = 21.75
    Range("country_level_data_0[[composition food organic waste ]]").Select
    ActiveWindow.SmallScroll Down:=-357
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Range("D7").Select
    Cells.Replace What:="NA", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False, _
        FormulaVersion:=xlReplaceFormula2
    Range("country_level_data_0[[#Headers],[region id]]").Select
End Sub
