Attribute VB_Name = "ModuleAscsort"
Option Explicit

Sub 疫苗數量遞增()
Attribute 疫苗數量遞增.VB_Description = "第一劑疫苗遞增"
Attribute 疫苗數量遞增.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 疫苗數量遞增 巨集
' 第一劑疫苗遞增
'
' 快速鍵: Ctrl+q
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("G2").Select
End Sub
