Attribute VB_Name = "Module1"
Option Explicit

Sub 口罩特約藥局排序從大到小()
Attribute 口罩特約藥局排序從大到小.VB_Description = "口罩庫存量遞減"
Attribute 口罩特約藥局排序從大到小.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩特約藥局排序從大到小 巨集
' 口罩庫存量遞減
'
' 快速鍵: Ctrl+q
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
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
    
End Sub
Sub 口罩庫存遞增()
Attribute 口罩庫存遞增.VB_Description = "口罩庫存遞增練習"
Attribute 口罩庫存遞增.VB_ProcData.VB_Invoke_Func = "y\n14"
'
' 口罩庫存遞增 巨集
' 口罩庫存遞增練習
'
' 快速鍵: Ctrl+y
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
    
End Sub
