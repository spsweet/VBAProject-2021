Attribute VB_Name = "ModuleAscsort"
Option Explicit

Sub �̭]�ƶq���W()
Attribute �̭]�ƶq���W.VB_Description = "�Ĥ@���̭]���W"
Attribute �̭]�ƶq���W.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �̭]�ƶq���W ����
' �Ĥ@���̭]���W
'
' �ֳt��: Ctrl+q
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
