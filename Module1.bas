Attribute VB_Name = "Module1"
Option Explicit

Sub �f�n�S���ħ��ƧǱq�j��p()
Attribute �f�n�S���ħ��ƧǱq�j��p.VB_Description = "�f�n�w�s�q����"
Attribute �f�n�S���ħ��ƧǱq�j��p.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�S���ħ��ƧǱq�j��p ����
' �f�n�w�s�q����
'
' �ֳt��: Ctrl+q
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
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
    
End Sub
Sub �f�n�w�s���W()
Attribute �f�n�w�s���W.VB_Description = "�f�n�w�s���W�m��"
Attribute �f�n�w�s���W.VB_ProcData.VB_Invoke_Func = "y\n14"
'
' �f�n�w�s���W ����
' �f�n�w�s���W�m��
'
' �ֳt��: Ctrl+y
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
    
End Sub
