Attribute VB_Name = "Module1"
Option Explicit

Sub ��ƲM�z��k()
Dim totalrow, totalcolumn As Long
totalrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row '���oA��Ҧ��C��

'���J�ӷ��@�̮ɶ����p�����
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'��J�Ĥ@�椽��
Range("D2").Select
ActiveCell.FormulaR1C1 = "=SEARCH(""by"",RC[-1])"
Range("E2").Select
ActiveCell.FormulaR1C1 = "=SEARCH(""/"",RC[-2])"
Range("F2").Select
ActiveCell.FormulaR1C1 = _
"=IFERROR(LEFT(RC[-3],RC[-2]-2),LEFT(RC[-3],RC[-1]-3))"
Range("G2").Select
ActiveCell.FormulaR1C1 = "=IFERROR(MID(RC[-4],RC[-3]+3,RC[-2]-RC[-3]-4),"""")"
Range("H2").Select
ActiveCell.FormulaR1C1 = "=RIGHT(RC[-5],LEN(RC[-5])-RC[-3]-1)"

'�۰ʶ�J���
Range("D2:H2").Select
Selection.AutoFill Destination:=Range(Cells(2, "D"), Cells(totalrow, "H"))

'���J�ƻs��ƪ����
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("E:E").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'�ƻs��ƨöK�W
Columns("I:K").Select
Selection.Copy
Columns("C:E").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Range("C1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "�ӷ�"
Range("D1").Select
ActiveCell.FormulaR1C1 = "�@��"
Range("E1").Select
ActiveCell.FormulaR1C1 = "�פJ�ɶ�"
Columns("F:K").Select
Selection.Delete Shift:=xlToLeft

'�w�qtotalcolum
totalcolumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1 '���o�Ĥ@�C�Ҧ���ơA�]���᭱�n�s�W�ҥH�n�[�@

'���s�W�n�\�񤺤��ƪ����
ActiveSheet.Columns("G:G").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow

'�[�J����
ActiveSheet.Cells(2, totalcolumn + 1).Select
ActiveCell.FormulaR1C1 = "=TEXTJOIN("" "",1,RC8:RC[-1])"
ActiveSheet.Cells(2, totalcolumn + 1).Select

'�۰ʶ�J����
Selection.AutoFill Destination:=Range(Cells(2, totalcolumn + 1), Cells(totalrow, totalcolumn + 1))
'��w�n�[�J��������m
ActiveSheet.Cells(2, totalcolumn + 2).Select
ActiveCell.FormulaR1C1 = "=SUBSTITUTE(RC[-1],""no_link"","" "")"
'�۰ʶ�J����
Selection.AutoFill Destination:=Range(Cells(2, totalcolumn + 2), Cells(totalrow, totalcolumn + 2))
Range(Cells(2, totalcolumn + 2), Cells(totalrow, totalcolumn + 2)).Select
Selection.Copy
Range(Cells(2, "G"), Cells(totalrow, "G")).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Range("H2").Select
Application.CutCopyMode = False
Selection.Copy
Columns("G:G").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("G1").Select
ActiveCell.FormulaR1C1 = "����"
Columns("H:H").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft
Columns("H:I").Select
Selection.Delete Shift:=xlToLeft

'�s�����W��
Range("A1").Select
ActiveCell.FormulaR1C1 = "���D"
Range("B1").Select
ActiveCell.FormulaR1C1 = "������}"
Range("F1").Select
ActiveCell.FormulaR1C1 = "�ӷ����}"

End Sub
