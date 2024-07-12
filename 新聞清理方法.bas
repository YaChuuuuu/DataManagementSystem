Attribute VB_Name = "Module1"
Option Explicit

Sub 資料清理方法()
Dim totalrow, totalcolumn As Long
totalrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row '取得A欄所有列數

'插入來源作者時間的計算欄位
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'輸入第一行公式
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

'自動填入資料
Range("D2:H2").Select
Selection.AutoFill Destination:=Range(Cells(2, "D"), Cells(totalrow, "H"))

'插入複製資料的欄位
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("E:E").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'複製資料並貼上
Columns("I:K").Select
Selection.Copy
Columns("C:E").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Range("C1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "來源"
Range("D1").Select
ActiveCell.FormulaR1C1 = "作者"
Range("E1").Select
ActiveCell.FormulaR1C1 = "匯入時間"
Columns("F:K").Select
Selection.Delete Shift:=xlToLeft

'定義totalcolum
totalcolumn = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column + 1 '取得第一列所有欄數，因為後面要新增所以要加一

'先新增要擺放內文資料的欄位
ActiveSheet.Columns("G:G").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow

'加入公式
ActiveSheet.Cells(2, totalcolumn + 1).Select
ActiveCell.FormulaR1C1 = "=TEXTJOIN("" "",1,RC8:RC[-1])"
ActiveSheet.Cells(2, totalcolumn + 1).Select

'自動填入公式
Selection.AutoFill Destination:=Range(Cells(2, totalcolumn + 1), Cells(totalrow, totalcolumn + 1))
'選定要加入公式的位置
ActiveSheet.Cells(2, totalcolumn + 2).Select
ActiveCell.FormulaR1C1 = "=SUBSTITUTE(RC[-1],""no_link"","" "")"
'自動填入公式
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
ActiveCell.FormulaR1C1 = "內文"
Columns("H:H").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft
Columns("H:I").Select
Selection.Delete Shift:=xlToLeft

'編輯欄位名稱
Range("A1").Select
ActiveCell.FormulaR1C1 = "標題"
Range("B1").Select
ActiveCell.FormulaR1C1 = "內文網址"
Range("F1").Select
ActiveCell.FormulaR1C1 = "來源網址"

End Sub
