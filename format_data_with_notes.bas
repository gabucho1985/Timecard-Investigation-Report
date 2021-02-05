Attribute VB_Name = "format_data_with_notes"
'Column manipulation
Sub StackDataToOneColumn() 'code used to put 2 columns into 1

Dim Rng1 As Range, Rng2 As Range, rng As Range

Dim RowIndex As Integer

Set Rng1 = Application.Selection

Set Rng1 = Application.InputBox("Select Range:", "StackDataToOneColumn", Rng1.Address, Type:=8)

Set Rng2 = Application.InputBox("Destination Column:", "StackDataToOneColumn", Type:=8)

RowIndex = 0

Application.ScreenUpdating = False


For Each rng In Rng1.Rows

    rng.Copy
     
    Rng2.Offset(RowIndex, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True

    RowIndex = RowIndex + rng.Columns.Count

Next
Range("G1:G5000").Delete
Application.CutCopyMode = False

Application.ScreenUpdating = True


End Sub

Sub Sort_Desc()

Range("F1") = "Index"
   Columns("F").Sort key1:=Range("F2"), _
      order1:=xlDescending, Header:=xlYes
End Sub

'split columns into date and time
Sub SplitTime()

Dim ws As Worksheet

Dim lastRow As Long
Dim Count As Long
Dim test As Double

Set ws = ActiveSheet


'Find last data point
With ws
    '.Columns(7).Insert
    lastRow = .Cells(.Rows.Count, "F").End(xlUp).Row

    For Count = 2 To lastRow
        'split date
        test = .Cells(Count, 6).Value2
        .Cells(Count, 6).Value2 = Int(test)
        .Cells(Count, 6).NumberFormat = "m/d/yyyy"
        .Cells(Count, 6).Offset(0, 1).Value2 = test - Int(test)
        .Cells(Count, 6).Offset(0, 1).NumberFormat = "hh:mm AM/PM"

    Next Count
End With
End Sub

'# code converts time to h:mm
Sub Convert_Time()
Columns("H:I").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("G" & Rows.Count).End(xlUp).Row
'
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""h:mm"")"
    Range("H1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("H1").AutoFill Destination:=Range("H1:H" & lastRow)
    Range("H1:H" & Range("G" & Rows.Count).End(xlUp).Row).Rows.Copy

    Range("G1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("H:I").Delete Shift:=xlShiftToLeft
End Sub
Sub Calle()
Call StackDataToOneColumn
Call Sort_Desc
Call SplitTime
Call Convert_Time
End Sub
'Fill in the blanks. Populate missing time
Sub Inut2() 'this works
For Each rw In UsedRange.Rows
  If rw.Columns("C") = "" Then
     rw.Columns("C") = rw.Columns("F")
  
    'rw.Columns("C") = rw.Columns("F") & " " & rw.Columns("G")
  End If
Next rw
End Sub











