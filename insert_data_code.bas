Attribute VB_Name = "insert_data_code"


'first check for missing data from first row and last row. Always check for that
'and go through data so there is no nissing data for ex: from sat to wed. If that's the case
'then code needs to be modified.
'Also check for Sundays, get them out then put them back in at the end of everything'first add missing data along PTO's
'code for investors_data to fill missing data this is managers data used to pay employee
'convert column A into date format
Sub Convertdate()
Columns("B:B").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""m/d/yyyy"")"
    Range("B1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("B1").AutoFill Destination:=Range("B1:B" & lastRow)
    Range("B1:B" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
    'Range("B1").Copy
'    Range("A" & Rows.Count).End(xlUp).Row.Select
'    Range("A" & Rows.Count).End(xlUp).Row.Copy
    
'    Range("B1:B").Select
'    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("B:B").Delete Shift:=xlShiftToLeft
End Sub
'code creates weekdays on col E
Sub Todo()
    'Range("E1").Select
    'ActiveCell.FormulaR1C1 = "Weekday"
    Range("E1") = "Weekday"
    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-4],""dddd"")"
    Range("E2").Select
    Range("E2").AutoFill Destination:=Range("E2:E" & lastRow)
    Range("E2:E" & Range("E" & Rows.Count).End(xlUp).Row).Rows.Copy
    Range("E2").Select
    'Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub MacroWeekDa()
   Call Convertdate
   Call Todo
End Sub

'from friday to monday input saturday
Sub InSat()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Friday" And Cells(j - 1, 5).Value = "Monday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Saturday"
        Cells(j, 2).Value = "0"
        'Cells(j + 1, 3).Value = "0"
  End If
  
Next
Next
End Sub
'from fr to t input saturday
Sub InMon()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 5).Value = "Friday" And Cells(j - 1, 5).Value = "Tuesday" _
     Then
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Saturday"
        Cells(j, 2).Value = "0"

        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 2
        Cells(j, 5).Value = "Monday"
        Cells(j, 2).Value = "0"
  End If
  
Next
Next
End Sub
'sat to t input monday
Sub InMonday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Saturday" And Cells(j - 1, 5).Value = "Tuesday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 2
        Cells(j, 5).Value = "Monday"
        Cells(j, 2).Value = "0"
  End If
  
Next
Next
End Sub
'Monday to Wednesday input Tuesday
Sub InTuesday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Monday" And Cells(j - 1, 5).Value = "Wednesday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Tuesday"
        Cells(j, 2).Value = "0"
  End If
  
Next
Next
End Sub
'TUESDAY to Thursday input Wednesday
Sub InWed()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Tuesday" And Cells(j - 1, 5).Value = "Thursday" _
     Then
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Wednesday"
        Cells(j, 2).Value = "0"
  End If
  
Next
Next
End Sub
'wednesday to F input Thursday
Sub InThursday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Wednesday" And Cells(j - 1, 5).Value = "Friday" _
     Then
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Thursday"
        Cells(j, 2).Value = "0"
  End If
   
Next
Next
End Sub
'THursday to Sat input Friday
Sub InFriday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  If Cells(j, 5).Value = "Thursday" And Cells(j - 1, 5).Value = "Saturday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = CDate(Cells(j + 1, 1).Value) + 1
        Cells(j, 5).Value = "Friday"
        Cells(j, 2).Value = "0"
  End If
  
Next
Next
End Sub
 
Sub Macro11()
Columns("B:E").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
 
'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""m/d/yyyy"")"
    Range("B1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("B1").AutoFill Destination:=Range("B1:B" & lastRow)
    Range("B1:B" & Range("B" & Rows.Count).End(xlUp).Row).Rows.Copy
    'Range("B1").Copy
'    Range("A" & Rows.Count).End(xlUp).Row.Select
'    Range("A" & Rows.Count).End(xlUp).Row.Copy
    
'    Range("B1:B").Select
'    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("B:E").Delete Shift:=xlShiftToLeft
End Sub
 

Sub HG()
Call InSat
Call InMon
Call InMonday
Call InTuesday
Call InWed
Call InThursday
Call InFriday
'Call Macro11
End Sub








