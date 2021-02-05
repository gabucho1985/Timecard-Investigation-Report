Attribute VB_Name = "Employee_inv_code"


'first add missing data along PTO's. Check for dates that have more than 4 entries(pivot tables/python)
'codes below clean time and dates after inserting missing data manually. This is done before converting time to integers
'make sure to click on column C
'#1 before running code check latest week to fill missing weeks/ do the same with oldest date at the bottom end
'this code below converts data to good date format "m/d/yyyy"
Sub Convertdate()
Columns("C:C").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("B" & Rows.Count).End(xlUp).Row
'
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""m/d/yyyy"")"
    Range("C1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("C1").AutoFill Destination:=Range("C1:C" & lastRow)
    Range("C1:C" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
    'Range("B1").Copy
'    Range("A" & Rows.Count).End(xlUp).Row.Select
'    Range("A" & Rows.Count).End(xlUp).Row.Copy
    
'    Range("B1:B").Select
'    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("C:C").Delete Shift:=xlShiftToLeft
End Sub



''make sure to click on column F
'Sub Macro12()
'Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'Dim lastRow As Long
'    lastRow = Range("F" & Rows.Count).End(xlUp).Row
''
'    Range("G1").Select
'    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""m/d/yyyy"")"
'    Range("G1").Select
'    'Selection.AutoFill Destination:=Range("B1:B3")
'    Range("G1").AutoFill Destination:=Range("G1:G" & lastRow)
'    Range("G1:G" & Range("H" & Rows.Count).End(xlUp).Row).Rows.Copy
'    'Range("B1").Copy
''    Range("A" & Rows.Count).End(xlUp).Row.Select
''    Range("A" & Rows.Count).End(xlUp).Row.Copy
'
''    Range("B1:B").Select
''    Selection.Copy
'    Range("F1").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'    Columns("G:G").Delete Shift:=xlShiftToLeft
'End Sub
'#2 code converts time to h:mm
Sub Convert_Time()
Columns("D:E").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
'
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""h:mm"")"
    Range("D1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("D1").AutoFill Destination:=Range("D1:D" & lastRow)
    Range("D1:D" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
    'Range("B1").Copy
'    Range("A" & Rows.Count).End(xlUp).Row.Select
'    Range("A" & Rows.Count).End(xlUp).Row.Copy
    
'    Range("B1:B").Select
'    Selection.Copy
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:E").Delete Shift:=xlShiftToLeft
End Sub
Sub MacroWeekDays()
    Range("D1") = "Weekday"
    Dim lastRow As Long
    lastRow = Range("B" & Rows.Count).End(xlUp).Row
    'ActiveCell.FormulaR1C1 = "Weekday"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""dddd"")"
    Range("D2").Select
    Range("D2").AutoFill Destination:=Range("D2:D" & lastRow)
    Range("D2:D" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
'    Range("D2:D615").Select
'    Selection.Copy
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub Fcall()
Call Convertdate
Call Convert_Time
Call MacroWeekDays
End Sub

'step#3
'from friday to monday input saturday
Sub InSat()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Friday" And Cells(j - 1, 4).Value = "Monday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Saturday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Saturday"
        Cells(j, 3).Value = "0:00"
        Cells(j + 1, 3).Value = "0:00"
  End If
  
Next
Next
End Sub
 
'from fr to tuesday input saturday/monday
Sub InMon()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Friday" And Cells(j - 1, 4).Value = "Tuesday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Saturday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Saturday"
        Cells(j, 3).Value = "0"
        Cells(j - 1, 3).Value = "0"
  End If
  
  If Cells(j - 1, 2).Value = "2/17/2020" And Cells(j, 2).Value = "2/17/2020" _
     Or Cells(j - 1, 2).Value = "1/20/2020" And Cells(j, 2).Value = "1/20/2020" _
     Or Cells(j - 1, 2).Value = "11/11/2019" And Cells(j, 2).Value = "11/11/2019" _
     Or Cells(j - 1, 2).Value = "10/14/2019" And Cells(j, 2).Value = "10/14/2019" _
     Or Cells(j - 1, 2).Value = "9/2/2019" And Cells(j, 2).Value = "9/2/2019" _
     Then
     'Range(Cells(1,1),Cells(5,5))
        Cells(j, 3).Value = "8:00"
        Cells(j - 1, 3).Value = "16:00"
  
  
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
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Saturday" And Cells(j - 1, 4).Value = "Tuesday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 2
        Cells(j, 4).Value = "Monday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Monday"
        Cells(j, 3).Value = "0"
        Cells(j + 1, 3).Value = "0"
  End If
  
  If Cells(j - 1, 2).Value = "2/17/2020" And Cells(j, 2).Value = "2/17/2020" _
     Or Cells(j - 1, 2).Value = "1/20/2020" And Cells(j, 2).Value = "1/20/2020" _
     Or Cells(j - 1, 2).Value = "11/11/2019" And Cells(j, 2).Value = "11/11/2019" _
     Or Cells(j - 1, 2).Value = "10/14/2019" And Cells(j, 2).Value = "10/14/2019" _
     Or Cells(j - 1, 2).Value = "9/2/2019" And Cells(j, 2).Value = "9/2/2019" _
     Then
     'Range(Cells(1,1),Cells(5,5))
        Cells(j, 3).Value = "8:00"
        Cells(j - 1, 3).Value = "16:00"
  
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
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Monday" And Cells(j - 1, 4).Value = "Wednesday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Tuesday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Tuesday"
        Cells(j, 3).Value = "0"
        Cells(j + 1, 3).Value = "0"
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
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Tuesday" And Cells(j - 1, 4).Value = "Thursday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Wednesday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Wednesday"
        Cells(j, 3).Value = "0"
        Cells(j + 1, 3).Value = "0"
  End If
  
   If Cells(j - 1, 2).Value = "1/1/2020" And Cells(j, 2).Value = "1/1/2020" _
     Or Cells(j - 1, 2).Value = "12/25/2019" And Cells(j, 2).Value = "12/25/2019" _
     Then
     'Range(Cells(1,1),Cells(5,5))
        Cells(j, 3).Value = "8:00"
        Cells(j - 1, 3).Value = "16:00"
  
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
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Wednesday" And Cells(j - 1, 4).Value = "Friday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Thursday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Thursday"
        Cells(j, 3).Value = "0"
        Cells(j + 1, 3).Value = "0"
  End If
  
  If Cells(j - 1, 2).Value = "11/28/2019" And Cells(j, 2).Value = "11/28/2019" _
     Then
     
        Cells(j, 3).Value = "8:00"
        Cells(j - 1, 3).Value = "16:00"
  
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
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If Cells(j, 4).Value = "Thursday" And Cells(j - 1, 4).Value = "Saturday" _
     Then
        
        Cells(j, 1).EntireRow.insert
        Cells(j, 1).Value = "OUT"
        Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value) + 1
        Cells(j, 4).Value = "Friday"
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value)
        Cells(j + 1, 4).Value = "Friday"
        Cells(j, 3).Value = "0"
        Cells(j + 1, 3).Value = "0"
  End If
  
Next
Next
End Sub
 'converts time to "h:mm"
Sub Macro111()
Columns("D:F").insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Dim lastRow As Long
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
'
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""h:mm"")"
    Range("D1").Select
    'Selection.AutoFill Destination:=Range("B1:B3")
    Range("D1").AutoFill Destination:=Range("D1:D" & lastRow)
    Range("D1:D" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
    'Range("B1").Copy
'    Range("A" & Rows.Count).End(xlUp).Row.Select
'    Range("A" & Rows.Count).End(xlUp).Row.Copy
    
'    Range("B1:B").Select
'    Selection.Copy
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:F").Delete Shift:=xlShiftToLeft
End Sub
 
'Range(Cells(1,1),Cells(5,5))
'If Cells(j - 1, 2).Value = "2/17/2020" And Cells(j, 2).Value = "2/17/2020" _
'     Then
'        Cells(j, 3).Value = Format("8:00", "h:mm")
'        Cells(j - 1, 3).Value = Format("16:00", "h:mm")
'
'
'  End If
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
Sub HG()
Call InSat
Call InMon
Call InMonday
Call InTuesday
Call InWed
Call InThursday
Call InFriday
Call Macro111
Call Fcall
End Sub
'#step 4
'0 in,in=meal,mael,out
Sub In_missing()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
  If Cells(j, 2).Value <> Cells(j - 1, 2).Value And Cells(j - 1, 2).Value = Cells(j - 2, 2).Value _
    And Cells(j, 1).Value = "IN" And Cells(j - 1, 1).Value = "INN" Then

    Cells(j, 1).EntireRow.insert
    Cells(j, 1).Value = "MEAL"
    Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value)
     Cells(j, 1).EntireRow.insert
    Cells(j, 1).Value = "MAEL"
    Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value)
    Cells(j, 1).EntireRow.insert
    Cells(j, 1).Value = "OUT"
    Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value)
   End If
   
'   If CDate(Cells(j, 2).Value) = CDate(Cells(j - 1, 2).Value) _
'  And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
'  And CDate(Cells(j - 2, 2).Value) = CDate(Cells(j - 3, 2).Value) Then
'
'    'Cells(j, 1).EntireRow.Insert
'    Cells(j, 1).Value = "INN"
'    'Cells(j, 1).EntireRow.Insert
'    Cells(j - 1, 1).Value = "MEAL"
'    Cells(j - 2, 1).Value = "MAEL"
'    Cells(j - 3, 1).Value = "OUT"
'  End If
   Next
   Next

    End Sub

'1 adds inn, meal, mael, out
Sub Inuti()
LR = Cells(Rows.Count, "A").End(xlUp).Row
MM = Cells(Rows.Count, "B").End(xlUp).Row
For H = MM To 2 Step -1
For j = LR To 2 Step -1
  'If Cells(j, 1).Value = "IN" Then Cells(j + 1, 1).Value = "INN"
  'If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
  If CDate(Cells(j, 2).Value) = CDate(Cells(j - 1, 2).Value) _
  And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
  And CDate(Cells(j - 2, 2).Value) = CDate(Cells(j - 3, 2).Value) Then
 
    'Cells(j, 1).EntireRow.Insert
    Cells(j, 1).Value = "INN"
    'Cells(j, 1).EntireRow.Insert
    Cells(j - 1, 1).Value = "MEAL"
    Cells(j - 2, 1).Value = "MAEL"
    Cells(j - 3, 1).Value = "OUT"
  End If
  'If Cells(j, 1).Value = "IN" Then Cells(j, 1).Value = "INN"
 
Next
Next
End Sub
'2 after getting values convert dates again and run step # 1 again to 'convert IN to INN
Sub In_missing_wo()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
  If CDate(Cells(j, 2).Value) <> CDate(Cells(j - 1, 2).Value) And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
    And CDate(Cells(j - 2, 2).Value) <> CDate(Cells(j - 3, 2).Value) _
    And Cells(j - 1, 1).Value = "IN" And Cells(j - 2, 1).Value = "OUT" Then
    'Cells(j - 2, 1).Value And Cells(j - 2, 1).Value <> Cells(j - 3, 1).Value Then
     
    Cells(j - 1, 1).EntireRow.insert
    Cells(j - 1, 1).Value = "MEAL"
    Cells(j - 1, 2).Value = Cells(j - 2, 2).Value
 
    
    Cells(j - 1, 1).EntireRow.insert
    Cells(j - 1, 1).Value = "MAEL"
    Cells(j - 1, 2).Value = Cells(j - 2, 2).Value
 
 
   End If
   
   If CDate(Cells(j, 2).Value) = CDate(Cells(j - 1, 2).Value) _
  And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
  And CDate(Cells(j - 2, 2).Value) = CDate(Cells(j - 3, 2).Value) Then
 
    'Cells(j, 1).EntireRow.Insert
    Cells(j, 1).Value = "INN"
    'Cells(j, 1).EntireRow.Insert
    Cells(j - 1, 1).Value = "MEAL"
    Cells(j - 2, 1).Value = "MAEL"
    Cells(j - 3, 1).Value = "OUT"
  End If
   
   
   Exit For
Next
Next
End Sub

'3 replace in for meal
Sub Finding()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
  If CDate(Cells(j, 2).Value) = CDate(Cells(j - 1, 2).Value) And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
  And CDate(Cells(j - 2, 2).Value) <> CDate(Cells(j - 3, 2).Value) _
    And Cells(j, 1).Value = "IN" And Cells(j - 1, 1).Value = "OUT" And Cells(j - 2, 1).Value = "IN" _
    Then
   ' Cells(j, 1).EntireRow.Insert
    Cells(j - 1, 1).Value = "MEAL"
    'Cells(j, 2).Value = CDate(Cells(j + 1, 2).Value)

   
  End If
  
Next
Next
End Sub

'4 need to add one missing element and convert dates again and run step # 1 again
Sub FindingFourt()
LR = Cells(Rows.Count, "A").End(xlUp).Row
For j = LR To 2 Step -1


  If Cells(j, 1).Value = "IN" And Cells(j - 1, 1).Value = "MEAL" And Cells(j - 2, 1).Value = "IN" _
  And Cells(j - 3, 1).Value <> "OUT" Then
   Cells(j - 2, 1).EntireRow.insert
   Cells(j - 2, 1).Value = "OUT"
   Cells(j - 2, 2).Value = Cells(j - 1, 2).Value

  ElseIf Cells(j, 1).Value = "IN" And Cells(j - 1, 1).Value = "MEAL" And Cells(j - 2, 1).Value = "OUT" Then
   Cells(j - 1, 1).EntireRow.insert
   Cells(j - 1, 1).Value = "MAEL"
   Cells(j - 1, 2).Value = Cells(j, 2).Value

  ElseIf Cells(j, 1).Value = "IN" And Cells(j - 1, 1).Value = "IN" And Cells(j - 2, 1).Value = "OUT" Then
      Cells(j, 1).EntireRow.insert
      Cells(j, 1).Value = "MEAL"
      Cells(j, 2).Value = Cells(j - 1, 2).Value

  ElseIf Cells(j + 1, 1).Value <> "IN" And Cells(j, 1).Value = "MEAL" _
    And Cells(j - 1, 1).Value = "IN" And Cells(j - 2, 1).Value = "OUT" Then
        Cells(j + 1, 1).EntireRow.insert
        Cells(j + 1, 1).Value = "IN"
        Cells(j + 1, 2).Value = Cells(j, 2).Value

   
  End If
  
   If CDate(Cells(j, 2).Value) = CDate(Cells(j - 1, 2).Value) _
  And CDate(Cells(j - 1, 2).Value) = CDate(Cells(j - 2, 2).Value) _
  And CDate(Cells(j - 2, 2).Value) = CDate(Cells(j - 3, 2).Value) Then
 
    'Cells(j, 1).EntireRow.Insert
    Cells(j, 1).Value = "INN"
    'Cells(j, 1).EntireRow.Insert
    Cells(j - 1, 1).Value = "MEAL"
    Cells(j - 2, 1).Value = "MAEL"
    Cells(j - 3, 1).Value = "OUT"
  End If
  
Next
End Sub

'' 5 if you have the 2 middle ones but not the first nor the last is missing convert dates again
'Sub In_missing_many()
'LR = Cells(Rows.Count, "A").End(xlUp).Row
'BB = Cells(Rows.Count, "B").End(xlUp).Row
'For j = LR To 2 Step -1
' For k = BB To 2 Step -1
'  If Cells(j, 2).Value <> Cells(j - 1, 2).Value And Cells(j - 1, 2).Value = Cells(j - 2, 2).Value And Cells(j - 2, 2).Value <> Cells(j - 3, 2).Value Then
'    Cells(j - 2, 1).EntireRow.Insert
'    Cells(j - 2, 1).Value = "MAEL"
'    Cells(j - 2, 2).Value = Cells(j - 1, 2).Value
'
'    Cells(j, 1).EntireRow.Insert
'    Cells(j, 1).Value = "OUT"
'    Cells(j, 2).Value = Cells(j - 1, 2).Value
'
'    'Cells(j, 1).EntireRow.Insert
'    'Cells(j, 1).Value = "MAEL"
'    'Cells(j, 2).Value = Cells(j - 1, 2).Value
'
'
'   End If
'   'Exit For
'Next
'Next
'End Sub
'6 convert IN TO MAEL
Sub MealMael()
LR = Cells(Rows.Count, "A").End(xlUp).Row
For j = LR To 2 Step -1
  If Cells(j, 1).Value = "MEAL" Then Cells(j - 1, 1).Value = "MAEL"
 
Next
End Sub
'7 convert IN TO INN
Sub InINN()
LR = Cells(Rows.Count, "A").End(xlUp).Row
For j = LR To 2 Step -1
  If Cells(j, 1).Value = "IN" Then Cells(j, 1).Value = "INN"
 
Next
End Sub
'8
'this code should be run at the end to put 0 for empty cells under time column
Sub FillBlanksColC()
  Dim rng As Range
  Dim i As Long
  Dim cell As Range
  Dim sht As Worksheet
  'Set sht = ActiveWorkbook.Sheets("Employee_inv_time_report")
  sht.Activate
  'Range("C12:AL12").Select
  'Range(Selection, Selection.End(xlDown)).Select
  Set rng = Range(Range("C2"), Range("C" & sht.UsedRange.Rows.Count))
  For Each cell In rng
    If cell.Value = "" Then cell.Value = "0"
  Next
End Sub
' one of the last codes whatever is in column b but not in f would show blank spaces
'get investors_data prior manipulation
'Code should be used to compare employee against investors_data code
'and viceverse and make updates and put on bold font matches
'get investors_data unclean in order to perform the comparison
Sub Compare()
Dim Range1 As Range, Range2 As Range, Rng1 As Range, Rng2 As Range
Dim x As Long
xTitleId = "Compare"
Set Range1 = Application.Selection
Set Range1 = Application.InputBox("Range1 :", xTitleId, Range1.Address, Type:=8)
Set Range2 = Application.InputBox("Range2:", xTitleId, Type:=8)
Application.ScreenUpdating = False
For Each Rng1 In Range1
    xValue = Left(Rng1.Value, 10)
     'xValue = Rng1.Value
    For Each Rng2 In Range2
    x = Rng2.Row
    'x = Rng2.Row
        If xValue = Left(Rng2.Value, 10) Then
        
             
           Rng1.Font.Bold = 3
           
        End If
    Next
Next
Application.ScreenUpdating = True
End Sub


Function BoldFont(CellRef As Range)
BoldFont = CellRef.Font.Bold
End Function

Sub Cale()
Call Compare
End Sub
'before running this code copy content of columnd(ctrl+x) and put in col I
Sub Macro5()
Dim lastRow, lastRow2 As Long
'    Columns("D:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    lastRow = Range("C" & Rows.Count).End(xlUp).Row
    lastRow2 = Range("D" & Rows.Count).End(xlUp).Row
    lastRow3 = Range("E" & Rows.Count).End(xlUp).Row
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-INT(RC[-1])"
    Range("D2").Select
    Range("D2").AutoFill Destination:=Range("D2:D" & lastRow) ' autofills
    Range("D2:D" & Range("E" & Rows.Count).End(xlUp).Row).Rows.Copy
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*1440"
    Range("E2").Select
    Range("E2").AutoFill Destination:=Range("E2:E" & lastRow2) 'autofills2
    Range("E2:E" & Range("E" & Rows.Count).End(xlUp).Row).Rows.Copy
''
    
    
    
    
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F1").Value = "Hours"
    Application.CutCopyMode = False
    'Columns("D:I").Delete Shift:=xlShiftToLeft
End Sub
Sub hiu()
    Columns("C:E").Select
    Range("E1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    End Sub

Sub MacroWD()
    Range("D1") = "Weekday"
    Dim lastRow As Long
    lastRow = Range("B" & Rows.Count).End(xlUp).Row
    'ActiveCell.FormulaR1C1 = "Weekday"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""dddd"")"
    Range("D2").Select
    Range("D2").AutoFill Destination:=Range("D2:D" & lastRow)
    Range("D2:D" & Range("D" & Rows.Count).End(xlUp).Row).Rows.Copy
'    Range("D2:D615").Select
'    Selection.Copy
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub Cal()
Call Macro5
Call Macro5
Call hiu
Call MacroWD
End Sub











