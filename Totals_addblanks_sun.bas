Attribute VB_Name = "Totals_addblanks_sun"


'code to add new column names
Sub CommandButton2_Click()
    
    'Dim myValue As Variant
    For Each Worksheet In ActiveWorkbook.Worksheets

            Range("A3").Value = "Weekday"
            Range("B3").Value = "Punch Date"
            Range("C3").Value = "Total time via ADP punches"
            Range("D3").Value = "Total for week via ADP punches"
            Range("E3").Value = "Variance"
            Range("F3").Value = "Total time paid via ADP"
            Range("G3").Value = "Total for week-paid"
            Range("H3").Value = "Pay Code"
            
    Next Worksheet
    
End Sub



'Code2 to add weekends
Sub In_missing_Saturday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
 
    'If Cells(j, 1).Value = "Monday" And Cells(j + 1, 1).Value <> "Sunday" Then
    If Cells(j, 1).Value = "Monday" Then
    Cells(j + 1, 1).EntireRow.insert
    'Cells(j, 1).Value = "Sunday"
    'Cells(j + 1, 1).EntireRow.Insert
    
   End If
   Exit For
Next
Next
End Sub


Sub In_missing_blank()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
    
  If Cells(j, 1).Value = "Saturday" Then
    Cells(j, 1).EntireRow.insert
    Cells(j, 1).EntireRow.insert
    Cells(j, 1).EntireRow.insert
   End If
 
   
   
   Exit For
  
Next
Next
End Sub
'this code should be run after Code3 below (In_Totals())
Sub In_missing_Sunday()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
 
   If Cells(j, 1).Value = "Monday" Then
    'Cells(j, 1).EntireRow.Insert
     Cells(j + 1, 1).Value = "Sunday"
     Cells(j + 1, 2).Value = CDate(Cells(j, 2).Value - 1)
   End If
   Exit For
Next
Next
   
   Rows(2).EntireRow.Delete
   Rows(3).EntireRow.Delete
   Rows(4).EntireRow.Delete
   Rows(2).EntireRow.Delete
End Sub

Sub En()
Call In_missing_blank
Call In_missing_Sunday
End Sub



'Code3 to add whether is equal, less or more and get the total hrs
Sub In_totals()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
 
    If Cells(j, 4).Value = Cells(j, 7).Value And Cells(j, 1).Value = "Monday" And Cells(j + 1, 1).Value <> "Sunday" Then
        Cells(j + 1, 2).Value = "No Variance"
    ElseIf Cells(j, 4).Value = Cells(j, 7).Value And Cells(j, 1).Value = "Sunday" Then
        Cells(j + 1, 2).Value = "No Variance"
    ElseIf Cells(j, 4).Value < Cells(j, 7).Value And Cells(j, 1).Value = "Monday" And Cells(j + 1, 1).Value <> "Sunday" Then
        Cells(j + 1, 2).Value = "Adjustment in employee's favor"
    ElseIf Cells(j, 4).Value < Cells(j, 7).Value And Cells(j, 1).Value = "Sunday" Then
        Cells(j + 1, 2).Value = "Adjustment in employee's favor"
        
    
    ElseIf Cells(j, 4).Value > Cells(j, 7).Value And Cells(j, 4).Value >= 40 And Cells(j, 7).Value >= 40 Then
        G = CDec(Cells(j, 4).Value - Cells(j, 7).Value)
        Cells(j + 1, 2).Value = "Variance of " & Format(Round(G, 2), "##.00") & " of" & " OT"
        
        Cells(j + 1, 5).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 6).Value = "OT"
        
    ElseIf Cells(j, 4).Value > Cells(j, 7).Value And Cells(j, 4).Value <= 40 And Cells(j, 7).Value < 40 Then
        G = CDec(Cells(j, 4).Value - Cells(j, 7).Value)
        Cells(j + 1, 2).Value = "Variance of " & Format(Round(G, 2), "##.00") & " of" & " RT"
        
        Cells(j + 1, 5).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 6).Value = "RT"
    
   End If
   Exit For
Next
Next
End Sub

Sub Fin()
Call CommandButton2_Click
Call In_missing_Saturday
Call In_missing_Sunday
Call In_totals
End Sub


Sub Formuoli3()
Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws
         .Select
            Call Fin
        End With
    Next
End Sub

'Code3 to add whether is equal, less or more and get the total hrs
Sub In_equal()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
'n = 40
For j = LR To 2 Step -1
 For k = BB To 2 Step -1
 
    If Cells(j, 4).Value = Cells(j, 8).Value And Cells(j, 1).Value = "Monday" And Cells(j + 1, 1).Value <> "Sunday" Then
        Cells(j + 1, 2).Value = "No Variance"
    ElseIf Cells(j, 4).Value = Cells(j, 8).Value And Cells(j, 1).Value = "Sunday" Then
        Cells(j + 1, 2).Value = "No Variance"
    ElseIf Cells(j, 4).Value < Cells(j, 8).Value And Cells(j, 1).Value = "Monday" And Cells(j + 1, 1).Value <> "Sunday" Then
        Cells(j + 1, 2).Value = "Adjustment in employee's favor"
    ElseIf Cells(j, 4).Value < Cells(j, 8).Value And Cells(j, 1).Value = "Sunday" Then
        Cells(j + 1, 2).Value = "Adjustment in employee's favor"
        
    ElseIf Cells(j, 4).Value > Cells(j, 8).Value Then
    G = CDec(Cells(j, 4).Value - Cells(j, 8).Value)
        If G < 0.02 Then
           Cells(j + 1, 2).Value = "No Variance"
    
    ElseIf Cells(j, 4).Value > Cells(j, 8).Value And Cells(j, 4).Value >= 40 And Cells(j, 8).Value >= 40 Then
        G = CDec(Cells(j, 4).Value - Cells(j, 8).Value)
        Cells(j + 1, 2).Value = "Variance of " & Format(Round(G, 2), "##.00") & " of" & " OT"
        
        Cells(j + 1, 6).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 7).Value = "OT"
        
    ElseIf Cells(j, 4).Value > Cells(j, 8).Value And Cells(j, 4).Value <= 40 And Cells(j, 8).Value < 40 Then
        G = CDec(Cells(j, 4).Value - Cells(j, 8).Value)
        Cells(j + 1, 2).Value = "Variance of " & Format(Round(G, 2), "##.00") & " of" & " RT"
        
        Cells(j + 1, 6).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 7).Value = "RT"
        
    
        
   End If
   End If
   Exit For
Next
Next
End Sub
'This code adds total as RT when a person took a SICK day or PTO during week
'during the week even if employee works >40 hrs employee still gets paid as RT
Sub po()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
For k = BB To 2 Step -1
If Cells(j, 10).Value = "STOP" And Cells(j, 9).Value = "SICK" Or Cells(j, 9).Value = "PTO" Or Cells(j, 9).Value = "HOLDAY" Or Cells(j, 9).Value = "SPECIALTIME" _
Or Cells(j - 1, 9).Value = "SICK" Or Cells(j - 1, 9).Value = "PTO" Or Cells(j - 1, 9).Value = "HOLDAY" Or Cells(j - 1, 9).Value = "SPECIALTIME" _
Or Cells(j - 2, 9).Value = "SICK" Or Cells(j - 2, 9).Value = "PTO" Or Cells(j - 2, 9).Value = "HOLDAY" Or Cells(j - 2, 9).Value = "SPECIALTIME" _
Or Cells(j - 3, 9).Value = "SICK" Or Cells(j - 3, 9).Value = "PTO" Or Cells(j - 3, 9).Value = "HOLDAY" Or Cells(j - 3, 9).Value = "SPECIALTIME" _
Or Cells(j - 4, 9).Value = "SICK" Or Cells(j - 4, 9).Value = "PTO" Or Cells(j - 4, 9).Value = "HOLDAY" Or Cells(j - 4, 9).Value = "SPECIALTIME" Then

   If Cells(j, 4).Value > Cells(j, 8).Value Then

        G = CDec(Cells(j, 4).Value - Cells(j, 8).Value)
        
        If G > 0.01 Then
            
            Cells(j + 1, 2).Value = "Variance of " & Format(Round(G, 2), "##.00") & " of" & " RT"
            Cells(j + 1, 6).Value = Format(Round(G, 2), "##.00")
            Cells(j + 1, 7).Value = "RT"
            
        
    ElseIf Cells(j, 10).Value <> "STOP" Then
    If Cells(j, 4).Value > Cells(j, 8).Value And Cells(j, 4).Value > 40 And Cells(j, 8).Value < 40 Then
        G = CDec(Cells(j, 4).Value - 40) 'OT
        H = 40 - CDec(Cells(j, 4).Value) 'RT
        Cells(j + 1, 6).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 7).Value = "OT"
        
        Cells(j + 2, 6).Value = Format(Round(H, 2), "##.00")
        Cells(j + 2, 7).Value = "RT"
End If
End If
End If
End If
'End If
Exit For
Exit For
Next
Next
End Sub

'code to perform calculations (total time per week times hourly rate or OT)
Sub Final_calculation()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
For k = BB To 2 Step -1
  x = Format(18.11, "##.00")
  y = (Format(18.11, "##.00") * 1.5)
  If Cells(j, 7).Value = "RT" Then
    'Cells(j, 3).Value = Format(Round(Cells(j, 3).Value * x, 2), "$##.00")
    Cells(j, 5).Value = Format(Round(Cells(j, 6).Value * x, 2), "$##.00")
     
  ElseIf Cells(j, 7).Value = "OT" Then
     Cells(j, 5).Value = Format(Round(Cells(j, 6).Value * y, 2), "$##.00")
      
     'Selection.Columns("C").NumberFormat = "$##.00"
     'Sheets("Sheet1").Columns("A").Style = "Currency"
   End If
   Exit For
Next
Next
End Sub



Sub llamar()
Call In_equal
Call po
'Call llamar2
'Call Final_calculation
End Sub

Sub llamarrr()
LR = Cells(Rows.Count, "A").End(xlUp).Row
BB = Cells(Rows.Count, "B").End(xlUp).Row
For j = LR To 2 Step -1
For k = BB To 2 Step -1
  
 If Cells(j, 10).Value <> "STOP" Then
    If Cells(j, 4).Value > Cells(j, 8).Value And Cells(j, 4).Value > 40 And Cells(j, 8).Value < 40 Then
        G = CDec(Cells(j, 4).Value - 40) 'OT
        H = 40 - CDec(Cells(j, 8).Value) 'RT
        Cells(j + 1, 6).Value = Format(Round(G, 2), "##.00")
        Cells(j + 1, 7).Value = "OT"
        
        Cells(j + 2, 6).Value = Format(Round(H, 2), "##.00")
        Cells(j + 2, 7).Value = "RT"
  
   End If
   End If
   Exit For
   Exit For
Next
Next
End Sub








