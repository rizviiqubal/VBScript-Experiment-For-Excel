Sub Button1_Click()
    Dim int_row
    Dim int_count_unresolved
    Dim int_count_resolved
    Dim int_cat
    Dim int_count_days
    Dim test As Worksheet
    int_rows = Sheets("Report").Cells(Rows.Count, "B").End(xlUp).Row
    
    
     'looping rows
        For i_counter_row = 2 To int_rows
        
           
            
            If Sheets("Report").Cells(i_counter_row, 2) = "New" Or Sheets("Report").Cells(i_counter_row, 2) = "In Progress" Or Sheets("Report").Cells(i_counter_row, 2) = "Reopened" Then
                
                'Sheets("Sheet1").Cells(i_counter_row, 2).Font.Size = 14
                int_count_unresolved = int_count_unresolved + 1
                Sheets("Report").Cells(i_counter_row, 12).Value = 5
                    
            End If
            
            
            
            If Sheets("Report").Cells(i_counter_row, 2) = "Fixed" Or Sheets("Report").Cells(i_counter_row, 2) = "Resolved" Or Sheets("Report").Cells(i_counter_row, 2) = "Verified" Then
                
                'Sheets("Sheet1").Cells(i_counter_row, 2).Font.Size = 14
                int_count_resolved = int_count_resolved + 1
                
                
                diff = TestDates(Sheets("Report").Cells(i_counter_row, 10).Value, Sheets("Report").Cells(i_counter_row, 11).Value)
                
                If diff < 1 Then
                    int_cat = 1
                End If
                
                If diff >= 1 And diff <= 3 Then
                    int_cat = 2
                End If
                
                If diff > 3 And diff <= 7 Then
                    int_cat = 3
                End If
                
                If diff > 7 Then
                    int_cat = 4
                End If
                
                Sheets("Report").Cells(i_counter_row, 12).Value = int_cat
                    
            End If
            
            
            
            
        Next
    
    
End Sub

Function TestDates(pDate1 As Date, pDate2 As Date) As Long

   TestDates = DateDiff("d", pDate1, pDate2)
   
End Function



