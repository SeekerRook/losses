Sub ComputeColumns()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim n As Double
    Dim prev_computed(1 To 1000) As Double
    Dim prev_results(1 To 1000) As Double
    
    ' Assuming data is on the first (active) worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Find the last row with data in Column B (prev_years)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 3 To lastRow
        n = ws.Cells(i, 2).Value
        If n <= 0 Then
            prev_computed(i - 1) = n
        Else
            For j = WorksheetFunction.Max(i - 6, 2) To i Step 1
                If n * prev_computed(j) >= 0 Then
                    GoTo Skip1:
                End If
                If n >= -1 * prev_computed(j) Then
                    n = n + prev_computed(j)
                    prev_computed(j) = 0
                Else
                    prev_computed(j) = n + prev_computed(j)
                    n = 0
                    Exit For
                End If
Skip1:
            Next j
            prev_computed(i - 1) = n
        End If
    Next i
    For l = 3 To WorksheetFunction.Max(lastRow - 6, 2) Step 1
        prev_computed(l - 1) = 0
    Next l
    
    Dim tempSum As Double
    tempSum = 0
    For j = i - 1 To WorksheetFunction.Max(lastRow - 5, 2) Step -1
        If prev_computed(j) < 0 Then
            tempSum = tempSum + prev_computed(j)
        End If
    Next j
    prev_result = tempSum
  
    
    ' Output to the sheet
    For i = 3 To lastRow
        ws.Cells(i, 3).Value = prev_computed(i - 1)
      '   ws.Cells(i, 4).Value = prev_results(i - 1)
    Next i
    ws.Cells(lastRow, 4).Value = prev_result
'''''''

  '  Dim ws As Worksheet
  '  Dim lastRow As Long, i As Long, j As Long
  '  Dim n As Double
  '  Dim prev_computed(1 To 1000) As Double
   ' Dim prev_results(1 To 1000) As Double
    
    ' Assuming data is on the first (active) worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Find the last row with data in Column B (prev_years)
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    For i = 3 To lastRow
        n = ws.Cells(i, 6).Value
        If n <= 0 Then
            prev_computed(i - 1) = n
        Else
            For j = WorksheetFunction.Max(i - 6, 2) To i Step 1
                If n * prev_computed(j) >= 0 Then
                    GoTo Skip2:
                End If
                If n >= -1 * prev_computed(j) Then
                    n = n + prev_computed(j)
                    prev_computed(j) = 0
                Else
                    prev_computed(j) = n + prev_computed(j)
                    n = 0
                    Exit For
                End If
Skip2:
            Next j
            prev_computed(i - 1) = n
        End If
    Next i
    For l = 3 To WorksheetFunction.Max(lastRow - 6, 2) Step 1
        prev_computed(l - 1) = 0
    Next l
    
   ' Dim tempSum As Double
    tempSum = 0
    For j = i - 1 To WorksheetFunction.Max(lastRow - 5, 2) Step -1
        If prev_computed(j) < 0 Then
            tempSum = tempSum + prev_computed(j)
        End If
    Next j
    prev_result = tempSum
  
    
    ' Output to the sheet
    For i = 3 To lastRow
        ws.Cells(i, 7).Value = prev_computed(i - 1)
      '   ws.Cells(i, 4).Value = prev_results(i - 1)
    Next i
    ws.Cells(lastRow, 8).Value = prev_result

End Sub
