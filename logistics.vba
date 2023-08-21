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
    
    For i = 2 To lastRow
        n = ws.Cells(i, 2).Value
        If n <= 0 Then
            prev_computed(i - 1) = n
        Else
            For j = i - 1 To WorksheetFunction.Max(i - 5, 1) Step -1
                If n >= -1 * prev_computed(j) Then
                    n = n + prev_computed(j)
                    prev_computed(j) = 0
                Else
                    prev_computed(j) = n + prev_computed(j)
                    n = 0
                    Exit For
                End If
            Next j
            prev_computed(i - 1) = n
        End If
    Next i
    
    For i = 2 To lastRow
        Dim tempSum As Double
        tempSum = 0
        For j = i - 1 To WorksheetFunction.Max(i - 5, 1) Step -1
            If prev_computed(j) < 0 Then
                tempSum = tempSum + prev_computed(j)
            End If
        Next j
        prev_results(i - 1) = tempSum
    Next i
    
    ' Output to the sheet
    For i = 2 To lastRow
        ws.Cells(i, 3).Value = prev_computed(i - 1)
        ws.Cells(i, 4).Value = prev_results(i - 1)
    Next i

End Sub