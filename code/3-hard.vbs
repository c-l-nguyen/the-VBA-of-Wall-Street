Sub Hard()
' Note: must run moderate exercise first!

    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    Dim max, min As Double
    Dim rownum As Integer
    Dim min_row_index, max_row_index, max_total_vol_index As Integer
    Dim max_total_vol As Double
    max = 0
    min = 0
    max_total_vol = 0
    rownum = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To rownum
        If Cells(i, 11) > max Then
            max = Cells(i, 11)
            max_row_index = i
        End If
        
        If Cells(i, 11) < min Then
            min = Cells(i, 11)
            min_row_index = i
        End If
        
        If Cells(i, 12) > max_total_vol Then
            max_total_vol = Cells(i, 12)
            max_total_vol_index = i
        End If
    Next i
    
    Range("P2") = Cells(max_row_index, 9).Value
    Range("P3") = Cells(min_row_index, 9).Value
    Range("P4") = Cells(max_total_vol_index, 9).Value
    
    Range("Q2") = max
    Range("Q3") = min
    Range("Q4") = max_total_vol
    
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"

    Columns("O").Autofit
    Columns("P").Autofit
    Columns("Q").Autofit
End Sub

Sub HardChallenge()
' Note: must run moderate exercise first!
    For Each ws In Worksheets
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        Dim max, min As Double
        Dim rownum As Integer
        Dim min_row_index, max_row_index, max_total_vol_index As Integer
        Dim max_total_vol As Double
        max = 0
        min = 0
        max_total_vol = 0
        rownum = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To rownum
            If ws.Cells(i, 11) > max Then
                max = ws.Cells(i, 11)
                max_row_index = i
            End If
            
            If ws.Cells(i, 11) < min Then
                min = ws.Cells(i, 11)
                min_row_index = i
            End If
            
            If ws.Cells(i, 12) > max_total_vol Then
                max_total_vol = ws.Cells(i, 12)
                max_total_vol_index = i
            End If
        Next i
        
        ws.Range("P2") = ws.Cells(max_row_index, 9).Value
        ws.Range("P3") = ws.Cells(min_row_index, 9).Value
        ws.Range("P4") = ws.Cells(max_total_vol_index, 9).Value
        
        ws.Range("Q2") = max
        ws.Range("Q3") = min
        ws.Range("Q4") = max_total_vol
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Columns("O").Autofit
        ws.Columns("P").Autofit
        ws.Columns("Q").Autofit
    
    Next ws
End Sub

Sub ClearHard()
    Columns("O:Q").ClearContents
    Columns("O:Q").ClearFormats
    Columns("O:Q").UseStandardWidth = True
End Sub

Sub ClearHardChallenge()
    For Each ws In Worksheets
        ws.Columns("O:Q").ClearContents
        ws.Columns("O:Q").ClearFormats
        ws.Columns("O:Q").UseStandardWidth = True
    Next ws
End Sub