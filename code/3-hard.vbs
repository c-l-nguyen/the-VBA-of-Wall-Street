Sub Hard()
' Note: must run moderate exercise first!

    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
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
    
    Range("O2") = Cells(max_row_index, 9).Value
    Range("O3") = Cells(min_row_index, 9).Value
    Range("O4") = Cells(max_total_vol_index, 9).Value
    
    Range("P2") = max
    Range("P3") = min
    Range("P4") = max_total_vol
    
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    
End Sub

Sub ClearHard()
    Range("N1:P4").ClearContents
    Range("IN1:P4").ClearFormats
End Sub