Sub ModerateChallenge()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_end As Double
    
    For Each ws In Worksheets
        total_vol = 0
        ticker_counter = 2              ' keep track of row to write out ticker summary
        ticker_open_close_counter = 2   ' keep track of row to save off open and close values
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(ticker_open_close_counter, 3)
            
            ' If different ticker value, then summarize
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearly_end = ws.Cells(i, 6)
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = yearly_end - yearly_open
                ' If we have opening value = 0, then just set cell to null
                ' to avoid dividing by 0
                If yearly_open = 0 Then
                    ws.Cells(ticker_counter, 11).Value = Null
                Else
                    ws.Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
                End If
                ws.Cells(ticker_counter, 12).Value = total_vol
                
                ' Color the cell green if > 0, red if < 0
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
                
                ' reset volume count to 0,
                ' move to next row to write ticker summary to in new table,
                ' update to first row of ticker group
                total_vol = 0
                ticker_counter = ticker_counter + 1
                ticker_open_close_counter = i + 1 
            End If
            
        Next i

        ws.Columns("J").Autofit
        ws.Columns("K").Autofit
        ws.Columns("L").Autofit

    Next ws
End Sub

Sub HardChallenge()
    'Note: must run moderate exercise first!
    Call ModerateChallenge()

    For Each ws In Worksheets
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        Dim max, min As Double
        Dim min_row_index, max_row_index, max_total_vol_index As Integer
        Dim max_total_vol As Double
        
        max = 0
        min = 0
        max_total_vol = 0
        
        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            ' replace min/max percentage change if we find lower/higher value
            If ws.Cells(i, 11) > max Then
                max = ws.Cells(i, 11)
                max_row_index = i
            End If
            
            If ws.Cells(i, 11) < min Then
                min = ws.Cells(i, 11)
                min_row_index = i
            End If
            
            ' replace max total volume value if higher value found
            If ws.Cells(i, 12) > max_total_vol Then
                max_total_vol = ws.Cells(i, 12)
                max_total_vol_index = i
            End If
        Next i
        
        ' Write out the values to specified cells
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

Sub ClearModerateChallenge()
    For Each ws In Worksheets
        ws.Columns("I:L").ClearContents
        ws.Columns("I:L").ClearFormats
        ws.Columns("I:L").UseStandardWidth = True
    Next ws
End Sub

Sub ClearHardChallenge()
    Call ClearModerateChallenge

    For Each ws In Worksheets
        ws.Columns("O:Q").ClearContents
        ws.Columns("O:Q").ClearFormats
        ws.Columns("O:Q").UseStandardWidth = True
    Next ws
End Sub