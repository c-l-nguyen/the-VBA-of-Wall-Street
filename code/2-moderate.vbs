Sub Moderate()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter As Integer
    Dim yearly_open As Double
    Dim yearly_end As Double
    Dim ticker_row_counter As Double
    
    total_vol = 0
    ticker_counter = 2
    ticker_row_counter = 2
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        total_vol = total_vol + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_row_counter, 3)
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_end = Cells(i, 6)
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = yearly_end - yearly_open
            If yearly_open = 0 Then
                Cells(ticker_counter, 11).Value = Null
            Else
                Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
            End If
            Cells(ticker_counter, 12).Value = total_vol
            
            If Cells(ticker_counter, 10).Value > 0 Then
                Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
                Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            
            Cells(ticker_counter, 11).NumberFormat = "0.00%"
            
            total_vol = 0
            ticker_counter = ticker_counter + 1
            ticker_row_counter = i + 1
        End If
        
    Next i

    Columns("J").Autofit
    Columns("K").Autofit
    Columns("L").Autofit

End Sub

Sub ModerateChallenge()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter As Integer
    Dim yearly_open As Double
    Dim yearly_end As Double
    Dim ticker_row_counter As Double
    
    For Each ws In Worksheets
        total_vol = 0
        ticker_counter = 2
        ticker_row_counter = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(ticker_row_counter, 3)
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearly_end = ws.Cells(i, 6)
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = yearly_end - yearly_open
                If yearly_open = 0 Then
                    ws.Cells(ticker_counter, 11).Value = Null
                Else
                    ws.Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
                End If
                ws.Cells(ticker_counter, 12).Value = total_vol
                
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
                
                total_vol = 0
                ticker_counter = ticker_counter + 1
                ticker_row_counter = i + 1
            End If
            
        Next i

        ws.Columns("J").Autofit
        ws.Columns("K").Autofit
        ws.Columns("L").Autofit

    Next ws
End Sub

Sub ClearModerate()
    Columns("I:L").ClearContents
    Columns("I:L").ClearFormats
    Columns("I:L").UseStandardWidth = True
End Sub

Sub ClearModerateChallenge()
    For Each ws In Worksheets
        ws.Columns("I:L").ClearContents
        ws.Columns("I:L").ClearFormats
        ws.Columns("I:L").UseStandardWidth = True
    Next ws
End Sub