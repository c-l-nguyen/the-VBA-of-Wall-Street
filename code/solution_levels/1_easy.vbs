Sub Easy()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter As Integer
    
    total_vol = 0
    ticker_counter = 2
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        total_vol = total_vol + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        ' If different ticker value, then summarize and reset volume count
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = total_vol
            total_vol = 0
            ticker_counter = ticker_counter + 1
        End If
    Next i

    Columns("J").Autofit

End Sub

Sub EasyChallenge()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter As Integer
    
    For Each ws In Worksheets
    
        total_vol = 0
        ticker_counter = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            ' If different ticker value, then summarize and reset volume count
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = total_vol
                total_vol = 0
                ticker_counter = ticker_counter + 1
            End If
        Next i

        ws.Columns("J").Autofit
    
    Next ws

End Sub

Sub ClearEasy()
    Columns("I:J").ClearContents
    Columns("I:J").ClearFormats
    Columns("I:J").UseStandardWidth = True
End Sub

Sub ClearEasyChallenge()
    For Each ws In Worksheets
        ws.Columns("I:J").ClearContents
        ws.Columns("I:J").ClearFormats
        ws.Columns("I:J").UseStandardWidth = True
    Next ws
End Sub