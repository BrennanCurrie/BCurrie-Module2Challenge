'Create a script that loops through all the stocks for one year and outputs the following information:
    'The ticker symbol
    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    
Sub StockData()

'establishing loop to run code below across all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

    'setting variable to hold ticker
    Dim Ticker As String
    'setting variable to hold yearly open prices
    Dim YearOpenPrice As Double
    'set variable to hold yearly close price
    Dim YearClosePrice As Double
    'set variable to hold yearly price change
    Dim YearChange As Double
    'set variable to hold percent change
    Dim PercentChange As Double
    'set variable for total stock volume
    Dim TotalStockVolume As Double
        TotalStockVolume = 0
    'set summary row tracker
    Dim SummaryTableRow
        SummaryTableRow = 2
    'set variable to track first row
    Dim FirstRow As Double
        FirstRow = 2
    'set variable to track last row
    'last row function Cells(Rows.Count, "A").End(xlUp).Row
    Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'print headers and format
    ws.Range("J" & 1).Value = "Ticker"
    ws.Range("J" & 1).Font.Bold = True
    ws.Range("J" & 1).Font.Underline = True
    
    ws.Range("K" & 1).Value = "Yearly Change"
    ws.Range("K" & 1).Font.Bold = True
    ws.Range("K" & 1).Font.Underline = True
    
    ws.Range("L" & 1).Value = "Percent Change"
    ws.Range("L" & 1).Font.Bold = True
    ws.Range("L" & 1).Font.Underline = True
    
    ws.Range("M" & 1).Value = "Total Stock Volume"
    ws.Range("M" & 1).Font.Bold = True
    ws.Range("M" & 1).Font.Underline = True
    
    ws.Range("P" & 1).Value = "Ticker"
    ws.Range("P" & 1).Font.Bold = True
    ws.Range("P" & 1).Font.Underline = True

    ws.Range("Q" & 1).Value = "Value"
    ws.Range("Q" & 1).Font.Bold = True
    ws.Range("Q" & 1).Font.Underline = True
        
    ws.Range("O" & 2).Value = "Greatest % Increase"
    ws.Range("O" & 2).Font.Bold = True
    ws.Range("O" & 2).Font.Italic = True
    
    ws.Range("O" & 3).Value = "Greatest % Decrease"
    ws.Range("O" & 3).Font.Bold = True
    ws.Range("O" & 3).Font.Italic = True
    
    ws.Range("O" & 4).Value = "Greatest Total Volume"
    ws.Range("O" & 4).Font.Bold = True
    ws.Range("O" & 4).Font.Italic = True
    
        'FIRST NESTED LOOP
        For i = FirstRow To LastRow

            CurrentTicker = ws.Cells(i, 1).Value
            NextTicker = ws.Cells(i + 1, 1).Value
            PreviousTicker = ws.Cells(i - 1, 1).Value
    
        'looking for middle row
        'Need: stock volume
            If CurrentTicker = PreviousTicker And CurrentTicker = NextTicker Then
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        'looking for first row
        'Need: open price, stock volume
            ElseIf CurrentTicker <> PreviousTicker Then
                YearOpenPrice = ws.Cells(i, 3).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        'looking for last row
        'Need: ticker, closing price, stock volume
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                Ticker = ws.Cells(i, 1).Value
                YearClosePrice = ws.Cells(i, 6)
        
                YearChange = YearClosePrice - YearOpenPrice
                PercentChange = YearChange / YearOpenPrice
        
            'print results
                ws.Range("J" & SummaryTableRow).Value = Ticker
                ws.Range("K" & SummaryTableRow).Value = YearChange
                ws.Range("L" & SummaryTableRow).Value = FormatPercent(PercentChange)
                ws.Range("M" & SummaryTableRow).Value = TotalStockVolume
    
                    If YearChange < 0 Then
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
    
                    ElseIf YearChange > 0 Then
                    ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                    End If
            
            'next summary table row
                SummaryTableRow = SummaryTableRow + 1
             'reset TotalStockVolume
                TotalStockVolume = 0
            End If
        Next i

        '2nd NESTED LOOP-find high percent value

        Dim LastPercentRow As Double
        Dim MaxPercent As Double
        Dim MaxPercentTic As String
            MaxPercent = ws.Cells(2, 12).Value
        Dim Index As Double

        For Index = 2 To LastRow
            If ws.Cells(Index, 12).Value > MaxPercent Then
                MaxPercent = ws.Cells(Index, 12).Value
                MaxPercentTic = ws.Cells(Index, 10).Value
            End If
        Next Index

        'print result
        ws.Range("Q" & 2).Value = FormatPercent(MaxPercent)
        ws.Range("P" & 2).Value = MaxPercentTic

        '3rd NESTED LOOP-find low percent value

        Dim LowPercent As Double
        LowPercent = ws.Cells(2, 12).Value
        Dim LowPercentTic As String
        Dim IndexLow As Double

        For IndexLow = 2 To LastRow
            If Cells(IndexLow, 12).Value < LowPercent Then
                LowPercent = ws.Cells(IndexLow, 12).Value
                LowPercentTic = ws.Cells(IndexLow, 10).Value
            End If
        Next IndexLow

        'print result
        ws.Range("P" & 3).Value = LowPercentTic
        ws.Range("Q" & 3).Value = FormatPercent(LowPercent)
    
        '4th NESTED LOOP-find largest volume

        Dim HighVol As Double
            HighVol = ws.Cells(2, 13).Value
        Dim HighVolTic As String
        Dim IndexVol As Double

        For IndexVol = 2 To LastRow
            If ws.Cells(IndexVol, 13).Value > HighVol Then
                HighVol = ws.Cells(IndexVol, 13).Value
                HighVolTic = ws.Cells(IndexVol, 10).Value
            End If
    
        Next IndexVol

        'print result
        ws.Range("P" & 4).Value = HighVolTic
        ws.Range("Q" & 4).Value = HighVol
        
Next
End Sub
