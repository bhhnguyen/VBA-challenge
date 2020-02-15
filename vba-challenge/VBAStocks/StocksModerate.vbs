Sub StockModerates():

    dim curTicker as String
    dim curVolume as LongLong
    dim ws as Worksheet
    dim lastRow as Long
    dim tickerRow as Integer
    dim curOpening as Double
    dim lastClosing as Double
    dim change as Double
    dim percentChange as Double
    
    For Each ws in Worksheets

        'Initialize certain values
        tickerRow = 2
        curVolume = 0
        curTicker = ws.Cells(2, 1).Value
        curOpening = ws.Cells(2, 3).Value
        lastClosing = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Add column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        for i = 2 to lastRow

            curVolume = curVolume + ws.Cells(i, 7).value

            'Handle new ticker
            If ws.Cells(i + 1, 1).Value <> curTicker Then
                'Calculate change-related values
                lastClosing = ws.Cells(i, 6).Value
                change = lastClosing - curOpening

                'Divide by 0 handling for percent change
                If curOpening = 0 Then
                    percentChange = 0
                Else
                    percentChange = change/curOpening
                End If

                'Fill in the rows
                ws.Cells(tickerRow, 9).Value = curTicker
                ws.Cells(tickerRow, 10).Value = change
                'Color handling
                If change < 0 Then
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 3
                ElseIf change > 0 Then
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(tickerRow, 11).Value = percentChange
                ws.Cells(tickerRow, 11).Style = "Percent"
                ws.Cells(tickerRow, 12).Value = curVolume

                'Reset handling for next ticker
                curVolume = 0
                curOpening = ws.Cells(i + 1, 3)
                tickerRow = tickerRow + 1
                curTicker = ws.Cells(i + 1, 1).Value


            End If        

        next i

    Next ws



End Sub