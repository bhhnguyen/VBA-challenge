Sub StockHard():

    dim curTicker as String
    dim curVolume as LongLong
    dim ws as Worksheet
    dim lastRow as Long
    dim tickerRow as Integer
    dim curOpening as Double
    dim lastClosing as Double
    dim change as Double
    dim percentChange as Double

    'HARD variables
    dim increaseTicker as String
    dim increaseValue as Double
    dim decreaseTicker as String
    dim decreaseValue as Double
    dim volumeTicker as String
    dim volumeValue as LongLong

    increaseTicker = 0
    increaseValue = 0
    decreaseTicker = 0
    decreaseValue = 0
    volumeTicker = 0
    volumeValue = 0
    
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

        'Add new rows/headers for HARD challenge
        ws.Cells(1, 16).Value  = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

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

                'Compare current values to stored ones
                If percentChange > increaseValue Then
                    increaseValue = percentChange
                    increaseTicker = curTicker
                End If

                If percentChange < decreaseValue Then
                    decreaseValue = percentChange
                    decreaseTicker = curTicker
                End If

                If curVolume > volumeValue Then
                    volumeValue = curVolume
                    volumeTicker = curTicker
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

        'Summary stats
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(2, 17).Value = increaseValue
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(3, 17).Value = decreaseValue
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(4, 17).Value = volumeValue

    Next ws



End Sub