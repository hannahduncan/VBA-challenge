Sub stocks()

'Runs through all sheets

For Each ws In Worksheets

    ticker = 0
    StockCounter = 1
    openS = 0
    closeS = 0
    Vol = 0
    YearlyChange = 0
    PercChange = 0
    maxPercChange = 0
    minPercChange = 0
    maxVol = 0

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To lastRow + 1

        ticker = ws.Cells(i + 1, 1).Value

        If i <> 1 Then
            prevTicker = ws.Cells(i, 1).Value
        Else
            prevTicker = ws.Cells(i + 1, 1).Value
            openS = ws.Cells(i + 1, 3).Value
        End If

        If prevTicker = ticker Then
            Vol = Vol + ws.Cells(i + 1, 7).Value
            GoTo NextLoop
        Else
            closeS = ws.Cells(i, 6).Value

            If openS <> 0 Then
                YearlyChange = closeS - openS
                PercChange = YearlyChange / openS
            Else
                YearlyChange = 0
                PercChange = 0
            End If

            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Volume"

            ws.Cells(StockCounter + 1, 9).Value = prevTicker
            ws.Cells(StockCounter + 1, 10).Value = YearlyChange
            ws.Cells(StockCounter + 1, 11).Value = PercChange
            ws.Cells(StockCounter + 1, 11).Style = "Percent"
            ws.Cells(StockCounter + 1, 12).Value = Vol


            If YearlyChange > 0 Then
                ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 3
            End If

            openS = ws.Cells(i + 1, 3).Value
            Vol = 0
            StockCounter = StockCounter + 1

        End If


NextLoop:
    Next i

    For j = 2 To StockCounter - 1

        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"

        currentTicker = ws.Cells(j, 9).Value
        currentVol = ws.Cells(j, 12).Value
        currentPercChange = ws.Cells(j, 11).Value


        If currentPercChange > maxPercChange Then
            maxPercChange = currentPercChange
            currentMaxTicker = currentTicker
            ws.Cells(2, 16).Value = currentMaxTicker
            ws.Cells(2, 17).Value = maxPercChange
            ws.Cells(2, 17).Style = "Percent"

       ElseIf currentPercChange < minPercChange Then
            minPercChange = currentPercChange
            currentMinTicker = currentTicker
            ws.Cells(3, 16).Value = currentMinTicker
            ws.Cells(3, 17).Value = minPercChange
            ws.Cells(3, 17).Style = "Percent"
        End If


        If currentVol > maxVol Then
            maxVol = currentVol
            currentVolTicker = currentTicker
            ws.Cells(4, 16).Value = currentVolTicker
            ws.Cells(4, 17).Value = maxVol
        End If


    Next j


Next ws


End Sub
