Sub stocks1()

'Runs through only one sheet

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

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To lastRow + 1

        ticker = Cells(i + 1, 1).Value

        If i <> 1 Then
            prevTicker = Cells(i, 1).Value
        Else
            prevTicker = Cells(i + 1, 1).Value
            openS = Cells(i + 1, 3).Value
        End If

        If prevTicker = ticker Then
            Vol = Vol + Cells(i + 1, 7).Value
            GoTo NextLoop
        Else
            closeS = Cells(i, 6).Value

            If openS <> 0 Then
                YearlyChange = closeS - openS
                PercChange = YearlyChange / openS
            Else
                YearlyChange = 0
                PercChange = 0
            End If

            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Volume"

            Cells(StockCounter + 1, 9).Value = prevTicker
            Cells(StockCounter + 1, 10).Value = YearlyChange
            Cells(StockCounter + 1, 11).Value = PercChange
            Cells(StockCounter + 1, 11).Style = "Percent"
            Cells(StockCounter + 1, 12).Value = Vol


            If YearlyChange > 0 Then
                Cells(StockCounter + 1, 10).Interior.ColorIndex = 4
            Else
                Cells(StockCounter + 1, 10).Interior.ColorIndex = 3
            End If

            openS = Cells(i + 1, 3).Value
            Vol = 0
            StockCounter = StockCounter + 1

        End If


NextLoop:
    Next i

    For j = 2 To StockCounter - 1

        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"

        currentTicker = Cells(j, 9).Value
        currentVol = Cells(j, 12).Value
        currentPercChange = Cells(j, 11).Value


        If currentPercChange > maxPercChange Then
            maxPercChange = currentPercChange
            currentMaxTicker = currentTicker
            Cells(2, 16).Value = currentMaxTicker
            Cells(2, 17).Value = maxPercChange
            Cells(2, 17).Style = "Percent"

       ElseIf currentPercChange < minPercChange Then
            minPercChange = currentPercChange
            currentMinTicker = currentTicker
            Cells(3, 16).Value = currentMinTicker
            Cells(3, 17).Value = minPercChange
            Cells(3, 17).Style = "Percent"
        End If


        If currentVol > maxVol Then
            maxVol = currentVol
            currentVolTicker = currentTicker
            Cells(4, 16).Value = currentVolTicker
            Cells(4, 17).Value = maxVol
        End If


    Next j


End Sub
