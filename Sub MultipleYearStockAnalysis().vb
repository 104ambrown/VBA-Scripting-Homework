Sub loopingWorksheets ()

    ' Declaring and initializing worksheets
    dim WSCount as integer
    WSCount = ActiveWorkbook.Worksheets.count
    dim activeWS as Worksheet

    ' For loop to loop through worksheets
    For i = 1 to WSCount
        ActiveWorkbook.Worksheets(i).Activate 
        Call MultipleYearStockAnalysis
    Next i
End Sub

Sub MultipleYearStockAnalysis()

    ' Declaring worksheet variable
    dim ws As Worksheet
    ' Declaring variables for columns in the worksheet
    dim ticker as string
    dim newTicker as string
    dim rowIndex as double
    dim last_row as double
    dim open as double 
    dim close as double
    dim yearly_change as double
    dim volume as double
    dim total_stockVol as double
    dim percent_change as double
    dim starting_point as integer
    dim percent_min as double
    dim percent_max as double
    dim max_volume as double
    dim tickerPMin as string
    dim tickerPMax as string
    dim tickerMaxVol as string
    
    ' Creating Analysis Table headers
    ws.Range("I1").value = "Ticker"
    ws.Range("J1").value = "Yearly Change"
    ws.Range("K1").value = "% Change"
    ws.Range("L1").value = "Total Stock Volume"
    ' Creating Analysis Table and labels for High/Low Stats
    ws.Range("O2").value = "Greatest % Increase"
    ws.Range("O3").value = "Greatest % Decrease"
    ws.Range("O4").value = "Greatest Total Volume"
    ws.Range("P1").value = "Ticker"
    ws.Range("Q1").value = "Value"
    
    'Instantiating data for the first ticker
    open = range("C2")
    percent_min = 100
    percent_max = -100
    total_stockVol = -1

    ' Iterating through the data
    rowIndex = 2
    total_stockVol = 0
    last_row = cells(rows.count, 1).end(x1up).row
    
    for i = 2 to last_row
        ticker = cells(i, 1).value
        newTicker = cells(i + 1, 1).value
        total_stockVol = cells(i, 7).value
        total_stockVol = total_stockVol + volume

        ' Adding new ticker values and setting volum back to 0 when a new ticker is found
        if (ticker <> newTicker) then
            close = cells(i, 6).value
            cells(rowIndex, 9).value = ticker
            cells(rowIndex, 10).value = close - open
            cells(rowIndex, 10).numberformat = "0.0000000000"

            if (cells(rowIndex, 10).value >=0) Then
                cells(rowIndex, 10).interior.colorindex = 10
            else
                cells(rowIndex, 10).interior.colorindex = 9
            end if

            cells(rowIndex, 11).numberformat = "0.00%"
            If (open <> 0) then
                cellscells(rowIndex, 11).value = (close - open) / open
            else
                cells(rowIndex, 11).value = 0
            end if

            cells(rowIndex, 12).value = total_stockVol

            ' Updating pecent minimum and maximum values and total stock volume values
            if (cells(rowIndex, 11).value < percent_max) then
                percent_maz = cells(rowIndex, 11).value
                tickerPMax = ticker
            else if (cells(rowIndex, 11).value < percent_min) then
                tickerPMin = ticker
            end if 
            if (total_stockVol > max_volume) then
                max-volume = total_stockVol
                tickerMaxVol = ticker
            end if

            rowIndex = rowIndex + 1
            open = cells(i + 1, 3)
            total_stockVol = 0

        end if 

    next i

    ' Printing all of the results
    range("O2").value = "Greatest % Increase"
    range("O3").value = "Greatest % Decrease"
    range("O4").value = "Greatest Total Volume"
    range("P1").value = "Ticker"
    range("P2").value = tickerPMax
    range("P3").value = tickerPMin
    range("P4").value = tickerMaxVol
    range("Q1").value = "Value"
    range("Q2").numberformat = "0.00%"
    range("Q2").value = percent_max
    range("Q3").numberformat = "0.00%"
    range("Q3").value = percent_min
    range("Q4").value = max_volume

End Sub

