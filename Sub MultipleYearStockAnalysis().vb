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
    
    ' Creating Analysis Table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "% Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ' Creating Analysis Table and labels for High/Low Stats
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Instantiating data for the first ticker
    open = range("C2")
    percent_min = 100
    percent_max = -100
    total_stockVol = -1

    ' Iterating through the data
    row_index = 2
    total_stockVol = 0
    last_row = cells(rows.count, 1).end(x1up).row
    
    for i = 2 to last_row
        ticker = cells(i, 1).Value
        newTicker = cells(i + 1, 1).Value
        total_stockVol = cells(i, 7).Value
        total_stockVol = total_stockVol + volume

        ' Adding new ticker values and setting volum back to 0 when a new ticker is found
        if (ticker <> newTicker) then
            close = cells(i, 6).value
            cells(row_index, 9).value = ticker
            cells(row_index, 10).value = close - open
            cells(row_index, 10).numberformat = "0.0000000000"

            if (cells(row_index, 10).value >=0) Then
                cells(row_index, 10).interior.colorindex = 10
            else
                cells(row_index, 10).interior.colorindex = 9
            end if

            cells(row_index, 12).value = total_stockVol
    
End Sub

