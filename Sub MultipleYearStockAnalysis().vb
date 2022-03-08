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
    dim open as double 
    dim close as double
    dim yearly_change as double
    dim total_stockVol as double
    dim percent_change as double
    dim starting_point as integer

    ' Looping through the worksheets
    For Each ws In Worksheets
    
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
    
    ' Declaring row variable

    last_row = cells(rows.count)
    
    
End Sub

