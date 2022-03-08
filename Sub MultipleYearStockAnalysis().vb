Sub MultipleYearStockAnalysis()

    ' Declaring worksheet variable
    Dim ws As Worksheet
    
    ' Looping through the worksheets
    For Each ws In Worksheets
    
    ' Creating Analysis Tables for All Tickers
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
    
    ' Declaring variables
    
    
    
End Sub

