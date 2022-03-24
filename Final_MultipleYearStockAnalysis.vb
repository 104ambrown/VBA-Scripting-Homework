


Sub MultipleYearStockAnalysis():

    'Loop through all the worksheets
    For each ws in Worksheets


    ' Declaring variables for columns in the worksheet
    Dim ticker As String
    Dim last_row As Long
    Dim SummaryTableRow as Long
    SummaryTableRow = 2
    Dim tickerVol as Double
    tickerVol = 0
    Dim openingYr As Double
    Dim closingYr As Double
    Dim yearly_change As Double
    Dim volume As Double
    volume = 0
    Dim PrevAmt as Long
    PrevAmt = 2
    Dim total_stockVol As Double
    Dim percent_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim max_volume As Double
    max_volume = 0
    Dim tickerPMin As String
    Dim tickerPMax As String
    Dim tickerMaxVol As String
    
    ' Creating Analysis Table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "% Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ' Creating Analysis Table and labels for High/Low Stats
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    

    ' Iterating through the data and finding the last row
    total_stockVol = 0
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
        ' Adding volume
        tickerVol = tickerVol + ws.Cells(i, 7).Value
        ' if the tickers the same keep adding, if it's different reset to 0 and start again
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Setting the ticker and printing it to the summary table
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & SummaryTableRow).value = ticker
        ws.Range("L" & SummaryTableRow).Value = tickerVol
        tickerVol = 0

        'Calculating Yearly open, close, and annual change
        openingYr = ws.Range("C" & PrevAmt)
        closingYr = ws.Range("F" & i)
        yearly_change = closingYr - openingYr
        ws.Range("J" & SummaryTableRow).Value = yearly_change

        'Doing the percentages thing with conditional formating
        if openingYr = 0 Then
            percent_change = 0
        Else
            openingYr = ws.Range("C" & PrevAmt)
            percent_change = yearly_change / openingYr
        End if
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            ws.Range("K" & SummaryTableRow).Value = percent_change
        If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 10
        Else
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 9
        End if

        'Adding another row to the summary table
        SummaryTableRow = SummaryTableRow + 1
        PrevAmt = i + 1
        End if 
    Next i 

    'Finding greatest percent increase/decrease and total volume
    last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To last_row
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
        End If

    Next i 
    ' Formatting the cells to round to 2 decimal places and include a % sign for visual appeasement
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
Next ws

End Sub


