
Sub loopingWorksheets()

    ' Declaring and initializing worksheets
    Dim WSCount As Integer
    WSCount = ActiveWorkbook.Worksheets.Count
    Dim activeWS As Worksheet

    ' For loop to loop through worksheets
    For i = 1 To WSCount
        ActiveWorkbook.Worksheets(i).Activate
        Call MultipleYearStockAnalysis
    Next i
End Sub


Sub MultipleYearStockAnalysis()

    ' Declaring variables for columns in the worksheet
    Dim ticker As String
    Dim newTicker As String
    Dim rowIndex As Double
    Dim last_row As Double
    Dim opening As Double
    Dim closing As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim total_stockVol As Double
    Dim percent_change As Double
    Dim starting_point As Integer
    Dim percent_min As Double
    Dim percent_max As Double
    Dim max_volume As Double
    Dim tickerPMin As String
    Dim tickerPMax As String
    Dim tickerMaxVol As String
    
    ' Creating Analysis Table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Total Stock Volume"
    ' Creating Analysis Table and labels for High/Low Stats
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Instantiating data for the first ticker
    opening = Range("C2")
    percent_min = 100
    percent_max = -100
    total_stockVol = -1

    ' Iterating through the data
    rowIndex = 2
    total_stockVol = 0
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row
        ticker = Cells(i, 1).Value
        newTicker = Cells(i + 1, 1).Value
        total_stockVol = Cells(i, 7).Value
        total_stockVol = total_stockVol + volume

        ' Adding new ticker values and setting volum back to 0 when a new ticker is found
        If (ticker <> newTicker) Then
            closing = Cells(i, 6).Value
            Cells(rowIndex, 9).Value = ticker
            Cells(rowIndex, 10).Value = closing - opening
            Cells(rowIndex, 10).NumberFormat = "0.0000000000"

            If (Cells(rowIndex, 10).Value >= 0) Then
                Cells(rowIndex, 10).Interior.ColorIndex = 10
            Else
                Cells(rowIndex, 10).Interior.ColorIndex = 9
            End If

            Cells(rowIndex, 11).NumberFormat = "0.00%"
            If (opening <> 0) Then
                Cells(rowIndex, 11).Value = (closing - opening) / opening
            Else
                Cells(rowIndex, 11).Value = 0
            End If

            Cells(rowIndex, 12).Value = total_stockVol

            ' Updating pecent minimum and maximum values and total stock volume values
            If (Cells(rowIndex, 11).Value > percent_max) Then
                percent_max = Cells(rowIndex, 11).Value
                tickerPMax = ticker
            Else
                If (Cells(rowIndex, 11).Value < percent_min) Then
                tickerPMin = ticker
            End If
            If (total_stockVol > max_volume) Then
                max_volume = total_stockVol
                tickerMaxVol = ticker
            End If

            rowIndex = rowIndex + 1
            opening = Cells(i + 1, 3)
            total_stockVol = 0

        End If
        Call loopingWorksheets
    End If
    
Next i

    ' Printing all of the results
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("P2").Value = tickerPMax
    Range("P3").Value = tickerPMin
    Range("P4").Value = tickerMaxVol
    Range("Q1").Value = "Value"
    Range("Q2").NumberFormat = "0.00%"
    Range("Q2").Value = percent_max
    Range("Q3").NumberFormat = "0.00%"
    Range("Q3").Value = percent_min
    Range("Q4").Value = max_volume

End Sub


