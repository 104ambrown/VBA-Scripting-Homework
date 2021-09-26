Sub VBAChallengeHomework()

'Create a script that will loop through workbook (outer loop)
'Create a script that will loop through all the stocks for one year (inner loop)

'Define variables

'Set title row for columns A to G and corresponding cell references
    Ticker1 = ColumnA Or Cell(i, 1)
    Date = ColumnB Or Cell(i, 2)
    Open = ColumnC or Cell(i, 3)
    High = ColumnD Or Cell(i, 4)
    Low = Column Or Cell(i, 5)
    Close = ColumnF or Cell(i, 6)
    TotalVolume = ColumnG Or Cell(i, 7)

'Output

'Ticker Symbol
    'Find the last row of data with rowCount
    
'Create new columns
    'Title it
        Ticker2 = ColumnI Or Cell(i, 9)
        YearlyChange = ColumnJ Or Cell(i, 10)
        PercentageChange = ColumnK Or Cell(i, 11)
        TotalStockVolume = ColumnL Or Cell(i, 12)
        
'Identify all the different tickers from ColumnA aka ("Ticker 1")
        'Everytime a new ticker is identified in Ticker1, create a value for it in Ticker2
    
'Yearly and percent change
    'Find the first opening price value for each ticker in ColumnC
    'Find the last closing price value for each ticker in ColumnF
        'Subtract Value of ColumnC "Open" from ColumnF "Close"
            'Store that value in ColumnJ "YearlyChange" in the same row as its corresponding ticker.
                
                'Take the value from TickerX,ColumnJ divide it by the first TickerX value in ColumnC
                    'Multiply the above quotient by 100 to find the percentage change
                        'Store that value in ColumnK

            'Conditionally formatting ColumnJ
                If Cell(i, J).Value <> 0 Then Interior.Color.Index = 10  '(I don't like the Neon green of 4)"
                    Else Cell(i, J).Value < 0 then Interior.color.index = 9 '(Again, I don't like the flourescent red 3)"
                Else if (Cell(i, J).Value = 0 then Interior.color.index = 2
                
            
'Calculate total stock volume
    'Find the starting row for a ticker
    'Find the last row for a ticker
        'In columnL, sum the values (iFirst, L) to (iLast, L)
            'Assign this value a new home in
    
    
'I need clarification...........
'For the Greatest % increase and decrease........
'Is the percentage increase and decrease supposed to be of stock price or stock volume?
        
End Sub

