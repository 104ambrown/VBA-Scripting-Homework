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
