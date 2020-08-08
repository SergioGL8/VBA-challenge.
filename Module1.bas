Attribute VB_Name = "Module1"
Sub Multipleyear_stockdata_testing()

    'For all the sheets
    For Each ws In Worksheets
    
    'Define varibles
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Stock As LongLong
    Dim ClosePrice As Double
    Dim OpenPrice As Double
    'Create a counter for the different open prices
    Dim openp As Double
    Dim SummaryTableRow As Integer
    'Create a counter for the different tickers name
    SummaryTableRow = 2
    openp = 2
    YearlyChange = 0
    PercentChange = 0
    Stock = 0


    'Set headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2:Q3").Style = "Percent"
    ws.Columns("I:L").AutoFit
    ws.Columns("O:Q").AutoFit
    
   'Loop through all daily tickers in first column and last row and if they dont match...
   For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
        'Look for the values
        Ticker = ws.Cells(i, 1).Value
        OpenPrice = ws.Cells(openp, 3).Value
        ClosePrice = ws.Cells(i, 6).Value
        'and do the operations that we need
        Stock = Stock + ws.Cells(i, 7).Value
        YearlyChange = ClosePrice - OpenPrice
            If YearlyChange > 0 Then
            ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            End If
        
            If OpenPrice = 0 Then
            PercentChange = 1
            Else
            PercentChange = (YearlyChange) / OpenPrice
            End If


        'Values into summary
        ws.Cells(SummaryTableRow, 9).Value = Ticker
        ws.Cells(SummaryTableRow, 10).Value = YearlyChange
        ws.Cells(SummaryTableRow, 11).Value = PercentChange
        ws.Cells(SummaryTableRow, 12).Value = Stock
            
        'Writes the next result in the next row
        SummaryTableRow = SummaryTableRow + 1
        'Reset the volume values per each change of ticker
        Stock = 0
        'Writes the next result in the next open price
        openp = i + 1
            
        Else
        
        'Otherwise do the stock summary
        Stock = Stock + ws.Cells(i, 7).Value
        
        End If

    Next i
    
        'Change format for percent
        ws.Columns("K").NumberFormat = "0.00%"
    Next ws

End Sub




