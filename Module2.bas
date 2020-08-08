Attribute VB_Name = "Module2"
Sub Multipleyear_stockdata_testing()

    'Do the same operations for each sheet in te book
    For Each ws In Worksheets
  
    'Set variables
    Dim Percent As Range
    Dim Vol As Range
    Dim Max As Double
    Dim Min As Double
    Dim VolMax As Double
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim TotalVol As String
    
    'Define the last row
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
    'Set ranges
    Set Percent = ws.Range("K2:K" & LastRow)
    Set Vol = ws.Range("L2:L" & LastRow)
      
    'Determine min or max from the formula that I look at the following link
    '"https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest..
    '..-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475
    Max = Application.WorksheetFunction.Max(Percent)
    'MsgBox (Max)
    Min = Application.WorksheetFunction.Min(Percent)
    'MsgBox (Min)
    VolMax = Application.WorksheetFunction.Max(Vol)
    'MsgBox (VolMax)
    MaxTicker = Application.Index(ws.Range("I:I"), Application.Match(Max, ws.Range("K:K"), 0))
    'MsgBox (MaxTicker)
    MinTicker = Application.Index(ws.Range("I:I"), Application.Match(Min, ws.Range("K:K"), 0))
    'MsgBox (MinTicker)
    TotalVol = Application.Index(ws.Range("I:I"), Application.Match(VolMax, ws.Range("L:L"), 0))
    'MsgBox (TotalVol)
     
    'Set Values and Tickers name
    ws.Range("Q2") = Max
    ws.Range("Q3") = Min
    ws.Range("Q4") = VolMax
    ws.Range("P2") = MaxTicker
    ws.Range("P3") = MinTicker
    ws.Range("P4") = TotalVol
    ws.Range("Q2:Q3").Style = "Percent"
    ws.Columns("O:Q").AutoFit
    Next ws
    
End Sub

