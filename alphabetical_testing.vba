Attribute VB_Name = "Module1"
Sub Test_stockdata()
    ' Assign a variable to worksheet
    Dim ws As Worksheet
    
    ' Loopthrough all the worksheets
    For Each ws In Worksheets
    
    ' Assign headers for the summary table and final table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percent change"
    ws.Range("L1").Value = "Total Stockvolume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Create variables for to hold value
    Dim Tickername As String
    Dim LastRow As Long
    Dim summaryTableRow As Long
    summaryTableRow = 2
    
    Dim TotalTickerVOlume As Double
    TotalTickerVOlume = 0
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim Percentchange As Double
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    
    Dim lastrowPC As Long
          
    'Determine the lastrow of the raw data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set The initial open price
    OpenPrice = ws.Cells(2, 3).Value
    
    'loop through the ticker symbol
    For i = 2 To LastRow
    
          ' Add values to total ticker(stock) volume
          TotalTickerVOlume = TotalTickerVOlume + ws.Cells(i, 7).Value
       
         ' check whether the next row is same ticker as the previous one or not..
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         ' MsgBox (ws.Cells(i, 1).Value)
          Tickername = ws.Cells(i, 1).Value
       
         'Print the ticker name in Summary table under 'Ticker' header
          ws.Range("I" & summaryTableRow).Value = Tickername
       
         'Print the Total Ticker volume in Summary table under 'Tota Stockvolume' header
          ws.Range("L" & summaryTableRow).Value = TotalTickerVOlume
   
          ' Set the close price
          ClosePrice = ws.Cells(i, 6).Value
          
          ' Calculate yearly change
          YearlyChange = ClosePrice - OpenPrice
          
          'Print the Yearly change in the summarytable under "Yearly change" Header
          ws.Range("J" & summaryTableRow).Value = YearlyChange
          
          'Format the Yearly change in $ format
          ws.Range("J" & summaryTableRow).NumberFormat = "$0.00"
           
          
          'use the conditionals to determine the percent change
          If OpenPrice = 0 Then
          Percentchange = 0
          
          Else
          
          Percentchange = YearlyChange / OpenPrice
          
          End If
          
          'Print the Percent change in the summary table
          ws.Range("k" & summaryTableRow).Value = Percentchange
          
          'Format the percent change in % format
          ws.Range("k" & summaryTableRow).NumberFormat = "0.00%"
           
          'using the conditional formatting to use in Yearly change, green for positive and red for negative
          If YearlyChange < 0 Then
          ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
          
          Else
          
          ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
          
          End If
          
           'Add 1 to summaryTableRow
          summaryTableRow = summaryTableRow + 1
          
          ' Reset total tickervolume to 0
           TotalTickerVOlume = 0
           
           'reset the open price
       OpenPrice = ws.Cells(i + 1, 3).Value
       
       End If
       
    Next i

    'Determine the last row of percent change
    lastrowPC = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'loop through rows for summary table
    For i = 2 To lastrowPC
    
    ' Determine Greatest% increase
    If ws.Range("K" & i).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowPC)) Then
    ws.Range("Q2").Value = ws.Range("K" & i).Value
    ws.Range("P2").Value = ws.Range("I" & i).Value
    ws.Range("Q2").NumberFormat = "0.00%"
    End If
    
   If ws.Range("K" & i).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowPC)) Then
    ws.Range("Q3").Value = ws.Range("K" & i).Value
    ws.Range("P3").Value = ws.Range("I" & i).Value
    ws.Range("Q3").NumberFormat = "0.00%"
    End If
    
    If ws.Range("L" & i).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowPC)) Then
    ws.Range("Q4").Value = ws.Range("L" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    ws.Range("Q4").NumberFormat = "$0.00"
    End If
    
    Next i
    
    Next ws
End Sub
