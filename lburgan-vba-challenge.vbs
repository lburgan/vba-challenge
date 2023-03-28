Sub StockData()

' Define all variables for this project

Dim vol_total As Double
Dim ticker As String
Dim stock_open As Double
Dim stock_close As Double
Dim lastrow As Long
Dim summary_row As Long
Dim max_percent As Double
Dim min_increase As Double
Dim max_vol As Double

'iterates through all worksheets

For Each ws In Worksheets
    
    'sets up all column titles/ summry table headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'grabbing last row and setting up row counter for summary table

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    
    'loops through all rows
    

    For i = 2 To lastrow
    
    'checking if the ticker value is different than the one after it, if so:
    'set ticker value, put it in correct column, and calculate/color summary varables

        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ticker = ws.Cells(i, 1).Value
            stock_close = ws.Cells(i, 6).Value
            vol_total = vol_total + ws.Cells(i, 7).Value
            ws.Range("I" & summary_row).Value = ticker
            ws.Range("L" & summary_row).Value = vol_total
            vol_total = 0
            ws.Range("J" & summary_row).Value = stock_close - stock_open
            ws.Range("K" & summary_row).Value = ((stock_close - stock_open) / (stock_open)) * 100
            If ws.Cells(summary_row, 11).Value < 0 Then
                ws.Cells(summary_row, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(summary_row, 11).Interior.ColorIndex = 4
            End If
             If ws.Cells(summary_row, 10).Value < 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(summary_row, 10).Value >= 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            End If
            summary_row = summary_row + 1
            
        ' if ticker is not different, just add to total volume
        
        Else
            vol_total = vol_total + ws.Cells(i, 7).Value
        End If
        
          'check if the ticker is different than the one before it, if so, then grab stock opening value
        
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
            stock_open = ws.Cells(i, 3).Value
        End If
    
    Next i
    
   'find the values for the max and min of values
   
    
   max_change = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
   min_change = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
   max_vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
   
   ws.Cells(2, 17).Value = max_change
   ws.Cells(3, 17).Value = min_change
   ws.Cells(4, 17).Value = max_vol
   
   'look for corresponding tickers and print in correct rows
   
   For i = 2 To lastrow
   
        If ws.Cells(i, 11).Value = max_change Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        End If
   
        If ws.Cells(i, 11).Value = min_change Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        End If
   
        If ws.Cells(i, 12).Value = max_vol Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        End If
    Next i
    
   
        

   
   
   
Next ws
End Sub
