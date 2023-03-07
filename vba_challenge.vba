Sub stockdata()

' create a loop and define variable to iterate through each worksheet
For Each ws In Worksheets
    
'lable cell location of each sheet for the summary table
'keep count for each row in the summary table, start row 2 so that headers are in row 1
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim summary_row As Integer
    summary_row = 2
    
'within each worksheet, count the total number of rows and store as "lastrow"
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'for each workskeet assign "ticker" as a string and "volume, close, open, yearly, and percent change as a double", set initial volume to 0
'assign opening price
    Dim ticker As String
    Dim volume As Double
    Dim close_price As Double
    Dim open_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    
    open_price = ws.Cells(2, 3).Value
    volume = 0
    
' creat for loop to check each row in a worksheet, check to see if that row is different then the following row
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
' if the rows are different, ie ticker changes, then calculate yearly change by taking the opening price of the first row and the closing price of the lastrow
        close_price = Cells(i, 6).Value
        yearly_change = (close_price - open_price)
        ws.Cells(summary_row, 10).Value = yearly_change
        
'use conditional formatting to change cell color, red for negative, green for positive
        If yearly_change >= 0 Then
            ws.Cells(summary_row, 10).Interior.ColorIndex = 10
        Else
            ws.Cells(summary_row, 10).Interior.ColorIndex = 3
        End If
        
'print ticker name in summary row under the "ticker"column
        ticker = ws.Cells(i, 1).Value
        ws.Cells(summary_row, 9).Value = ticker
        
'sum volume and print in summary table
'reset volume count to 0 for next ticker summation
        volume = volume + ws.Cells(i, 7).Value
        ws.Cells(summary_row, 12).Value = volume
        volume = 0
'calculate percent change and print to summary row, format to percent cell type
        percent_change = (close_price - open_price) / open_price
        ws.Cells(summary_row, 11).Value = percent_change
        ws.Cells(summary_row, 11).NumberFormat = "0.00%"
         
'add count to summary row
        summary_row = summary_row + 1
        
'reset open price
        open_price = ws.Cells(i + 1, 3).Value
        
'if the rows are the same ticker, sum the volume of the row
        Else
        volume = volume + ws.Cells(i, 7).Value
        
        End If
    
    Next i
'new for loop to go through summary data to find greatest % increase, decrease, and largest volume and label cells
'________________________________________________________________________________________________________________

    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
   
' first find number of rows in summary table
    last_row_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
' for loop to go through each row of summary table
    For i = 2 To last_row_summary
        
' find the the ticker associated with greatest % increase and put value in table
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_summary)) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            
' if value is lower then store  ticker,  greatest decrease in table
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row_summary)) Then
            ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            
' find ticker associated with greatest volume
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_summary)) Then
        ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
          
        End If
    Next i
Next ws

End Sub
