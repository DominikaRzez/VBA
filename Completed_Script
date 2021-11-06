Sub Market_Stock()

'_______________________________________
'Loop through all the sheets
'_______________________________________
For Each ws In Worksheets


'_______________________________________________________
'Instructions
'_________________________________________________________

Dim last_row As Long
Dim Yearly_Change As Double
Dim open_price As Double
Dim close_price As Double
Dim percentage As Double

    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_row = 2
    Stock_row = 2
    Total_stock = 0
    yearlychange_row = 2
    openprice_row = 2
    closeprice_row = 2
    
    'Loop through the first column to get the Ticker symbol
    For i = 2 To last_row
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            'MsgBox (Cells(i, 1).Value)
            'identify the column to display ticker symbols
             ws.Cells(Ticker_row, 9).Value = ws.Cells(i, 1).Value
             Ticker_row = Ticker_row + 1
        End If
    Next i
    
    
    'Define headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Perecent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    '________________________________________________________________________
    
    'find the opening price at the beginning of the year
    For i = 2 To last_row
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            open_price = ws.Cells(i, 3).Value
            'store the value in seperate column
            ws.Cells(openprice_row, 15).Value = open_price
            openprice_row = openprice_row + 1
        End If
    Next i
   
   'find the closing price at the end of the year
     For i = 2 To last_row
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            close_price = ws.Cells(i, 6).Value
            'store the value in the seperate column
            ws.Cells(closeprice_row, 16).Value = close_price
            closeprice_row = closeprice_row + 1
        End If
    Next i
    
    '__________________________________________________________________________
    'Find last row in the summary table
    summary_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    '___________________________________________________________________
        
    'Yearly change
        For i = 2 To summary_row
            Yearly_Change = ws.Cells(i, 16).Value - ws.Cells(i, 15).Value
            ws.Cells(i, 10).Value = Yearly_Change
    Next i
    
    'Apply conditional formatting
    For i = 2 To summary_row
        If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
        
   '______________________________________________________________________________
   
    'percentage change
        For i = 2 To summary_row
            If ws.Cells(i, 15).Value <> 0 Then
                percentage = (ws.Cells(i, 10).Value / ws.Cells(i, 15).Value)
                ws.Cells(i, 11).Value = percentage
                ws.Range("K2:K" & summary_row).NumberFormat = "0.00%"
            
            Else
                ws.Cells(i, 11).Value = "-"
            
            End If
        Next i
        
        
    '______________________________________________________________________________
        
    'clear the columns storing open price and closing price
        ws.Range("O2:O" & summary_row).Value = ""
        ws.Range("P2:P" & summary_row).Value = ""
    '________________________________________________________________________________
    
    
    'loop through the table to get the stock volume for each ticker
    For i = 2 To last_row
        'count when tickers are the same
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            Total_stock = Total_stock + ws.Cells(i, 7).Value
            
        'define what happens when tickers are different
         ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
            'add last value to the counter
            Total_stock = Total_stock + ws.Cells(i, 7).Value
            'display the value
            ws.Cells(Stock_row, 12).Value = Total_stock
            Stock_row = Stock_row + 1
            'reset counters
            Total_stock = 0
        End If
    Next i
    
    
    '_________________________________________________________________________________
    
        'BONUS
    
    'Define headers
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest total volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    'Find the greatest percent increase
    greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & summary_row))
    ws.Range("P2").Value = greatest_increase
    'Find the row number of the ticker with the greatest_increase
    ticker_increase = WorksheetFunction.Match(greatest_increase, ws.Range("K2:K" & summary_row), 0)
    ws.Range("O2").Value = ticker_increase
    'Adjust the row, to consider the header and match the ticker with the row number
    ws.Range("O2").Value = ws.Cells(ticker_increase + 1, 9)
    
    
    'Find the greatest percent decrease
    greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & summary_row))
    ws.Range("P3").Value = greatest_decrease
    'Find the row number of the ticker with the greatest decrease
    ticker_decrease = WorksheetFunction.Match(greatest_decrease, ws.Range("K2:K" & summary_row), 0)
    'Adjust the row to consider the header and match the ticker with the row number
    ws.Range("O3").Value = ws.Cells(ticker_decrease + 1, 9)
    
    'Find the greatest total volume
    greatest_volume = WorksheetFunction.Max(ws.Range("L2:L" & summary_row))
    ws.Range("P4").Value = greatest_volume
    'Find the row number of the ticker with the greatest volume
    ticker_volume = WorksheetFunction.Match(greatest_volume, ws.Range("L2:L" & summary_row), 0)
    ws.Range("P4").Style = "Normal"
    ws.Range("O4").Value = ticker_volume
    'Adjust the row, to consider the header and match the ticker with the row number
    ws.Range("O4").Value = ws.Cells(ticker_volume + 1, 9)
    
    
    'Autofit the rows and columns
    ws.UsedRange.EntireColumn.AutoFit
    ws.UsedRange.EntireRow.AutoFit
    
'________________________________________________________________
'Instructions complete
'Move to the next sheet
'_____________________________________________________________________

Next ws

    MsgBox ("Completed")
    
    
End Sub
