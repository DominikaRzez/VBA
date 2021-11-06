Sub Summary()

Dim last_row As Long
Dim Yearly_Change As Double
Dim open_price As Double
Dim close_price As Double
Dim percentage As Double

    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_row = 2
    Stock_row = 2
    Total_stock = 0
    yearlychange_row = 2
    openprice_row = 2
    closeprice_row = 2
    
    'Loop through the first column to get the Ticker symbol
    For i = 2 To last_row
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            'MsgBox (Cells(i, 1).Value)
            'identify the column to display ticker symbols
             Cells(Ticker_row, 9).Value = Cells(i, 1).Value
             Ticker_row = Ticker_row + 1
        End If
    Next i
    
    
    'Define headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Perecent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    '________________________________________________________________________
    
    'find the opening price at the beginning of the year
    For i = 2 To last_row
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            open_price = Cells(i, 3).Value
            'store the value in seperate column
            Cells(openprice_row, 15).Value = open_price
            openprice_row = openprice_row + 1
        End If
    Next i
   
   'find the closing price at the end of the year
     For i = 2 To last_row
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            close_price = Cells(i, 6).Value
            'store the value in the seperate column
            Cells(closeprice_row, 16).Value = close_price
            closeprice_row = closeprice_row + 1
        End If
    Next i
    
    '__________________________________________________________________________
    'Find last row in the summary table
    summary_row = Cells(Rows.Count, 9).End(xlUp).Row
    
    '___________________________________________________________________
        
    'Yearly change
        For i = 2 To summary_row
            Yearly_Change = Cells(i, 16).Value - Cells(i, 15).Value
            Cells(i, 10).Value = Yearly_Change
    Next i
        
   '______________________________________________________________________________
   
    'percentage change
        For i = 2 To summary_row
            If Cells(i, 15).Value <> 0 Then
                percentage = (Cells(i, 10).Value / Cells(i, 15).Value)
                Cells(i, 11).Value = percentage
                ActiveSheet.Range("K2:K" & summary_row).NumberFormat = "0.00%"
            
            Else
                Cells(i, 11).Value = "-"
            
            End If
        Next i
        
        
    '______________________________________________________________________________
        
    'clear the columns storing open price and closing price
        ActiveSheet.Range("O2:O" & summary_row).Value = ""
        ActiveSheet.Range("P2:P" & summary_row).Value = ""
    '________________________________________________________________________________
    
    
    'loop through the table to get the stock volume for each ticker
    For i = 2 To last_row
        'count when tickers are the same
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            Total_stock = Total_stock + Cells(i, 7).Value
            
        'define what happens when tickers are different
         ElseIf Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            'add last value to the counter
            Total_stock = Total_stock + Cells(i, 7).Value
            'display the value
            Cells(Stock_row, 12).Value = Total_stock
            Stock_row = Stock_row + 1
            'reset counters
            Total_stock = 0
        End If
    Next i
    
    'Autofit the rows and columns
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.EntireRow.AutoFit
    
    
End Sub
