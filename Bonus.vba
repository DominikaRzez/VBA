Sub Bonus()
'BONUS
    
    'Define headers
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest total volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    'Find last row in the summary table
    summary_row = Cells(Rows.Count, 9).End(xlUp).Row
    
    ActiveSheet.Range("P2:P3").NumberFormat = "0.00%"
    
    'Find the greatest percent increase
    greatest_increase = WorksheetFunction.Max(ActiveSheet.Range("K2:K" & summary_row))
    Range("P2").Value = greatest_increase
    'Find the row number of the ticker with the greatest_increase
    ticker_increase = WorksheetFunction.Match(greatest_increase, ActiveSheet.Range("K2:K" & summary_row), 0)
    Range("O2").Value = ticker_increase
    'Adjust the row, to consider the header and match the ticker with the row number
    Range("O2").Value = Cells(ticker_increase + 1, 9)
    
    
    'Find the greatest percent decrease
    greatest_decrease = WorksheetFunction.Min(ActiveSheet.Range("K2:K" & summary_row))
    Range("P3").Value = greatest_decrease
    'Find the row number of the ticker with the greatest decrease
    ticker_decrease = WorksheetFunction.Match(greatest_decrease, ActiveSheet.Range("K2:K" & summary_row), 0)
    'Adjust the row to consider the header and match the ticker with the row number
    Range("O3").Value = Cells(ticker_decrease + 1, 9)
    
    'Find the greatest total volume
    greatest_volume = WorksheetFunction.Max(ActiveSheet.Range("L2:L" & summary_row))
    Range("P4").Value = greatest_volume
    'Find the row number of the ticker with the greatest volume
    ticker_volume = WorksheetFunction.Match(greatest_volume, ActiveSheet.Range("L2:L" & summary_row), 0)
    Range("P4").Style = "Normal"
    Range("O4").Value = ticker_volume
    'Adjust the row, to consider the header and match the ticker with the row number
    Range("O4").Value = Cells(ticker_volume + 1, 9)
    
    'Autofit the rows and columns
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.EntireRow.AutoFit
    
End Sub
