Sub formatting()
 'Find last row in the summary table
    summary_row = Cells(Rows.Count, 9).End(xlUp).Row

'Apply conditional formatting
    For i = 2 To summary_row
        If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Else
        Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
    
End Sub
