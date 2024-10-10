Sub great()
'Define Variables
    Dim Row As Long
    Dim Column As Double
    Dim Ticker_Name As String
    Dim Value As String
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Totes_V As Double
    Dim dataRange As Range
    
    Dim Result_Row As Long
    Dim RowCount As Long
    Dim lastRow As Long
    Dim matchRow As Variant
    
    
    'Print column headers for the summary table
    Cells(1, 18).Value = "Ticker Name"
    Cells(1, 19).Value = "Value"
    
    'Print row labels for summary table
    Cells(2, 17).Value = "Greatest Percent Increase"
    Cells(3, 17).Value = "Greatest Percent Decrease"
    Cells(4, 17).Value = "Greatest Stock Volume"

    ' Find the last used row in column L
    lastRow = Cells(Rows.Count, "L").End(xlUp).Row
    
    ' Set the range dynamically from L1 to the last used row
    Set dataRange = Range("L1:L" & lastRow)
    
    'Maximum
    ' Find the maximum value in the dynamically defined range
    Greatest_Increase = WorksheetFunction.Max(dataRange)
    
     'Print Greatest Increase Value for summary table
    Cells(2, 19).Value = Greatest_Increase
    Cells(2, 19).NumberFormat = "0.00%"
    
    ' Use the Match function to find the value in the range
    matchRow = WorksheetFunction.Match(Greatest_Increase, dataRange, 0)
    
    'Pull Ticker Name that matches greatest increase value
    Ticker_Name = Cells(matchRow, 10).Value

    'Print Ticker name for max Increase
    Cells(2, 18).Value = Ticker_Name
    
    'Minimum
    ' Find the minimum value in the dynamically defined range
    Greatest_Decrease = WorksheetFunction.Min(dataRange)
    
    'Print Greatest Deacrease Value for summary table
    Cells(3, 19).Value = Greatest_Decrease
    Cells(3, 19).NumberFormat = "0.00%"
    
    ' Use the Match function to find the value in the range
    matchRow = WorksheetFunction.Match(Greatest_Decrease, dataRange, 0)
    
     'Pull Ticker Name that matches greatest increase value
    Ticker_Name = Cells(matchRow, 10).Value
    
    'Print Ticker name for Decrease
    Cells(3, 18).Value = Ticker_Name
    
    
   'Greatest Stock Volume
    ' Find the last used row in column M
    lastRow = Cells(Rows.Count, "M").End(xlUp).Row
    
    ' Set the range dynamically from L1 to the last used row
    Set dataRange = Range("M1:M" & lastRow)
    
    ' Find the greatest stock total in the dynamically defined range
    Greatest_Totes_V = WorksheetFunction.Max(dataRange)
    
     'Print greatest stock total for summary table
    Cells(4, 19).Value = Greatest_Totes_V
    
    ' Use the Match function to find the value in the range
    matchRow = WorksheetFunction.Match(Greatest_Totes_V, dataRange, 0)
    
    'Pull Ticker Name that matches greatest stock total
    Ticker_Name = Cells(matchRow, 10).Value

    'Print Ticker name for greatest stock total
    Cells(4, 18).Value = Ticker_Name
    
 
 
 ' Print Ticker Name
        'Ticker_Name = Cells(Row, 10).Value
        'Cells(Result_Row, 16).Value = Ticker_Name

End Sub
