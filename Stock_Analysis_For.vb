Sub Stock_Analysis()

    ' Define variables
    Dim Row As Long
    
    'Define other variables we are collecting
    Dim Ticker_Name As String
    Dim TotalStockVol As Double
         'opening value
            Dim O_Val As Double
        'closing value
            Dim closingVal As Double
        'change from opening and closting
            Dim change As Double
            
    Dim Result_Row As Long
    Dim RowCount As Long
    Dim ws As Worksheet

    
    ' Initialize Result Row (start from row 2 for the summary table)
    Result_Row = 2

    ' Calculate the number of rows in the data set automatically
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    ' Print column headers for the summary table
    Cells(1, 10).Value = "Ticker Name"
    Cells(1, 11).Value = "Quarterly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    'Loop sheet operations through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    ' Loop through rows
        For Row = 2 To RowCount
        
            ' Check if it's the first occurrence of a ticker (new ticker)
            If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' Store the opening value (now in column C)
                O_Val = ws.Cells(Row, 3).Value
                
                ' Print Ticker Name
                Ticker_Name = ws.Cells(Row, 1).Value
                Cells(Result_Row, 10).Value = Ticker_Name
                
                ' Reset TotalStockVol for the new ticker
                TotalStockVol = 0
            End If
    
            ' Add current row's volume to total volume
            TotalStockVol = TotalStockVol + ws.Cells(Row, 6).Value
    
            ' Check if it's the last occurrence of the ticker
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' Calculate the closing value
                closingVal = ws.Cells(Row, 6).Value
                
                ' Calculate change: closing value - opening value
                change = closingVal - O_Val
                
                   ' Print Quarterly Change
                Cells(Result_Row, 11).Value = change
    
    
                ' Color the cell based on the change value
                If change > 0 Then
                    Cells(Result_Row, 11).Interior.ColorIndex = 4 ' Green for positive change
                ElseIf change < 0 Then
                    Cells(Result_Row, 11).Interior.ColorIndex = 3 ' Red for negative change
                Else
                    Cells(Result_Row, 11).Interior.ColorIndex = 0 ' No color for no change
                End If
                
                ' Calculate and Print Percent Change
                If O_Val <> 0 Then
                    Cells(Result_Row, 12).Value = (change / O_Val) * 100
                ' Format the result as a percentage
                    Cells(Result_Row, 12).NumberFormat = "0.00%"
                End If
                
                ' Print Total Stock Volume in the summary table
                Cells(Result_Row, 13).Value = TotalStockVol
    
                ' Move to the next row in compiled results
                Result_Row = Result_Row + 1
            End If
            
        Next Row
        
    Next ws

    MsgBox ("You did it!")
    
End Sub
