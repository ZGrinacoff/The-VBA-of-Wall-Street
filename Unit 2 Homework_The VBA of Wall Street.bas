Attribute VB_Name = "Module1"
Sub Stocks()
    
'Activate variable for each worksheet.
Dim WS As Worksheet

'Activate and loop through all sheets.
For Each WS In Worksheets

    'Set an initial variable for holding the ticker symbol.
    Dim Ticker As String

    'Set an initial variable for holding the Total Volume per stock.
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    'Set an initial variable for holding the Yearly Change per stock.
    Dim Ann_Change As Double
    Ann_Change = 0
    Dim Ann_Close As Double
    Dim Ann_Open As Double
    Ann_Close = 0
    Ann_Open = 0
    
    'Set an initial variable for percent change.
    Dim Percent_Change As Double

    'Keep track of the location for each Stock in the summary brand.
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Set variable for last row.
    Dim Last_Row As Long
    Last_Row = WS.Cells(Rows.Count, 1).End(xlUp).Row

    'Set Column Headers for Summary Table.
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"

    'Loop through all stock activity.
        For i = 2 To Last_Row
            
            'Check if the preceding stock is the same as current stock
            If (WS.Cells(i - 1, 1).Value <> WS.Cells(i, 1).Value) Then
            
                'Set the Annual Open for the current stock.
                Ann_Open = WS.Cells(i, 3).Value
            
            'Check if we are still within the same Stock, if not then...
            ElseIf (WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value) Then
            
                'Set the Stock Ticker.
                Ticker = WS.Cells(i, 1).Value
            
                'Add to the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
            
                'Set the Annual Close for the current stock.
                Ann_Close = WS.Cells(i, 6).Value
            
                'Calculate Annual Change per Stock.
                Ann_Change = Ann_Close - Ann_Open
                    
                    'Handle error for zero values. Skip Percent Change if case is met.
                    If (Ann_Open = 0) Then
                        Percent_Change = 0
                        GoTo Numerr
                    End If
            
                'Calculate the Percent Change per stock.
                Percent_Change = Ann_Change / Ann_Open
                    
Numerr:
                    
                'Print the Stock Ticker in the Summary Table.
                WS.Range("I" & Summary_Table_Row).Value = Ticker
            
                'Print the Total Stock Volume to the Summary Table.
                WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
                'Print the Annual Change per stock in the Summary Table.
                WS.Range("J" & Summary_Table_Row).Value = Ann_Change
                
                    'Conditional Formatting for Yearly Change in Summary Table.
                    If Ann_Change >= 0 Then
                        WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    ElseIf Ann_Change < 0 Then
                        WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    End If
                    
                'Print the Percent Change per stock in the Summary Table.
                WS.Range("K" & Summary_Table_Row).Value = Percent_Change
            
                'Format Percent Change as Percent.
                WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Add one to the summary table row.
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Total Stock Volume.
                Total_Stock_Volume = 0
            
                'Reset Annual Change
                Ann_Change = 0
                Ann_Open = 0
                Ann_Close = 0
            
                'Reset Percent Change
                Percent_Change = 0
            
            'If the cell in the next row is the same Stock Ticker, then...
            Else
        
                'Add to the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
            
            End If
        
        Next i
    
    'Set Column Headers for Greatest Value Summary.
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O3").Value = "Greatest % Decrease"
    WS.Range("O4").Value = "Greatest Total Volume"
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"

    'Create variable for last row in summary table.
    Dim Sum_Last_Row As Long
    Sum_Last_Row = WS.Cells(Rows.Count, "I").End(xlUp).Row

        'Loop through summary table to find max/min % change, and greatest Total Volume, then print ticker along with result. Formatting included.
        For x = 2 To Sum_Last_Row
            
            If (WS.Cells(x, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & Sum_Last_Row))) Then
                WS.Range("P2").Value = WS.Cells(x, 9).Value
                WS.Range("Q2").Value = WS.Cells(x, 11).Value
                WS.Range("Q2").NumberFormat = "0.00%"
            
            ElseIf (WS.Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & Sum_Last_Row))) Then
                WS.Range("P3").Value = WS.Cells(x, 9).Value
                WS.Range("Q3").Value = WS.Cells(x, 11).Value
                WS.Range("Q3").NumberFormat = "0.00%"
            
            ElseIf (WS.Cells(x, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & Sum_Last_Row))) Then
                WS.Range("P4").Value = WS.Cells(x, 9).Value
                WS.Range("Q4").Value = WS.Cells(x, 12).Value
            
            End If
            
        Next x
    
    'Autofit columns A to Q.
    WS.Columns("A:Q").AutoFit

Next WS

End Sub
