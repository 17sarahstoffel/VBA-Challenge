Attribute VB_Name = "Module1"
Sub StockMarket()

 For Each ws In Worksheets
    
        Dim Worksheet_Name As String
        Worksheet_Name = ws.Name
        
        'Creating the column names for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Setting Variable for ticker
        Dim Ticker As String
    
        ' Keeping track of the location for the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        ' Having VBA look for the last row of data
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Setting Yearly Change, when the stock opened and closed as variables
        Dim Yearly_Change As Double
        Dim Stock_Open As Double
        Dim Stock_Close As Double
    
        'Setting variable for percent change
        Dim Percent_Change As Double
    
        ' Setting Variable for the total stock volume
        Dim Stock_Volume As Long
        Total_Stock_Volume = 0
    
    
        For i = 2 To last_row
    
            'Finding the open price
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Stock_Open = ws.Cells(i, 3).Value
            
            End If
    
            'Looking for when the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Putting the ticker names in the new chart
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                ' Finding the ending price
                Stock_Close = ws.Cells(i, 6).Value
            
                ' Finding the yearly change and putting it in the new chart
                Yearly_Change = Stock_Close - Stock_Open
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                'Finding the percent change and putting it in the new chart
                Percent_Change = Yearly_Change / Stock_Open
                ws.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)
            
                ' adding the volume and printing it to the new chart
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
                'Resting the Stock Volume
                Total_Stock_Volume = 0
            
                'Adding one to the summary table row to place next value in the next row
                Summary_Table_Row = Summary_Table_Row + 1
            
            ' When the ticker is the same
            Else
            
                'add the stock voulmes together
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            
            End If
            
        Next i
    
        'Finding the last row in the summary table
        Dim Summary_Last_Row As Integer
        Summary_Last_Row = ws.Cells(Rows.Count, 10).End(xlUp).Row
       
        'Color formating the yearly change
        For j = 2 To Summary_Last_Row
    
            If ws.Cells(j, 10).Value < 0 Then
        
                ws.Cells(j, 10).Interior.ColorIndex = 3 'Red
            
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 4 'Green
            
            End If
            
        Next j
        
        'creating column names for the Greatest table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
       'Creating variables for the greatest increase, decrease, and total volume
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Total_Volume As LongLong
       
        Dim Ticker_Increase As String
        Dim Ticker_Decrease As String
        Dim Ticker_Total As String
        
       
        'Finding the greatest increase, decrease, and total and placing it in the greatest table
        Greatest_Increase = WorksheetFunction.Max(ws.Range("K2:K3001"))
        ws.Cells(2, 17).Value = FormatPercent(Greatest_Increase)
        
        Greatest_Decrease = WorksheetFunction.Min(ws.Range("K2:K3001"))
        ws.Cells(3, 17).Value = FormatPercent(Greatest_Decrease)
        
        Greatest_Total_Volume = WorksheetFunction.Max(ws.Range("L2:L3001"))
        ws.Cells(4, 17).Value = Greatest_Total_Volume
        
        
        'Finding the tickers for the greatest increase, decrease, and total and placing it in the greatest table
        For k = 2 To Summary_Last_Row
            
            If ws.Cells(k, 11).Value = Greatest_Increase Then
                
                Ticker_Increase = ws.Cells(k, 9).Value
                ws.Cells(2, 16).Value = Ticker_Increase
                
            ElseIf ws.Cells(k, 11).Value = Greatest_Decrease Then
            
                Ticker_Decrease = ws.Cells(k, 9).Value
                ws.Cells(3, 16).Value = Ticker_Decrease
                
            ElseIf ws.Cells(k, 12).Value = Greatest_Total_Volume Then
                
                Ticker_Total = ws.Cells(k, 9).Value
                ws.Cells(4, 16).Value = Ticker_Total
                   
            End If
            
            
       Next k
        
        
        
        
    Next ws
        

End Sub
