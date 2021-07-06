Attribute VB_Name = "Module2"
Sub StockAnalysis():

'Creating Variables headers for stock_summary_table on each worksheet
    For Each ws In Worksheets
        Dim Ticker_Sym As String
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Ticker_Volume As Double
        Dim Stock_Summary_Row As Integer
    
        'Define Initial Values for variables

        Ticker_Volume = 0
        Stock_Summary_Row = 2
        Open_Price = ws.Cells(2, 3).Value
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Count the numbe of rows in first column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        'Loop through rows by Ticker_Sym
        For T = 2 To lastrow
        
            If ws.Cells(T + 1, 1).Value <> ws.Cells(T, 1).Value Then
            
                Ticker_Sym = ws.Cells(T, 1).Value
                
                Ticker_Volume = Ticker_Volume + ws.Cells(T, 7).Value
                
                'Printing the Stock_Summary_Row values
                ws.Range("I" & Stock_Summary_Row).Value = Ticker_Sym
                
                ws.Range("L" & Stock_Summary_Row).Value = Ticker_Volume
                
                Close_Price = ws.Cells(T, 6).Value
                
                Yearly_Change = (Close_Price - Open_Price)
                
                ws.Range("J" & Stock_Summary_Row).Value = Yearly_Change
                
                If Open_Price = 0 Then
                        
                        Percent_Change = 0
                        
                    Else
                        
                        Percent_Change = Yearly_Change / Open_Price
                        
                End If
                
                ws.Range("K" & Stock_Summary_Row).Value = Percent_Change
                
                ws.Range("K" & Stock_Summary_Row).NumberFormat = "0.00%"
                
                'ws.Range("L" & Stock_Summary_Row).Value = Ticker_Volume
                
                
                'Reset the Variables
                
                Stock_Summary_Row = Stock_Summary_Row + 1
                
                Ticker_Volume = 0
                
                Open_Price = ws.Cells(T + 1, 3)
                
                
                Else
                
                    'Add volume of Stock Trade
                
                    Ticker_Volume = Ticker_Volume + ws.Cells(T, 7).Value
                
            End If
        
        Next T
    'conditional formatting to highlight positive change in green and negative change in red

    Stock_Summary_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For T = 2 To Stock_Summary_Row
            If ws.Cells(T, 10).Value > 0 Then
                ws.Cells(T, 10).Interior.ColorIndex = 10

            Else
                ws.Cells(T, 10).Interior.ColorIndex = 3
            End If
        
        Next T
        
    'Cell labels

            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"

    'Stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
        For T = 2 To Stock_Summary_Row
        
            'Find the maximum and maximum percent change
            If ws.Cells(T, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Stock_Summary_Row)) Then
                ws.Cells(2, 16).Value = ws.Cells(T, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(T, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(T, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Stock_Summary_Row)) Then
                ws.Cells(3, 16).Value = ws.Cells(T, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(T, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(T, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Stock_Summary_Row)) Then
                ws.Cells(4, 16).Value = ws.Cells(T, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(T, 12).Value
            
            End If
        
        Next T

        
    Next ws
End Sub

