Sub StockMarketReview()

    ' Set inital variable for ticker symbol, stock volume, open price, close price
    Dim Ticker_Name, GITicker, GDTicker, GVTicker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim GreatestIncrease, GreatestDecrease, GreatestVolume As Double
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    ' First time seeing ticker symbol tracker
    Dim First_Time As Integer
    First_Time = 1
    
    ' Summary Row Tracker
    Dim Summary_Row As Integer
    Summary_Row = 2
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
        ws.Activate
        
        'Adding Summary Table Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Determine Last Row and Loop through all rows of stocks
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
        
            ' Check if still within same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set Ticker Name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Final stock sum
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'Set Close_Price
                Close_Price = ws.Cells(i, 6).Value
                
                'Print ticker name to summary table
                ws.Range("I" & Summary_Row).Value = Ticker_Name
                
                'Print total volume to summary table
                ws.Range("L" & Summary_Row).Value = Total_Volume
                
                'Print Yearly Change
                ws.Range("J" & Summary_Row).Value = Close_Price - Open_Price
                
                'Print Percent Change and Conditional Formating for Yearly Change
                If Close_Price = Open_Price Then
                    ws.Range("K" & Summary_Row).Value = 0
                    ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                Else
                    If Close_Price > Open_Price Then
                        ws.Range("K" & Summary_Row).Value = (Close_Price / Open_Price)
                        ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("K" & Summary_Row).Value = (1 - (Close_Price / Open_Price))
                        ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                        ws.Range("K" & Summary_Row).Value = Range("K" & Summary_Row).Value * -1
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                    End If
                End If
                
                'Track Greatest Values
                If ws.Range("K" & Summary_Row).Value > GreatestIncrease Then
                    GreatestIncrease = ws.Range("K" & Summary_Row).Value
                    GITicker = ws.Cells(i, 1).Value
                End If
                
                If ws.Range("K" & Summary_Row).Value < GreatestDecrease Then
                    GreatestDecrease = ws.Range("K" & Summary_Row).Value
                    GDTicker = ws.Cells(i, 1).Value
                End If
                If ws.Range("L" & Summary_Row).Value > GreatestVolume Then
                    GreatestVolume = ws.Range("L" & Summary_Row).Value
                    GVTicker = ws.Cells(i, 1).Value
                End If
                
                ' Increment summary row by 1
                Summary_Row = Summary_Row + 1
                
                'Reset total volume
                Total_Volume = 0
                
                ' Reset First time tracker
                First_Time = 1
                
                
            'If ticker symbol is the same
            Else
                If First_Time = 1 Then
                    Open_Price = Cells(i, 3).Value
                    First_Time = 0
                End If
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        'Adding Challenge Table Headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(2, 15).Value = GITicker
        ws.Cells(2, 16).Value = GreatestIncrease
            ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = GDTicker
        ws.Cells(3, 16).Value = GreatestDecrease
            ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = GVTicker
        ws.Cells(4, 16).Value = GreatestVolume
        
        'Reset Summary Table Rows for next sheet
        Summary_Row = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        ws.Columns.AutoFit
        
    Next ws

End Sub
