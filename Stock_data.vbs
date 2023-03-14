Sub StockAnalysis():
     
    'Declare variables
    Dim LastRow, TickerIncrease, TickerDecrease, TickerTotal As String
    Dim Ticker As String
    Dim Volume, OpenValue, CloseValue, YearlyChange, PercentChange, GreatestIncrease, GreatestDecrease, GreatestTotal As Double
    Dim StockCount As Integer
        
    For Each ws In ActiveWorkbook.Worksheets
        
        'Set titles for Columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        'Set titles for Rows for O2: Greatest % Increase, O3: Greatest % Decrease and 04: Greatest Total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Find the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Sset initial values for StockCount and OpenValue
        StockCount = 1
        OpenValue = ws.Cells(2, 3).Value
            
        'Loop through all the Rows with data
        For i = 2 To LastRow
        
            'Check if the next ticker in the list is different from the current ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Increase StockCount by 1
                StockCount = StockCount + 1
                
                'Assign the Ticker and Close Value. Increase Volume by the volume amount for that day
                Ticker = ws.Cells(i, 1).Value
                CloseValue = ws.Cells(i, 6).Value
                Volume = ws.Cells(i, 7).Value + Volume
                
                'Calculate the Yearly Change and the Percent Change for the current stock
                YearlyChange = CloseValue - OpenValue
                PercentChange = YearlyChange / OpenValue
                                
                'Print the values for the current stock Ticker, Yearly Change, PercentChange, and Volume
                ws.Cells(StockCount, 9).Value = Ticker
                ws.Cells(StockCount, 10).Value = YearlyChange
                ws.Cells(StockCount, 11).Value = FormatPercent(PercentChange, 2)
                ws.Cells(StockCount, 12).Value = Volume
                ws.Cells(StockCount, 13).Value = OpenValue
                ws.Cells(StockCount, 14).Value = CloseValue
                
                'Check if the Yearly Change was positive/negative and format the cell color to green/red.
                If YearlyChange >= 0 Then

                       ws.Cells(StockCount, 10).Interior.ColorIndex = 4
                       ws.Cells(StockCount, 11).Interior.ColorIndex = 4

                   Else

                       ws.Cells(StockCount, 10).Interior.ColorIndex = 3
                       ws.Cells(StockCount, 11).Interior.ColorIndex = 3

                   End If
                   
                'Compare the current total volume to the previous greatest total volume. Replace if it is greater.
                If GreatestTotal < Volume Then

                    GreatestTotal = Volume
                    TickerTotal = Ticker

                End If

                'Compare the current percent increase to the previous greatest percent increase. Replace if it is greater.
                If GreatestIncrease < PercentChange Then

                    GreatestIncrease = PercentChange
                    TickerIncrease = Ticker

                End If

                'Compare the current percent decrease to the previous greatest percent decrease. Replace if it is greater.
                If GreatestDecrease > PercentChange Then

                    GreatestDecrease = PercentChange
                    TickerDecrease = Ticker

                End If
                
                'Reset Volume to 0 and OpenValue to the Opening Value of the next stock
                Volume = 0
                OpenValue = ws.Cells(i + 1, 3).Value
                
                
            Else
                
                'Increase the Volume of the current row to the total Volume
                Volume = ws.Cells(i, 7).Value + Volume
            
            End If
    
        Next i
    
        'Print the ticker and values for the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
        ws.Cells(2, 16).Value = TickerIncrease
        ws.Cells(2, 17).Value = FormatPercent(GreatestIncrease, 2)
        ws.Cells(3, 16).Value = TickerDecrease
        ws.Cells(3, 17).Value = FormatPercent(GreatestDecrease, 2)
        ws.Cells(4, 16).Value = TickerTotal
        ws.Cells(4, 17).Value = Format(GreatestTotal, "Scientific")
        
        
        'Reset GreatestIncrease, TickerIncrease, GreatestDecrease, TickerDecrease, GreatestTotal, TickerTotal variables
        GreatestIncrease = 0
        TickerIncrease = ""
        GreatestDecrease = 0
        TickerDecrease = ""
        GreatestTotal = 0
        TickerTotal = ""
        
        'Reset OpenValue and CloseValue
        OpenValue = 0
        CloseValue = 0
    
        'Autfit Columns to
        ws.Columns("I:Q").AutoFit
    
    Next

End Sub