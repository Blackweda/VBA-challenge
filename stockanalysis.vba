Sub StockAnalysis():

    '   Instructions
    '
    '   Create a script that loops through all the stocks for one year and
    '   outputs the following information:
    '
    '   1) The ticker symbol
    '
    '   2) Yearly change from the opening price at the beginning of a given year
    '   to the closing price at the end of that year.
    '
    '   3) The percentage change from the opening price at the beginning of a given
    '   year to the closing price at the end of that year.
    '
    '   4) The total stock volume of the stock
    '
    '   5) Add functionality to your script to return the stock with the "Greatest %
    '   increase", "Greatest % decrease", and "Greatest total volume"
    '
    '   6) Make the appropriate adjustments to your VBA script to enable it to run on
    '   every worksheet (that is, every year) at once.
    '
    
    
    
    
 
    '   ---------------------------
    '   LOOP THROUGH ALL THE SHEETS
    '   ---------------------------

    For Each ws In Worksheets
    
        ' Assign memory for variables
        Dim StockTicker As String
        
        Dim StockOpenPrice As Double
        Dim StockClosePrice As Double
                
        Dim StockPriceChange As Double
        Dim StockPercentChange As Double
        Dim StockTotalVolume As LongLong
        
        ' https://stackoverflow.com/questions/31436397/vba-integer-overflow-at-70-000
        
        ' Fill in Headers For Analysis Results
        ws.Range("I1").Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
                
        
        ' Determine Last Row of Worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' We will start loop soon, only the headers exist, get first line data
        StockTicker = ws.Cells(2, 1).Value
        StockOpenPrice = ws.Cells(2, 3).Value
        StockTotalVolume = ws.Cells(2, 7).Value
        
        ' Start looping through rows for information
        For i = 2 To LastRow
        
            ' Stock Ticker Hasn't Changed Since Previous Row
            If (StockTicker = ws.Cells(i + 1, 1).Value And i <> 2) Then
            
                ' Add volume to total count
                StockTotalVolume = StockTotalVolume + ws.Cells(i, 7).Value
                
            End If
            
            ' Current Row is Last For Current Ticker
            If (StockTicker <> ws.Cells(i + 1, 1).Value) Then
            '   https://www.automateexcel.com/vba/invalid-qualifier/    --> used to solve the error below:
            '   If (StockTicker.Value <> Cells(i + 1, 1).Value) Then
            
                ' Determine the last row of individual ticker list (Column 9 / I)
                Ticker_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                
                ' Add volume to total count
                StockTotalVolume = StockTotalVolume + ws.Cells(i, 7).Value
                                
                StockClosePrice = ws.Cells(i, 6).Value
                    
                ' Calculate the total year changes
                StockPriceChange = StockClosePrice - StockOpenPrice
                StockPercentChange = StockClosePrice / StockOpenPrice
                    
                'Add a new row to individual ticker analysis row(Column I)
                ws.Cells(Ticker_LastRow + 1, 9).Value = StockTicker
                
                ws.Cells(Ticker_LastRow + 1, 10).Value = StockPriceChange
		' If the Price Change is nothing, its still considered not being a positive move
                If (StockPriceChange <= 0#) Then
                    ws.Cells(Ticker_LastRow + 1, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Ticker_LastRow + 1, 10).Interior.ColorIndex = 4
                End If
                
                ws.Cells(Ticker_LastRow + 1, 11).Value = StockPercentChange
		' If the Percent Change is nothing, its still considered not being a positive move
                If (StockPercentChange > 1#) Then
                    ws.Cells(Ticker_LastRow + 1, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(Ticker_LastRow + 1, 11).Interior.ColorIndex = 3
                End If
                ' Further editing the values for human readability
                ws.Cells(Ticker_LastRow + 1, 11).Value = str((StockPercentChange - 1) * 100) + "%"
                
                ws.Cells(Ticker_LastRow + 1, 12).Value = StockTotalVolume
                
                'Reinitialize Open Price to New Value and Volume to Zero
                StockOpenPrice = ws.Cells(i + 1, 3).Value
                StockTotalVolume = 0
                StockTicker = ws.Cells(i + 1, 1).Value
                
            End If
        
        Next i
        
        ' Determine the Greatest Ticker counts
        Ticker_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim GPI As Double       ' Greatest % Increase
        Dim GPD As Double       ' Greatest % Decrease
        Dim GTV As LongLong     ' Greatest Total Volume
        
        GPI = 1#
        GPD = 1#
        GTV = 0
        
        
        For j = 2 To Ticker_LastRow
        
            If (ws.Cells(j, 11).Value > GPI) Then
                GPI = ws.Cells(j, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
            End If
            If (ws.Cells(j, 11).Value < GPD) Then
                GPD = ws.Cells(j, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(j, 11).Value
            End If
            If (ws.Cells(j, 12).Value > GTV) Then
                GTV = ws.Cells(j, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
            End If
        
        Next j
    
    Next ws
        

End Sub
