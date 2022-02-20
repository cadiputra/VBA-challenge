Sub TickerSummary()
' set variable to hold ticker symbol
    Dim TickerSymbol As String
    
'set variable to hold opening and closing price, yearly change, and percent change
    Dim OpeningPrice, ClosingPrice As Double
    Dim YearlyChange, PercentChange As Double


'set opening date variable
    Dim OpeningDate As String
        
'set variable to hold stock volume total
    Dim StockTotal As Double
    StockTotal = 0

'Bonus: set variable to hold max & min percent change and volume value
'Bonus: set variable to hold max & min ticker symbol
    Dim MaxPercentChange As Double
    MaxPercentChange = 0
    
    Dim MinPercentChange As Double
    MinPercentChange = 0
    
    Dim MaxVolume As Double
    
    MaxVolume = 0
    Dim MaxMinTicker As String


'anchor the location for summary table
    Dim SummaryHeader, SummaryRow As Integer
    SummaryHeader = 1
    SummaryRow = SummaryHeader + 1

'loop through all worksheets
For Each ws In Worksheets
    
    'SUMMARY TABLE
        'print summary header
        ws.Cells(SummaryHeader, 9).Value = "Ticker"
        ws.Cells(SummaryHeader, 10).Value = "Yearly Change"
        ws.Cells(SummaryHeader, 11).Value = "Percent Change"
        ws.Cells(SummaryHeader, 12).Value = "Total Stock Volume"
        ws.Cells(SummaryRow, 15).Value = "Greatest % Increase"
        ws.Cells(SummaryRow + 1, 15).Value = "Greatest % Decrease"
        ws.Cells(SummaryRow + 2, 15).Value = "Greatest Total Volume"
        ws.Cells(SummaryHeader, 16).Value = "Ticker"
        ws.Cells(SummaryHeader, 17).Value = "Value"
        
    
        'determine last row in each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'determine opening date
        OpeningDate = ws.Cells(2, 2).Value
        
            'build ticker symbol summary
            For t = 2 To LastRow
            
            'retrieve the opening price
            'check if date is the first date
            If ws.Cells(t, 2).Value = OpeningDate Then
            OpeningPrice = ws.Cells(t, 3).Value
            'ws.Range("o" & SummaryRow).Value = OpeningPrice
            End If
            
            'check if next row has the same ticker symbol
            If ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
        
            ' retrieve the ticker symbol, closing and opening price
            TickerSymbol = ws.Cells(t, 1).Value
            ClosingPrice = ws.Cells(t, 6).Value
            
            
            'calculate yearly change, percent change and volume total
            YearlyChange = ClosingPrice - OpeningPrice
            PercentChange = YearlyChange / OpeningPrice
            StockTotal = StockTotal + ws.Cells(t, 7).Value
            
            ' print the ticker symbol in summary table
            ws.Range("i" & SummaryRow).Value = TickerSymbol
            'ws.Range("p" & SummaryRow).Value = ClosingPrice
            ws.Range("j" & SummaryRow).Value = YearlyChange
            ws.Range("k" & SummaryRow).Value = PercentChange
            ws.Range("k" & SummaryRow).Style = "Percent"
            ws.Range("l" & SummaryRow).Value = StockTotal
            
            'conditional formatting on yearly change
            If YearlyChange > 0 Then
            ws.Range("j" & SummaryRow).Interior.ColorIndex = 4
            Else
            ws.Range("j" & SummaryRow).Interior.ColorIndex = 3
            End If
            
            
            'Reset volume total
            StockTotal = 0
            
            ' set next summary table row
            SummaryRow = SummaryRow + 1
            
            Else
            
            ' if next row has the same ticker symbol, then add to volume total
            StockTotal = StockTotal + ws.Cells(t, 7).Value
            
            End If
            
            Next t
            
        'reset table summary row
        SummaryRow = SummaryHeader + 1
    
    
    'BONUS TABLE
        For s = 2 To 3001
    
        'check if next row has higher percent change
        If ws.Cells(s, 11).Value > ws.Cells(s + 1, 11).Value And ws.Cells(s, 11).Value > MaxPercentChange Then
        
        'retrieve and print max percent change and its ticker symnbol
        MaxPercentChange = ws.Cells(s, 11).Value
        MaxMinTicker = ws.Cells(s, 9).Value
        ws.Cells(2, 17).Value = MaxPercentChange
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(2, 16).Value = MaxMinTicker
        End If
        
      
    
        'check if next row has lower percent change
        If ws.Cells(s, 11).Value < ws.Cells(s + 1, 11).Value And ws.Cells(s, 11).Value < MinPercentChange Then
        
        'retrieve and print min percent change and its ticker symbol
        MinPercentChange = ws.Cells(s, 11).Value
        MaxMinTicker = ws.Cells(s, 9).Value
        ws.Cells(3, 17).Value = MinPercentChange
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(3, 16).Value = MaxMinTicker
        End If
        
       
    
        'check if next row has higher volume
        If ws.Cells(s, 12).Value > ws.Cells(s + 1, 12).Value And ws.Cells(s, 12).Value > MaxVolume Then
        
        'retrieve and print volume and its ticker symbol
        MaxVolume = ws.Cells(s, 12).Value
        MaxMinTicker = ws.Cells(s, 9).Value
        ws.Cells(4, 17).Value = MaxVolume
        ws.Cells(4, 16).Value = MaxMinTicker
        End If
        
        Next s
        
        'reset MaxPercentChange
        MaxPercentChange = 0
    
        'reset MinPercentChange
        MinPercentChange = 0
    
        'reset volume
        MaxVolume = 0


Next ws
    
End Sub
