Sub StockExchange()
    
    'SETUP
    '-----------------------------------------------------------------------------
    
    'Declare Variable Types
    'Activate this worksheet
    'Declare initial values for TotalVolume and SumTableRow
    'Create LastRow Variable
    'Set Up Summary Table
    'Set Up Change Table
    
    'Declare Variable Types
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim i As Double
    Dim SumTableRow As Double
    Dim ClosingValue As Double
    Dim OpeningValue As Double
    Dim YearlyChange As Double
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatTot As Double
    Dim GreatDecIndex As Integer
    Dim GreatIncIndex As Integer
    Dim GreatTotIndex As Integer
    Dim GreatIncName As String
    Dim GreatDecName As String
    Dim GreatTotName As String
    Dim ws As Worksheet
    Dim PercentChange As Variant

    
    'Activate this WS
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
    'Declare initial values for TotalVolume & SumTableRow
    TotalVolume = 0
    SumTableRow = 2
    
    'Create LastRow variable
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set Up Summary Table
    Range("I1:P1").Font.Bold = True
    Range("I" & 1).Value = "Tickers"
    Range("J" & 1).Value = "Total Volume"
    Range("K" & 1).Value = "Yearly Change"
    Range("L" & 1).Value = "Percent Change"
    
    'Set Up Change Table
    Range("N2:N5").Font.Bold = True
    Cells(1, 15).Value = "Tickers"
    Cells(1, 16).Value = "Values"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"

    
    ' LOOP 1: CALCULATE TOTAL VOLUME, YEARLY CHANGE, & % CHANGE
    '         OUTPUT THESE VALUES AND TICKER NAMES TO SUMMARY TABLE
    '-------------------------------------------------------------------------------
    
    'First row in the loop
    Ticker = Cells(2, 1).Value
    Cells(2, 9).Value = Ticker
    OpeningValue = Cells(2, 3).Value
    TotalVolume = Cells(2, 7).Value
    
    
    'Start for loop
    For i = 3 To LastRow
    
        'Starting points for ticker loop
        Ticker = Cells(i, 1).Value
        PreviousTicker = Cells(i - 1, 1).Value
        
        'If the ticker is different than the one that came before it:
        'Create a new row in the summary table and publish the previous Total Volume
        'and Ticker Name in the summary table. Reset total volume. Excluding first
        'iteration as it was done before the loop
        If Ticker <> Cells(i - 1, 1).Value And i <> 2 Then
            
            'If this ticker is different than the one before it and unique
            If Ticker <> "" Then
                Cells(SumTableRow, 9).Value = PreviousTicker
                Cells(SumTableRow, 10).Value = TotalVolume
                SumTableRow = SumTableRow + 1
                TotalVolume = 0
            End If
            
            'Get opening value for the new ticker row
            OpeningValue = Cells(i, 3).Value
            
        Else
            'If ticker is the same as the last one, add its volume to VolumeTotal
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            'If Ticker is the same as the ticker before it but different than the
            'next ticker, then we know it will be the last day and we can grab
            'our closing value
            If Ticker <> Cells(i + 1, 1).Value Then
                ClosingValue = Cells(i, 6).Value
            End If
        End If
        
        'Calculate Yearly Change and % Change and add to summary table
        'Change PercentChange format to 0.00%
        If Ticker <> "" Then
            Cells(SumTableRow, 11).Value = ClosingValue - OpeningValue
            PercentChange = (Cells(SumTableRow, 11).Value / OpeningValue)
            Cells(SumTableRow, 12).Value = FormatPercent(PercentChange)
        End If
    
    Next i
    
    'Last row in the loop
    Ticker = Cells(LastRow, 1).Value
    Cells(SumTableRow, 9).Value = Ticker
    Cells(SumTableRow, 10).Value = TotalVolume
    Cells(SumTableRow, 11).Value = Cells(LastRow, 6).Value - OpeningValue
    
    
    'LOOP 2: CONDITIONAL FORMATTING
    '-------------------------------------------------------------------------------
    
    'Conditional Formatting to make positive change green, neg change red
    'and no change yellow
    For i = 2 To LastRow
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.Color = vbGreen
            
        ElseIf Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.Color = vbRed
            
        Else
            Cells(i, 11).Interior.Color = vbYellow
            
        End If
        
    Next i
    
    
    'CHANGE TABLE OUTPUT
    '-------------------------------------------------------------------------------
    
    'Find Greatest Increase Name and Value
    GreatInc = WorksheetFunction.Max(Range("L:L"))
    GreatIncIndex = WorksheetFunction.Match(GreatInc, Range("L:L"), 0)
    GreatIncName = Cells(GreatIncIndex, "I").Value
    
    'List to summary table
    Cells(2, 16).Value = FormatPercent(GreatInc)
    Cells(2, 15).Value = GreatIncName
    
    'Find Greatest Decrease
    GreatDec = WorksheetFunction.Min(Range("L:L"))
    GreatDecIndex = WorksheetFunction.Match(GreatDec, Range("L:L"), 0)
    GreatDecName = Cells(GreatDecIndex, "I").Value
    
    'List to summary table
    Cells(3, 16).Value = FormatPercent(GreatDec)
    Cells(3, 15).Value = GreatDecName
    
    'Find Greatest Total
    GreatTot = WorksheetFunction.Max(Range("J:J"))
    GreatTotIndex = WorksheetFunction.Match(GreatTot, Range("J:J"), 0)
    GreatTotName = Cells(GreatTotIndex, "I").Value
    
    'List to summary table
    Cells(4, 16).Value = GreatTot
    Cells(4, 15).Value = GreatTotName
    
    
    'ADJUST VOLUMN WIDTHS FOR BETTER VISUALS
    '--------------------------------------------------------------------------------
    
    'Autofit column width
    Cells.EntireColumn.AutoFit

    'MOVE TO NEXT WORKSHEET AND END SCRIPT
    '--------------------------------------------------------------------------------

    Next ws
            
End Sub








