Attribute VB_Name = "Module1"
Sub StockDataAnalysis():

    'Looping through the worksheets
    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        DataTabLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        Dim TickerName As String
    
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
    
        Dim SummaryTableRowNumber As Integer
        SummaryTableRowNumber = 2
    
        Dim PriceClose As Double
        Dim PriceOpen As Double
        PriceClose = -1
        PriceOpen = -1
        
        Dim YearlyChange As Double
        Dim YearlyChangePercent As Double
        Dim YearlyChangePercentString As String
            
        'Looping through the data table rows within the sheet
        'Assumption: our data are already sorted by date (in descending order)
        'Otherwise, we would have to add the conditional for checking the dates or to sort our data table
        'To sort the data table uncomment block below
        
        ''''''''''''
        'ws.Sort.SortFields.Clear
        'ws.Sort.SortFields.Add Key:=Range("A2:A" & DataTabLastRow _
        '    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        'ws.Sort.SortFields.Add Key:=Range("B2:B" & DataTabLastRow _
        '    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        'With ws.Sort
        '    .SetRange Range("A1:G" & DataTabLastRow)
        '    .Header = xlYes
        '    .MatchCase = False
        '    .Orientation = xlTopToBottom
        '    .SortMethod = xlPinYin
        '    .Apply
        'End With
        ''''''''''''
        
        For i = 2 To DataTabLastRow
        
             If PriceOpen = -1 Then
                 PriceOpen = ws.Cells(i, 3).Value
             End If
             
             PriceClose = ws.Cells(i, 6).Value
            
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
                 TickerName = ws.Cells(i, 1).Value
                 ws.Range("I" & SummaryTableRowNumber).Value = TickerName
                 
                 If PriceOpen >= 0 And PriceClose >= 0 Then
                 
                    'Calculating Yearly Change
                    '''''''''
                    YearlyChange = PriceClose - PriceOpen
                    ws.Range("J" & SummaryTableRowNumber).Value = YearlyChange
                    ws.Range("J" & SummaryTableRowNumber).NumberFormat = "#,##0.000000000"
                    '''''''''
                    
                    'Calculating Yearly Change Percent
                    '''''''''
                    If PriceOpen > 0 Then
                       YearlyChangePercent = (PriceClose - PriceOpen) / PriceOpen
                       ws.Range("K" & SummaryTableRowNumber).Value = YearlyChangePercent
                       ws.Range("K" & SummaryTableRowNumber).NumberFormat = "#,##0.00%"
                    ElseIf PriceOpen = 0 And PriceClose = 0 Then
                       YearlyChangePercent = 0
                       ws.Range("K" & SummaryTableRowNumber).Value = YearlyChangePercent
                       ws.Range("K" & SummaryTableRowNumber).NumberFormat = "#,##0.00%"
                    Else:
                       YearlyChangePercentString = "NA"
                       ws.Range("K" & SummaryTableRowNumber).Value = YearlyChangePercentString
                    End If
                    '''''''''
                 
                 End If
                 
                 PriceClose = -1
                 PriceOpen = -1
             
                 'Calculating Total Stock Volume
                 '''''''''
                 TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                 ws.Range("L" & SummaryTableRowNumber).Value = TotalStockVolume
                 '''''''''
                 
                 'Formatting
                 '''''''''
                 If ws.Range("J" & SummaryTableRowNumber).Value < 0 Then
                     ws.Range("J" & SummaryTableRowNumber).Interior.Color = RGB(255, 0, 0)
                 ElseIf ws.Range("J" & SummaryTableRowNumber).Value > 0 Then
                     ws.Range("J" & SummaryTableRowNumber).Interior.Color = RGB(0, 255, 0)
                 End If
                 '''''''''
                
                 SummaryTableRowNumber = SummaryTableRowNumber + 1
             
                 TotalStockVolume = 0
             
             Else
             
                 TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
             
             End If
        
        Next i
        
        'Calculating "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Volume"
        ''''''''''
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
                        
        SummaryTabLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Dim MaxIncrease As Double
        Dim MaxIncreaseTicker As String
        
        MaxIncrease = ws.Cells(2, 11).Value
                
        For k = 2 To SummaryTabLastRow
            If ws.Cells(k, 11) > MaxIncrease And IsNumeric(ws.Cells(k, 11)) Then
                MaxIncrease = ws.Cells(k, 11)
                MaxIncreaseTicker = ws.Cells(k, 9)
            End If
        Next k
        
        ws.Range("P2").Value = MaxIncrease
        ws.Range("P2").NumberFormat = "#,##0.00%"
        ws.Range("O2").Value = MaxIncreaseTicker
        
        Dim MaxDecrease As Double
        Dim MaxDecreaseTicker As String
        
        MaxDecrease = ws.Cells(2, 11).Value
        
        For l = 2 To SummaryTabLastRow
            If ws.Cells(l, 11) < MaxDecrease And IsNumeric(ws.Cells(l, 11)) Then
                MaxDecrease = ws.Cells(l, 11)
                MaxDecreaseTicker = ws.Cells(l, 9)
            End If
        Next l
        
        ws.Range("P3").Value = MaxDecrease
        ws.Range("P3").NumberFormat = "#,##0.00%"
        ws.Range("O3").Value = MaxDecreaseTicker
        
        Dim MaxTotal As Double
        Dim MaxTotalTicker As String
        
        MaxTotal = ws.Cells(2, 12).Value
        
        For m = 2 To SummaryTabLastRow
            If ws.Cells(m, 12) > MaxTotal And IsNumeric(ws.Cells(m, 12)) Then
                MaxTotal = ws.Cells(m, 12)
                MaxTotalTicker = ws.Cells(m, 9)
            End If
        Next m
        
        ws.Range("P4").Value = MaxTotal
        ws.Range("O4").Value = MaxTotalTicker
        ''''''''''
        
        ws.Columns("A:P").EntireColumn.AutoFit
                                   
    Next ws

End Sub

