Attribute VB_Name = "Module1"

Sub stock_data()
    
    Dim ws As Worksheet
    
    'loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        '-----------------------------------------------------------------------
        ' Ticker, qtr change, percent change, & total stock volume columns
        '-----------------------------------------------------------------------
        
        'define var to hold ticker row num
        Dim tickerRow As Double
        tickerRow = 2
        
        'define vars to hold ticker name, open price, close price, qtr change,
        'percent change, total stock volume
        Dim openprice As Double
        Dim closeprice As Double
        Dim qtrchange As Double
        Dim percentchange As Double
        Dim totStockVol As Double
        openprice = ws.Cells(2, 3).Value
        closeprice = 0
        qtrchange = 0
        percentchange = 0
        totStockVol = 0
        
        'headers
        Dim ticker As String
        ticker = "Ticker"
        
        ws.Cells(1, 9).Value = ticker
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'last row of tickers, column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        
        
        'loop through column A, check if next row is the same ticker
        For i = 2 To lastRow
        
            'If next row is not the same ticker...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'record ticker name in column I
                ws.Cells(tickerRow, 9).Value = ws.Cells(i, 1).Value
                
                'grab close price
                closeprice = ws.Cells(i, 6).Value
                
                'calculate & save quarter change, record in column J
                qtrchange = closeprice - openprice
                ws.Cells(tickerRow, 10).Value = qtrchange
                
                'calculate & save percent change, record in column K
                percentchange = (qtrchange / openprice)
                ws.Cells(tickerRow, 11).Value = percentchange
                
                'add total stock volume, record in column L
                totStockVol = totStockVol + ws.Cells(i, 7).Value
                ws.Cells(tickerRow, 12).Value = totStockVol
                
                'grab  next open price
                openprice = ws.Cells(i + 1, 3).Value
                
                'reset total stock volume
                totStockVol = 0
                
                ' + 1 to ticker row
                tickerRow = tickerRow + 1
                
            'If next row is the same ticker...
            Else
                'add total stock volume
                totStockVol = totStockVol + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        
        '-----------------------------------------------------------------------
        ' Greatest % increase, greatest % decrease, greatest total stock volume
        '-----------------------------------------------------------------------
        
        'headers & labels
        Dim great_perc_inc As String
        Dim great_perc_dec As String
        Dim great_tot_vol As String
        Dim val As String
        great_perc_inc = "Greatest Percent Increase"
        great_perc_dec = "Greatest Percent Decrease"
        great_tot_vol = "Greatest Total Volume"
        val = "Value"
        
        'define vars
        Dim greatestInc As Double
        Dim greatestDec As Double
        Dim greatIncTicker As String
        Dim greatDecTicker As String
        Dim greatTotStockVol As Double
        Dim greatTSVTicker As String
        greatestInc = 0
        greatestDec = 0
        greatTotStockVol = 0
        
        'last row of percent change, column I
        lastRow2 = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        'last row of total stock volume, column L
        lastRow3 = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
        
        
        
        'loop through percent changes, column K
        For i = 2 To lastRow2
        
            'check if num is > greatest % increase
            If ws.Cells(i, 11).Value > greatestInc Then
                
                'record % value & ticker
                greatestInc = ws.Cells(i, 11).Value
                greatIncTicker = ws.Cells(i, 9).Value
            
            'check if num is < greatest % decrease
            ElseIf ws.Cells(i, 11).Value < greatestDec Then
            
                'record % value & ticker
                greatestDec = ws.Cells(i, 11).Value
                greatDecTicker = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        'loop through total stock volume
        For i = 2 To lastRow3
            
            'check if greater
            If ws.Cells(i, 12).Value > greatTotStockVol Then
                
                'record greatest total stock volume & ticker
                greatTotStockVol = ws.Cells(i, 12).Value
                greatTSVTicker = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        
        
        'display headers & labels
        ws.Cells(2, 15).Value = great_perc_inc
        ws.Cells(3, 15).Value = great_perc_dec
        ws.Cells(4, 15).Value = great_tot_vol
        ws.Cells(1, 16).Value = ticker
        ws.Cells(1, 17).Value = val
        
        'display tickers
        ws.Cells(2, 16).Value = greatIncTicker
        ws.Cells(3, 16).Value = greatDecTicker
        ws.Cells(4, 16).Value = greatTSVTicker
        
        'display greatest % inc, dec, tot vol of tickers
        ws.Cells(2, 17).Value = greatestInc
        ws.Cells(3, 17).Value = greatestDec
        ws.Cells(4, 17).Value = greatTotStockVol
        
        
        
        '-----------------------------------------------------------------------
        ' Conditional formatting
        '-----------------------------------------------------------------------
        
        'define vars
        Dim rng As Range
        Dim greencell As FormatCondition
        Dim redcell As FormatCondition
        
        'last row of qtrly change
        lastRow4 = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        lastRow5 = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        
        'define conditional formatting quarterly change, column J
        Set rng = ws.Range("J2:J" & lastRow4)
        Set rng2 = ws.Range("K2:K" & lastRow5)
        Set rng3 = ws.Range("Q2:Q3")
        
        'Clear conditional formatting
        rng.FormatConditions.Delete
        
        'Color cells with negative values red
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0)
        End With
        
        'Color positive values green
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0)
        End With
        
        'Format percents to two decimal places & greatest total volume
        rng2.NumberFormat = "0.00%"
        rng3.NumberFormat = "0.00%"
        Set cell = ws.Range("Q4")
        cell.NumberFormat = "0.00E+00"
        
    Next ws
        
End Sub

