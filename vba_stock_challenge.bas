Attribute VB_Name = "Module1"
Sub TickerAnalysis()

Dim ws As Worksheet

    'Create Loop in Worksheet'
    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Set Variables
        Dim ticker_sym As String
        
        Dim total_vol As Double
        
        Dim rowcount As Long
        rowcount = 2
        
        Dim yearopen As Double
        yearopen = 0
        
        Dim yearclose As Double
        yearclose = 0
        
        Dim year_change As Double
        yearchange = 0
        
        Dim percentchange As Double
        percentchange = 0
                
        Dim greatpercentincr As Double
        Dim greatpercentincrtick As String
        Dim greatpercentdecr As Double
        Dim greatpercentdecrtick As String
        Dim greatstockvol As Double
        Dim greatstockvoltick As String
            
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through ticker symbols
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                yearopen = ws.Cells(i, 3).Value
                
            End If
            
        'total up volume for each row to determine total stock
        
        total_vol = total_vol + ws.Cells(i, 7)
        
        'Conditional to determine if Ticker Symbol is changing
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'Move Ticker Symbol
            ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
            
            'Stock Volume
            ws.Cells(rowcount, 12).Value = total_vol
            
            'End year price
            yearclose = ws.Cells(i, 6).Value
            
            'Calculate Price Change
            yearchange = yearclose - yearopen
            ws.Cells(rowcount, 10).Value = yearchange
            
            'Conditional to see positive or negative changes
            
            If yearchange >= 0 Then
                ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                
            Else
                ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                
            End If
            
            'Percentchange for each year and move to table
            
            'If no change
            If yearopen = 0 And yearclose = 0 Then
            
                percentchange = 0
                ws.Cells(rowcount, 11).Value = percentchange
                ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                
            'If new stock -- no percent change will occur
            ElseIf yearopen = 0 Then
            
                Dim percentchangen As String
                percentchangen = "New Stock"
                ws.Cells(rowcount, 11).Value = percentchangen
                
            Else
            
                percentchange = yearclose / yearopen
                ws.Cells(rowcount, 11).Value = percentchange
                ws.Cells(rowcount, 11).NumberFormat = "0.00%"
            
            End If
            
            rowcount = rowcount + 1
            
            total_vol = 0
            yearopen = 0
            yearclose = 0
            yearchange = 0
            percentchange = 0
            
            End If
        
     Next i
            
 
    
    
                
    Next ws


End Sub

