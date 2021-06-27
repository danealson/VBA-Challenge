Attribute VB_Name = "Module1"
Sub TickerAnalysis()

Dim ws As Worksheet


    'Create Loop in Worksheet'
    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Value"
        
    Next ws


End Sub

