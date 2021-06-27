Attribute VB_Name = "Module1"
Sub TickerAnalysis()

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
    

Dim ws As Worksheet


    'Create Loop in Worksheet'
    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Value"
        
    Next ws


End Sub

