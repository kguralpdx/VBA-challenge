Attribute VB_Name = "Module1"
Sub ticker_totals()

For Each ws In Worksheets

Dim ticker As String
Dim last_row As Long
Dim last_totals_row As Long
Dim NewTotalsRow As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChangePrice As Double
Dim StockVolume As Variant
Dim greatest_per_inc As Double
Dim greatest_per_dec As Double
Dim greatest_tot_vol As Variant
Dim max_ticker_name As String
Dim min_ticker_name As String
Dim max_vol_ticker As String
'Dim ws As Worksheet


' make the current worksheet the active one
'    ws.Activate
    
    ' get the last row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Sort the current worksheet to make sure it is sorted by ticker then date ascending
    ' Code help for this was taken from https://trumpexcel.com/sort-data-vba/ and https://stackoverflow.com/questions/52619676/sort-multiple-columns-excel-vba
    With ws.Sort
         .SortFields.Add Key:=Range("A1"), Order:=xlAscending
         .SortFields.Add Key:=Range("B1"), Order:=xlAscending
         .SetRange ActiveSheet.Range("A1:G" & last_row)
         .Header = xlYes
         .Apply
    End With
    
    ' Add the column headers for the totals section
    ws.Range("J1").Value = "Ticker Symbol"
    ws.Range("K1").Value = "Yearly Change ($)"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    ' set the first totals row
    NewTotalsRow = 2
    
    ' get first ticker's opening price
    OpeningPrice = ws.Cells(2, 3).Value
    
    ' get each ticker symbol. First row is the header so start at row 2
    For i = 2 To last_row
        ' find when on the first row for a ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' get the ticker name, its opening value and the opening volume
            ticker = ws.Cells(i, 1).Value
            OpeningPrice = ws.Cells(i, 3).Value
            StockVolume = ws.Cells(i, 7).Value
        
        ' find when the value of the next cell is different than that of the current cell
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' put the current ticker value in the next Totals row
            ws.Cells(NewTotalsRow, 10).Value = ticker
        
            ' get the closing price for that ticker, calculate the yearly change and the yearly change percent and place those values in the Totals column for that ticker
            ClosingPrice = ws.Cells(i, 6).Value
            YearlyChangePrice = ClosingPrice - OpeningPrice
            ws.Cells(NewTotalsRow, 11).Value = YearlyChangePrice
            ' need to handle divide by zero issue
            If (YearlyChangePrice = 0 Or OpeningPrice = 0) Then
                ws.Cells(NewTotalsRow, 12).Value = 0
            Else
                ws.Cells(NewTotalsRow, 12).Value = YearlyChangePrice / OpeningPrice
            End If
            
            ' get the stock volume total and insert it into the totals section
            ws.Cells(NewTotalsRow, 13).Value = StockVolume + ws.Cells(i, 7).Value
            
            ' get the ototals row for the next ticker
            NewTotalsRow = NewTotalsRow + 1
        Else
            ' get the running total for stock volume
            StockVolume = StockVolume + ws.Cells(i, 7).Value
        End If
    ' calculate the yearly change for the current ticker
    
    Next i
    
    ' get the last row of the totals section
    last_totals_row = NewTotalsRow - 1
    
    
    ' ******CHALLENGE Section*******
    
    ' empty out the totals so that they're populated using data just from the current worksheet, not all of the worksheets
    greatest_per_inc = 0
    max_ticker_name = ""
    greatest_per_dec = 0
    min_ticker_name = ""
    greatest_tot_vol = 0
    max_vol_ticker = ""
    
    
    ' Add the column headers for the Challenge section
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker Symbol"
    ws.Range("R1").Value = "Value"
    
    ' get stock with the greatest % increase, greatest % decrease, and greatest total volume
    ' code help for this was taken from https://stackoverflow.com/questions/51977446/vba-find-highest-value-in-a-column-c-and-return-its-value-and-the-adjacent-ce
    For x = 2 To last_totals_row
        If ws.Cells(x, 12) > greatest_per_inc Then
           greatest_per_inc = ws.Cells(x, 12)
           max_ticker_name = ws.Cells(x, 10)
        End If
    
    Next x
    
    For n = 2 To last_totals_row
        If ws.Cells(n, 12) < greatest_per_dec Then
           greatest_per_dec = ws.Cells(n, 12)
           min_ticker_name = ws.Cells(n, 10)
        End If
    
    Next n
    
    For v = 2 To last_totals_row
        If ws.Cells(v, 13) > greatest_tot_vol Then
           greatest_tot_vol = ws.Cells(v, 13)
           max_vol_ticker = ws.Cells(v, 10)
        End If
    
    Next v
    
    ws.Cells(2, 18).Value = greatest_per_inc
    ws.Cells(2, 17).Value = max_ticker_name
    ws.Cells(3, 18).Value = greatest_per_dec
    ws.Cells(3, 17).Value = min_ticker_name
    ws.Cells(4, 18).Value = greatest_tot_vol
    ws.Cells(4, 17).Value = max_vol_ticker
    
    
    ' Conditional formatting applied to the YearlyChange column -- positive change = green, negative change = red
    For f = 2 To (last_totals_row)
        If ws.Cells(f, 11).Value < 0 Then
            ws.Cells(f, 11).Interior.ColorIndex = 3
        ElseIf ws.Cells(f, 11).Value > 0 Then
            ws.Cells(f, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(f, 11).Interior.ColorIndex = 2
        End If
    
    Next f
    
    ' Add percent formatting to the Percent Changed column
    ws.Range("L2:L" & last_totals_row).NumberFormat = "0.00%"
    
    ' Add percent formatting to the Challenge Greatest % Increase and Greatest % Decrease columns
    ws.Range("R2:R3").NumberFormat = "0.00%"
        
    ' Autofit to display data
    ws.Columns("J:M").AutoFit
    
    ' Autofit to Challenge data
    ws.Columns("P:R").AutoFit

Next ws

End Sub



