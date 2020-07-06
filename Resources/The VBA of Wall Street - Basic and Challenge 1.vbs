Sub SortWorksheet()

' Sort the current worksheet to make sure it is sorted by ticker then date ascending
' Code help for this was taken from https://trumpexcel.com/sort-data-vba/ and https://stackoverflow.com/questions/52619676/sort-multiple-columns-excel-vba
 
 Dim sheet As Worksheet
 Set sheet = ActiveSheet
 
' get the last row
last_row = Cells(Rows.Count, 1).End(xlUp).Row

With sheet.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange ActiveSheet.Range("A1:G" & last_row)
     .Header = xlYes
     .Apply
End With

End Sub

Sub ticker_totals()

Dim ticker As String
Dim last_row As Long
Dim last_totals_row As Long
Dim NewTotalsRow As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChangePrice As Double
Dim StockVolume As Variant
Dim sheet As Worksheet

' run the SortWorksheet subroutine to make sure all the records in the worksheet are sorted correctly
Call SortWorksheet

Set sheet = ActiveSheet

' Add the column headers for the totals section
Range("J1").Value = "Ticker Symbol"
Range("K1").Value = "Yearly Change ($)"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

' get the last row
last_row = Cells(Rows.Count, 1).End(xlUp).Row

' set the first totals row
NewTotalsRow = 2

' get first ticker's opening price
OpeningPrice = Cells(2, 3).Value

' get each ticker symbol. First row is the header so start at row 2
For i = 2 To last_row
    ' find when on the first row for a ticker
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        ' get the ticker name, its opening value and the opening volume
        ticker = Cells(i, 1).Value
        OpeningPrice = Cells(i, 3).Value
        StockVolume = Cells(i, 7).Value
    
    ' find when the value of the next cell is different than that of the current cell
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' put the current ticker value in the next Totals row
        Cells(NewTotalsRow, 10).Value = ticker 'Cells(i, 1).Value
    
        ' get the closing price for that ticker, calculate the yearly change and the yearly change percent and place those values in the Totals column for that ticker
        ClosingPrice = Cells(i, 6).Value
        YearlyChangePrice = ClosingPrice - OpeningPrice
        Cells(NewTotalsRow, 11).Value = YearlyChangePrice
        ' need to handle divide by zero issue
        If (YearlyChangePrice = 0 Or OpeningPrice = 0) Then
            Cells(NewTotalsRow, 12).Value = 0
        Else
            Cells(NewTotalsRow, 12).Value = YearlyChangePrice / OpeningPrice
        End If
        
        ' get the stock volume total and insert it into the totals section
        Cells(NewTotalsRow, 13).Value = StockVolume + Cells(i, 7).Value
        
        ' get the opening price for the next ticker and the totals row
'        OpeningPrice = Cells(i + 1, 3).Value
        NewTotalsRow = NewTotalsRow + 1
'        StockVolume = Cells(i + 1, 7).Value
    Else
        ' get the running total for stock volume
        StockVolume = StockVolume + Cells(i, 7).Value
    End If
' calculate the yearly change for the current ticker

Next i

' get the last row of the totals section
last_totals_row = NewTotalsRow - 1

' CHALLENGE Section

Dim greatest_per_inc As Double
Dim greatest_per_dec As Double
Dim greatest_tot_vol As Variant
Dim max_ticker_name As String
Dim min_ticker_name As String
Dim max_vol_ticker As String

' Add the column headers for the Challenge section
Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"
Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

' get stock with the greatest % increase, greatest % decrease, and greatest total volume
' code help for this was taken from https://stackoverflow.com/questions/51977446/vba-find-highest-value-in-a-column-c-and-return-its-value-and-the-adjacent-ce
For x = 2 To last_totals_row
    If sheet.Cells(x, 12) > greatest_per_inc Then
       greatest_per_inc = sheet.Cells(x, 12)
       max_ticker_name = sheet.Cells(x, 10)
    End If

Next x

For n = 2 To last_totals_row
    If sheet.Cells(n, 12) < greatest_per_dec Then
       greatest_per_dec = sheet.Cells(n, 12)
       min_ticker_name = sheet.Cells(n, 10)
    End If

Next n

For v = 2 To last_totals_row
    If sheet.Cells(v, 13) > greatest_tot_vol Then
       greatest_tot_vol = sheet.Cells(v, 13)
       max_vol_ticker = sheet.Cells(v, 10)
    End If

Next v

Cells(2, 18).Value = greatest_per_inc
Cells(2, 17).Value = max_ticker_name
Cells(3, 18).Value = greatest_per_dec
Cells(3, 17).Value = min_ticker_name
Cells(4, 18).Value = greatest_tot_vol
Cells(4, 17).Value = max_vol_ticker


' Conditional formatting applied to the YearlyChange column -- positive change = green, negative change = red
For f = 2 To (last_totals_row)
    If Cells(f, 11).Value < 0 Then
        Cells(f, 11).Interior.ColorIndex = 3
    ElseIf Cells(f, 11).Value > 0 Then
        Cells(f, 11).Interior.ColorIndex = 4
    Else
        Cells(f, 11).Interior.ColorIndex = 2
    End If

Next f

' Add percent formatting to the Percent Changed column
Range("L2:L" & last_totals_row).NumberFormat = "0.00%"

' Add percent formatting to the Challenge Greatest % Increase and Greatest % Decrease columns
Range("R2:R3").NumberFormat = "0.00%"
    
' Autofit to display data
sheet.Columns("J:M").AutoFit
'For Each ws In Worksheets
'    ws.Columns("J:M").AutoFit
'Next ws

' Autofit to Challenge data
sheet.Columns("P:R").AutoFit
'For Each ws In Worksheets
'    ws.Columns("P:R").AutoFit
'Next ws

End Sub




