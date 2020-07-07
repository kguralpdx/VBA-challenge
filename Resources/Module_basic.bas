Attribute VB_Name = "Module1"
Sub SortWorksheet()

' Sort the current worksheet to make sure it is sorted by ticker then date ascending
' Code help for this was taken from https://trumpexcel.com/sort-data-vba/ and https://stackoverflow.com/questions/52619676/sort-multiple-columns-excel-vba
 
' get the last row
last_row = Cells(Rows.Count, 1).End(xlUp).Row

With ActiveSheet.Sort
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

' run the SortWorksheet subroutine to make sure all the records in the worksheet are sorts correctly
Call SortWorksheet

' Add the column headers for the totals section
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
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

Range("L2:L" & last_totals_row).NumberFormat = "0.00%"
    
'    ' Autofit to display data
'For Each ws In Worksheets
'    ws.Columns("J:M").AutoFit
'Next ws

End Sub
