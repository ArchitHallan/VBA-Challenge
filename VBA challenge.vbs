Option Explicit
Dim greatestPIncreaseTicker As String
Dim greatestPIncrease As Double

Dim greatestPDecreaseTicker As String
Dim greatestPDecrease As Double

Dim greatestTotVolTicker As String
Dim greatestTotVol As Double

Public Sub PrintResult()
Dim ws As Worksheet
 For Each ws In Worksheets
Worksheets(ws.Name).Activate
Dim ticker As String
Dim rng As Range
Dim nxt As Boolean
Dim openstock As Double
Dim closestock As Double
Dim countRow As Integer
Dim totalStockVol As Double
Dim irow As Double
Dim icol As Double
countRow = 1
totalStockVol = 0
Set rng = Range("A2:" + Range("A2").End(xlToRight).End(xlDown).Address)
Dim data() As Variant
data = rng
ticker = data(1, 1)
openstock = data(1, 3)
For irow = 1 To UBound(data)
totalStockVol = totalStockVol + data(irow, 7)
For icol = 1 To 7
If data(irow, 1) <> ticker Or (irow = UBound(data) And icol = 7) Then
closestock = data(irow - 1, 6)
countRow = countRow + 1
InsertSummary countRow, ticker, (closestock - openstock), openstock, (totalStockVol - data(irow, 7))
ticker = data(irow, 1)
openstock = data(irow, 3)
totalStockVol = data(irow, 7)
End If

Next icol
Next irow
Next
MsgBox "Work Completed"
End Sub



Public Sub InsertSummary(row As Integer, ticker As String, yearlychange As Double, openingstock As Double, totalStockVol As Double)


If (row = 2) Then
Columns("J").ColumnWidth = 12
Columns("K").ColumnWidth = 14
Columns("L").ColumnWidth = 14
Columns("M").ColumnWidth = 17
Columns("O").ColumnWidth = 20
Columns("P").ColumnWidth = 8
Columns("Q").ColumnWidth = 20
Columns("N").ColumnWidth = 4
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
greatestPIncrease = (yearlychange / openingstock)
greatestPDecrease = (yearlychange / openingstock)
greatestTotVol = totalStockVol
End If
Range("J" & row).Value = ticker
If yearlychange < 0 Then
Range("K" & row).Interior.ColorIndex = 3
Else
Range("K" & row).Interior.ColorIndex = 4
End If

Range("K" & row).Value = yearlychange
Range("L" & row).NumberFormat = "0.00%"
Range("L" & row).Value = yearlychange / openingstock
Range("M" & row).Value = totalStockVol

If (yearlychange / openingstock) >= greatestPIncrease Then
greatestPIncreaseTicker = ticker
greatestPIncrease = (yearlychange / openingstock)
Range("Q2").NumberFormat = "0.00%"
Range("Q2").Value = greatestPIncrease
Range("P2").Value = greatestPIncreaseTicker
Range("O2").Value = "Greatest % Increase"
End If
If (yearlychange / openingstock) <= greatestPDecrease Then
greatestPDecreaseTicker = ticker
greatestPDecrease = (yearlychange / openingstock)
Range("Q3").NumberFormat = "0.00%"
 Range("Q3").Value = greatestPDecrease
 Range("P3").Value = greatestPDecreaseTicker
  Range("O3").Value = "Greatest % Decrease"
End If

If totalStockVol >= greatestTotVol Then
greatestTotVolTicker = ticker
greatestTotVol = totalStockVol
Range("Q4").Value = greatestTotVol
Range("P4").Value = greatestTotVolTicker
 Range("O4").Value = "Greatest Total Volumee"
End If
End Sub




