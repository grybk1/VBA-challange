Attribute VB_Name = "Module1"
Sub main()
Dim ws As Worksheet



For Each ws In Worksheets
    'StockChecker
    'Bonus
  ws.Activate
  StockChecker
  Bonus
 
Next ws
End Sub


Sub StockChecker()

Dim TickerSymbol As String
Dim NextTickerSymbol As String
Dim TotalStockVolumne As Long
Dim YearlyStartPrice As Double
Dim YearlyEndPrice As Double
Dim YearlyChange As Double
Dim PercentageOfChange As Double
Dim ResultPosition As Integer
Dim sht As Worksheet
Dim lastRow As Long


'Build Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volue"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volue"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"



'initialize variables
Set sht = ActiveSheet
lastRow = sht.Cells(Rows.Count, "A").End(xlUp).Row
TotalStockVolume = 0
YearlyStartPrice = Cells(2, 3).Value
YearlyEndPrice = 0
ResultPosition = 2

'create loop for summation
For i = 2 To lastRow

    TickerSymbol = Cells(i, 1).Value
    NextTickerSymbol = Cells(i + 1, 1).Value

        If TickerSymbol = NextTickerSymbol Then
         TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        Else
         'Accumulate summations
         TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
         YearlyEndPrice = Cells(i, 6).Value
         YearlyChange = YearlyEndPrice - YearlyStartPrice
         If YearlyStartPrice > 0 Then PercentageOfChange = YearlyChange / YearlyStartPrice
         If YearlyChange > 0 Then cellColor = 4 Else cellColor = 3
         
         'Present summations
     
         Cells(ResultPosition, 9).Value = TickerSymbol
         Cells(ResultPosition, 12).Value = TotalStockVolume
         Cells(ResultPosition, 10).Value = YearlyChange
         Cells(ResultPosition, 10).Interior.ColorIndex = cellColor
         Cells(ResultPosition, 11).Value = PercentageOfChange
         Cells(ResultPosition, 11).NumberFormat = ".00%"
         Cells(ResultPosition, 12).NumberFormat = "0000"
         
         'Reinitialize variables
         TickerSymbol = NextTickerSymbol
         YearlyStartPrice = Cells(i + 1, 3).Value
         ResultPosition = ResultPosition + 1
                          
        End If

Next i

End Sub

Sub Bonus()

Dim MaxPercentChange As Double
Dim NextPercentChange As Double
Dim MaxTicker As String
Dim MinPercentChange As Double
Dim MinTicker As String
Dim MaxVolume As Double
Dim NextVolume As Double
Dim MaxVolumeTicker As String


Dim sht As Worksheet
Dim lastRow As Long

'initialize variables
Set sht = ActiveSheet
lastRow = sht.Cells(Rows.Count, 11).End(xlUp).Row - 1
MaxPercentChange = 0
MinPercentChange = 0
MaxVolume = 0

'Loop dataset
For i = 2 To lastRow
    NextPercentChange = Cells(i, 11).Value
    If MaxPercentChange < NextPercentChange Then
        MaxPercentChange = NextPercentChange
        MaxTicker = Cells(i, 9).Value
    End If
    If MinPercentChange > NextPercentChange Then
        MinPercentChange = NextPercentChange
        MinTicker = Cells(i, 9).Value
    End If
    NextVolume = Cells(i, 12).Value
    If MaxVolume < NextVolume Then
    MaxVolume = NextVolume
    MaxVolumeTicker = Cells(i, 9).Value
    End If
    
Next i

'Present Results
Cells(2, 16).Value = MaxTicker
Cells(2, 17).Value = MaxPercentChange
Cells(2, 17).NumberFormat = ".00%"

Cells(3, 16).Value = MinTicker
Cells(3, 17).Value = MinPercentChange
Cells(3, 17).NumberFormat = ".00%"

Cells(4, 16).Value = MaxVolumeTicker
Cells(4, 17).Value = MaxVolume
Cells(4, 17).NumberFormat = "0000"

End Sub

Sub resest()
Dim sht As Worksheet
Dim lastRow As Long
Dim RangeString As String


Set sht = ActiveSheet
lastRow = sht.Cells(Rows.Count, "L").End(xlUp).Row
For i = 1 To lastRow
    Cells(i, 9).Value = ""
    Cells(i, 10).Value = ""
    Cells(i, 11).Value = ""
    Cells(i, 12).Value = ""
    Cells(i, 10).Interior.ColorIndex = 0

Next i


End Sub

