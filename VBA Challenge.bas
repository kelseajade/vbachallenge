Attribute VB_Name = "Module1"

Sub StockMarket():
Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim VolumeStock As Double
Dim PercentChange As Double
Dim WS As Worksheet

For Each WS In Worksheets

WS.Range("I1").Value = "Ticker"

WS.Range("J1").Value = "Yearly Change"

WS.Range("K1").Value = "Percent Change"

WS.Range("L1").Value = "Total Stock Volume"


' create for loop to find ticker, year open and year close
EndRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
j = 2
TotalStockVolume = 0

OpenPrice = WS.Cells(2, 3).Value
    For i = 2 To EndRow
    TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
    PercentChange = YearlyChange / OpenPrice
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    tickerName = WS.Cells(i, 1).Value
    ClosePrice = WS.Cells(i, 6).Value
    
    YearlyChange = ClosePrice - OpenPrice
    OpenPrice = WS.Cells(i + 1, 3).Value
' name of the cell = variable in the cell
    WS.Cells(j, 9).Value = tickerName
    WS.Cells(j, 10).Value = YearlyChange
    WS.Cells(j, 11).Value = PercentChange
    WS.Cells(j, 12).Value = TotalStockVolume
    j = j + 1
    TotalStockVolume = 0
    
    End If
    
    Next i
  

For i = 2 To EndRow


'Adding cell formatting
    If WS.Cells(i, 10).Value >= 0 Then
       WS.Cells(i, 10).Interior.ColorIndex = 4
    Else
        WS.Cells(i, 10).Interior.ColorIndex = 3
        
    End If
Next i

'Greatest increase, decrease, and total volume
 Dim percent_max As Double
      percent_max = 0
  Dim percent_min As Double
      percent_min = 0

For i = 2 To EndRow

'Add Conditional
    If percent_max < WS.Cells(i, 11).Value Then
        percent_max = WS.Cells(i, 11).Value
        WS.Cells(2, 17).Value = percent_max
        WS.Cells(2, 17).Style = "Percent"
        WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
    ElseIf percent_min > WS.Cells(i, 11).Value Then
        percent_min = WS.Cells(i, 11).Value
        WS.Cells(3, 17).Value = percent_min
        WS.Cells(3, 17).Style = "Percent"
        WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
    End If
Next i

  Dim TotalStockVolumeRow As Long
      TotalStockVolumeRow = WS.Cells(Rows.Count, 12).End(xlUp).Row
  Dim TotalStockVolumeRowMax As Double
      TotalStockVolumeRowMax = 0

 
 For i = 2 To Total

'Adding Conditional for greatest total volume
    If TotalStockVolumeRowMax < WS.Cells(i, 12).Value Then
       TotalStockVolumeRowMax = WS.Cells(i, 12).Value
       WS.Cells(4, 17).Value = TotalStockVolumeRowMax
       WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
       
    End If
Next i
    
Next WS
    
End Sub
