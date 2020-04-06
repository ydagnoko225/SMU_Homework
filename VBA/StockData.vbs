Sub stockData()

Dim i As Double
Dim lastRow As Long
Dim column As Long
Dim summaryTableIndex As Long
Dim runningVolume As Double
Dim start As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim ws As Worksheet
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim TotalVolume As Double
'Dim myRange As Range



For Each ws In Worksheets

Set K = Range("K2:K" & Rows.Count)

'LAST ROW
lastRow = Range("A" & Rows.Count).End(xlUp).Row
column = 1

start = 2

'this is how we'll keep track of G
summaryTableIndex = 2

'Sum
runningVolume = 0


   'write cells
    'Column headers
    ws.Range("I" & summaryTableIndex - 1).Value = "Ticker"
    ws.Range("J" & summaryTableIndex - 1).Value = "Yearly Change"
    ws.Range("K" & summaryTableIndex - 1).Value = "Percent Change"
    ws.Range("L" & summaryTableIndex - 1).Value = "Total Stock Volume"
    ws.Range("N" & summaryTableIndex + 2).Value = "Greatest % increase"
    ws.Range("N" & summaryTableIndex + 3).Value = "Greatest % decrease"
    ws.Range("N" & summaryTableIndex + 4).Value = "Greatest Total Volume"

' loop until the last row
For i = 2 To lastRow

'if the index of i is different to the index of the last row(aka running higher)
If i <> lastRow Then

       If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then

       runningVolume = runningVolume + Cells(i, 7).Value
       YearlyChange = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
       
       
       If (Cells(start, 3) <> 0) Then
       PercentChange = Round((YearlyChange / Cells(start, 3) * 100), 2)
       End If
       

         ws.Cells(summaryTableIndex, 12) = runningVolume
         ws.Cells(summaryTableIndex, 9) = ws.Cells(i, column).Value
         ws.Range("J" & summaryTableIndex).Value = YearlyChange
         ws.Range("K" & summaryTableIndex).Value = (CStr(PercentChange) & "%")
         
       'color
       If (YearlyChange > 0) Then
       ws.Range("J" & summaryTableIndex).Interior.ColorIndex = 4
       ElseIf (YearlyChange < 0) Then
       ws.Range("J" & summaryTableIndex).Interior.ColorIndex = 3
       End If
                     
        'Reset the the running total to 0
          runningVolume = 0
         'manually increment G
          summaryTableIndex = summaryTableIndex + 1
          start = i + 1
          
 
      Else
     runningVolume = runningVolume + ws.Cells(i, 7).Value

      End If


Else
'add to the running volume
 runningVolume = runningVolume + ws.Cells(i, 7).Value
 ws.Range("J" & summaryTableIndex).Value = YearlyChange
 ws.Range("K" & summaryTableIndex).Value = PercentChange

'write the cells
 ws.Cells(summaryTableIndex, 12) = runningVolume
 ws.Cells(summaryTableIndex, 9) = ws.Cells(i, column).Value
 
'RESETS
'Reset the the running total to 0
 runningVolume = 0

' manually increment G
 summaryTableIndex = summaryTableIndex + 1
 start = i + 1

End If

Next i

       'GreatestPercentIncrease = Application.WorksheetFunction.Max(K)
       'GreatestPercentDecrease = Application.WorksheetFunction.Min(K)
       'TotalVolume = Application.WorksheetFunction.Sum(ws.Range("L:L"))
        
       'ws.Range("O" & summaryTableIndex + 2) = GreatestPercentIncrease
       'ws.Range("O" & summaryTableIndex + 3) = GreatestPercentIncrease
       'ws.Range("O" & summaryTableIndex + 4) = TotalVolume
       

Next ws

End Sub




