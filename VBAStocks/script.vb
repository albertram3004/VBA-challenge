Sub StockYearMarketdata()

Dim Year_Change As Double
Dim Percent_Change As Double
Dim Volume_Total As Double
Dim Open_Amount As Double
Dim Close_Amount As Double
Dim Ticker_Value As String
Dim LastRow As Long
Dim PercentLast As Integer
Dim OpenCol As Integer
Dim CloseCol As Integer
Dim VolumeCol As Integer
Dim FirstRecord As Boolean
Dim Summary_Col As Integer
Dim Summary_Row As Integer
' Loop sheets & Populate Summary Sections
   For Each ws In Worksheets
  'Set Dim of ws and initialize variales
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  OpenCol = 3
  CloseCol = 6
  VolumeCol = 7
  Volume_Total = 0
  FirstRecord = True
  'Set Start Position and Headers of Summary Section
  Summary_Col = 9
  Summary_Row = 2
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  'Loop ws
  For i = 2 To LastRow
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
          Volume_Total = Volume_Total + ws.Cells(i, VolumeCol).Value
          Close_Amount = ws.Cells(i, CloseCol).Value
          Ticker_Value = ws.Cells(i, 1).Value
          Year_Change = Close_Amount - Open_Amount
          If Open_Amount = 0 Then
              Percent_Change = 0
          Else
              Percent_Change = Round((Year_Change / Open_Amount), 4)
          End If
          'Summary 1 Out
          ws.Cells(Summary_Row, Summary_Col).Value = Ticker_Value
          ws.Cells(Summary_Row, Summary_Col + 1).Value = Year_Change
          ws.Cells(Summary_Row, Summary_Col + 2).Value = Percent_Change
          ws.Cells(Summary_Row, Summary_Col + 3).Value = Volume_Total
          If Year_Change < 0 Then
              ws.Cells(Summary_Row, Summary_Col + 1).Interior.ColorIndex = 3
          Else
              ws.Cells(Summary_Row, Summary_Col + 1).Interior.ColorIndex = 4
          End If
          'Reset for next ticker type
          Summary_Row = Summary_Row + 1
          FirstRecord = True
          Volume_Total = 0
      Else
          If FirstRecord Then
              Open_Amount = ws.Cells(i, OpenCol).Value
              FirstRecord = False
          End If
          Volume_Total = Volume_Total + ws.Cells(i, VolumeCol).Value
      End If
  Next i
  'Summary 2 Out & Formatting
  PercentLast = ws.Cells(Rows.Count, Summary_Col + 2).End(xlUp).Row
  ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & PercentLast))
  ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & PercentLast), WorksheetFunction.Match _
                      (ws.Cells(2, 17).Value, ws.Range("K2:K" & PercentLast), 0))
  ws.Cells(2, 17).NumberFormat = "0.00%"
  ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & PercentLast))
  ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & PercentLast), WorksheetFunction.Match _
                      (ws.Cells(3, 17).Value, ws.Range("K2:K" & PercentLast), 0))
  ws.Cells(3, 17).NumberFormat = "0.00%"
  ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & PercentLast))
  ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & PercentLast), WorksheetFunction.Match _
                      (ws.Cells(4, 17).Value, ws.Range("L2:L" & PercentLast), 0))
  ws.Range("K2:K" & PercentLast).NumberFormat = "0.00%"
  ws.Columns("A:Q").AutoFit
Next ws
End Sub

