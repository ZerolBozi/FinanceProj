Attribute VB_Name = "CalculateIndicator"
Option Explicit

Sub CalculateVO(ws As Worksheet, stdLength As Integer)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim VOCol As Long
    Dim volume As Variant
    Dim volumeStd As Double
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < stdLength Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    VOCol = lastCol + 1
    
    volume = ws.Range("G2:G" & lastRow).value
    
    ws.Cells(1, VOCol).value = "VO"
    
    ws.Range(ws.Cells(2, 7), ws.Cells(lastRow, 7)).Copy ws.Cells(2, VOCol)
    
    For i = stdLength + 1 To lastRow
        volumeStd = Application.WorksheetFunction.StDev(ws.Range(ws.Cells(i - stdLength + 1, 7), ws.Cells(i, 7)))
        
        If ws.Cells(i, VOCol).value > volumeStd * 5 Then
            ws.Cells(i, VOCol).Interior.ColorIndex = 26
        ElseIf ws.Cells(i, VOCol).value > volumeStd * 3.5 Then
            ws.Cells(i, VOCol).Interior.ColorIndex = 44
        ElseIf ws.Cells(i, VOCol).value > volumeStd * 2 Then
            ws.Cells(i, VOCol).Interior.ColorIndex = 27
        End If
    Next i
    
End Sub

Sub CalculateBB(ws As Worksheet, period As Integer, numStdDev As Double)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim bbCol As Long, upperCol As Long, lowerCol As Long
    Dim i As Long
    Dim closePrices As Variant
    Dim movingAvg As Double
    Dim stdDev As Double

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < period Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    bbCol = lastCol + 1
    upperCol = lastCol + 2
    lowerCol = lastCol + 3
    
    closePrices = ws.Range("F2:F" & lastRow).value
    
    ws.Cells(1, bbCol).value = "BB_Mid"
    ws.Cells(1, upperCol).value = "BB_Upper"
    ws.Cells(1, lowerCol).value = "BB_Lower"
    
    For i = period To lastRow - 1
        movingAvg = Application.WorksheetFunction.Average(ws.Range(ws.Cells(i - period + 2, 6), ws.Cells(i + 1, 6)))
        stdDev = Application.WorksheetFunction.StDev(ws.Range(ws.Cells(i - period + 2, 6), ws.Cells(i + 1, 6)))
        
        ws.Cells(i + 1, bbCol).value = movingAvg
        ws.Cells(i + 1, upperCol).value = movingAvg + numStdDev * stdDev
        ws.Cells(i + 1, lowerCol).value = movingAvg - numStdDev * stdDev
    Next i
    
    For i = 2 To period
        ws.Cells(i, bbCol).value = ""
        ws.Cells(i, upperCol).value = ""
        ws.Cells(i, lowerCol).value = ""
    Next i
End Sub


Sub CalculateMACD(ws As Worksheet, shortPeriod As Integer, longPeriod As Integer, signalPeriod As Integer)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim closePrices As Variant
    Dim macdCol As Long, signalCol As Long, histCol As Long
    Dim shortEma As Variant, longEma As Variant
    Dim i As Long
    Dim kShort As Double, kLong As Double, kSignal As Double
    Dim macd As Double, signal As Double

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < longPeriod Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    macdCol = lastCol + 1
    signalCol = lastCol + 2
    histCol = lastCol + 3
    
    closePrices = ws.Range("F2:F" & lastRow).value
    
    ws.Cells(1, macdCol).value = "MACD"
    ws.Cells(1, signalCol).value = "Signal"
    ws.Cells(1, histCol).value = "Hist"
    
    ReDim shortEma(1 To lastRow - 1)
    ReDim longEma(1 To lastRow - 1)
    
    kShort = 2 / (shortPeriod + 1)
    kLong = 2 / (longPeriod + 1)
    kSignal = 2 / (signalPeriod + 1)
    
    shortEma(shortPeriod) = Application.WorksheetFunction.Average(ws.Range(ws.Cells(2, 6), ws.Cells(shortPeriod + 1, 6)))
    longEma(longPeriod) = Application.WorksheetFunction.Average(ws.Range(ws.Cells(2, 6), ws.Cells(longPeriod + 1, 6)))
    
    For i = shortPeriod + 1 To lastRow - 1
        shortEma(i) = closePrices(i, 1) * kShort + shortEma(i - 1) * (1 - kShort)
    Next i
    
    For i = longPeriod + 1 To lastRow - 1
        longEma(i) = closePrices(i, 1) * kLong + longEma(i - 1) * (1 - kLong)
    Next i
    
    For i = longPeriod + 1 To lastRow - 1
        macd = shortEma(i) - longEma(i)
        ws.Cells(i + 1, macdCol).value = macd
        If i >= longPeriod + signalPeriod Then
            signal = Application.WorksheetFunction.Average(ws.Range(ws.Cells(i - signalPeriod + 1, macdCol), ws.Cells(i, macdCol)))
            ws.Cells(i + 1, signalCol).value = signal
            ws.Cells(i + 1, histCol).value = macd - signal
        End If
    Next i
    
    For i = 2 To longPeriod + 1
        ws.Cells(i, macdCol).value = ""
        ws.Cells(i, signalCol).value = ""
        ws.Cells(i, histCol).value = ""
    Next i
End Sub

Sub CalculateEMA(ws As Worksheet, period As Integer)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim closePrices As Variant
    Dim emaCol As Long
    Dim k As Double
    Dim ema As Double
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < period Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    emaCol = lastCol + 1
    
    closePrices = ws.Range("F2:F" & lastRow).value
    
    ws.Cells(1, emaCol).value = "EMA_" & Str(period)
    
    ema = Application.WorksheetFunction.Average(ws.Range(ws.Cells(2, 6), ws.Cells(1 + period, 6)))
    
    ws.Cells(period + 1, emaCol).value = ema

    k = 2 / (period + 1)
    
    For i = period + 2 To lastRow
        ema = closePrices(i - 1, 1) * k + ema * (1 - k)
        ws.Cells(i, emaCol).value = ema
    Next i
    
    For i = 2 To period
        ws.Cells(i, emaCol).value = ""
    Next i
End Sub


Sub CalculateMA(ws As Worksheet, period As Integer)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim closePrices As Variant
    Dim maCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < period Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    maCol = lastCol + 1
    
    closePrices = ws.Range("F2:F" & lastRow).value
    
    ws.Cells(1, maCol).value = "MA_" & Str(period)
    
    For i = period + 1 To lastRow
        ws.Cells(i, maCol).value = Application.WorksheetFunction.Average(ws.Range(ws.Cells(i - period + 1, 6), ws.Cells(i, 6)))
    Next i
    
    For i = 2 To period
        ws.Cells(i, maCol).value = ""
    Next i
End Sub

Sub CalculateRSI(ws As Worksheet, period As Integer)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim closePrices As Variant
    Dim gains() As Double
    Dim losses() As Double
    Dim avgGain As Double
    Dim avgLoss As Double
    Dim rs As Double
    Dim rsi As Double
    Dim rsiCol  As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < period Then
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    rsiCol = lastCol + 1
    
    closePrices = ws.Range("F2:F" & lastRow).value
    
    ReDim gains(1 To lastRow - 1)
    ReDim losses(1 To lastRow - 1)
    
    For i = 2 To lastRow - 1
        If closePrices(i, 1) > closePrices(i - 1, 1) Then
            gains(i - 1) = closePrices(i, 1) - closePrices(i - 1, 1)
            losses(i - 1) = 0
        Else
            gains(i - 1) = 0
            losses(i - 1) = closePrices(i - 1, 1) - closePrices(i, 1)
        End If
    Next i
    
    ws.Cells(1, rsiCol).value = "RSI"
    For i = 2 To period + 1
        ws.Cells(i, rsiCol).value = ""
    Next i
    
    avgGain = WorksheetFunction.Average(gains, 1, period)
    avgLoss = WorksheetFunction.Average(losses, 1, period)
    
    For i = period + 2 To lastRow
        avgGain = ((avgGain * (period - 1)) + gains(i - 1)) / period
        avgLoss = ((avgLoss * (period - 1)) + losses(i - 1)) / period
        If avgLoss = 0 Then
            rs = 0
            rsi = 100
        Else
            rs = avgGain / avgLoss
            rsi = 100 - (100 / (1 + rs))
        End If
        ws.Cells(i, rsiCol).value = rsi
    Next i
End Sub
