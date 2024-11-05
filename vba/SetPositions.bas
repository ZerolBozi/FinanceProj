Attribute VB_Name = "SetPositions"
Option Explicit

Sub CleanPositions(ws As Worksheet)
    Dim row As Long
    Dim lastRow As Long
    row = 40
    lastRow = ws.Rows.Count
    
    ws.Range("I" & row & ":AB" & lastRow).Clear
End Sub

Sub SetAllPositions(ws As Worksheet, uid As String)
    Dim positions As Variant
    Dim i As Integer
    Dim json As Object
    Dim row As Long
    Dim longColor As Long
    Dim shortColor As Long
    Dim tmpPositionsCount As Long
    
    row = 40
    longColor = RGB(169, 208, 142)
    shortColor = RGB(247, 106, 91)
    
    positions = FetchPositions(uid)
    
    tmpPositionsCount = UBound(positions)
    
    Dim symbol As String
    Dim side As String
    Dim leverage As String
    Dim amount As String
    Dim price As String
    Dim liquidationPrice As String
    Dim marginRatio As String
    Dim margin As String
    Dim unrealizedPnl As String
    Dim percentage As String
    
    For i = LBound(positions) To UBound(positions)
        Set json = JsonConverter.ParseJSON(positions(i))
        
        symbol = json("symbol")
        side = json("side")
        leverage = json("leverage")
        amount = json("amount")
        price = json("price")
        liquidationPrice = json("liquidationPrice")
        marginRatio = json("marginRatio")
        margin = json("margin")
        unrealizedPnl = json("unrealizedPnl")
        percentage = json("percentage")
        
        If symbol = "" Then
            Exit Sub
        End If
        
        Dim tmpColor As Long
        
        If tmpPositionsCount <> positionsCount Then
            Call CleanPositions(ws)
        End If
        positionsCount = tmpPositionsCount
        
        If side = "long" Then
            tmpColor = longColor
        ElseIf side = "short" Then
            tmpColor = shortColor
        End If
        
        Dim tmpColor2 As Long
        
        If CSng(unrealizedPnl) > 0 Then
            tmpColor2 = RGB(0, 176, 80)
        Else
            tmpColor2 = RGB(255, 0, 0)
        End If
        
        With ws.Range("I" & row & ":J" & row)
            .Merge
            .value = "交易對"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("K" & row & ":L" & row)
            .Merge
            .value = "方向"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("M" & row & ":N" & row)
            .Merge
            .value = "槓桿倍數"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("O" & row & ":P" & row)
            .Merge
            .value = "數量"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("Q" & row & ":R" & row)
            .Merge
            .value = "開倉價位"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("S" & row & ":T" & row)
            .Merge
            .value = "強平價格"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("U" & row & ":V" & row)
            .Merge
            .value = "保證金比例"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("W" & row & ":X" & row)
            .Merge
            .value = "保證金"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("Y" & row & ":Z" & row)
            .Merge
            .value = "盈虧"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        With ws.Range("AA" & row & ":AB" & row)
            .Merge
            .value = "報酬率"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = tmpColor
        End With
        
        row = row + 1
        
        With ws.Range("I" & row & ":J" & row)
            .Merge
            .value = symbol
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("K" & row & ":L" & row)
            .Merge
            .value = side
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("M" & row & ":N" & row)
            .Merge
            .value = leverage
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0""x"""
        End With
        
        With ws.Range("O" & row & ":P" & row)
            .Merge
            .value = amount
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00000 ""USDT"""
        End With
        
        With ws.Range("Q" & row & ":R" & row)
            .Merge
            .value = price
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("S" & row & ":T" & row)
            .Merge
            .value = liquidationPrice
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("U" & row & ":V" & row)
            .Merge
            .value = marginRatio
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00000%"
        End With
        
        With ws.Range("W" & row & ":X" & row)
            .Merge
            .value = margin
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00000 ""USDT"""
        End With
        
        With ws.Range("Y" & row & ":Z" & row)
            .Merge
            .value = unrealizedPnl
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Color = tmpColor2
            .NumberFormatLocal = """+""0.00000 ""USDT"";""-""0.00000 ""USDT"""
        End With
        
        With ws.Range("AA" & row & ":AB" & row)
            .Merge
            .value = percentage & "%"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Color = tmpColor2
            .NumberFormatLocal = """+""0.00%;""-""0.00%"
        End With
        
        row = row + 2
    Next i
    
    With ws.Range("I" & row & ":AB" & row)
        .Merge
        .value = "*倉位與即時報價一同更新，每5秒更新一次倉位資料"
        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)
    End With
End Sub
