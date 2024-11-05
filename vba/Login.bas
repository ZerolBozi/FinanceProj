Attribute VB_Name = "Login"
Option Explicit

Dim comboBoxEventHandler As ClsComboBoxEvent
Dim buttonEventHandler As ClsBottonEvent

Public positionsCount As Long

Sub PerformLogout()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim http As Object
    Dim url As String
    Dim params As String
    Dim json As Object
    
    params = GetSettingValue("login")
    If params = "" Then
        MsgBox ("﹟ゼ祅")
        Exit Sub
    End If
    
    sheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = params Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' 惠璶ぶ
    If ThisWorkbook.Sheets.Count = 1 Then
        ThisWorkbook.Sheets.Add
    End If
    
    If sheetExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    params = "uid=" & params
    
    url = "http://127.0.0.1:8080/logout?" & params
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.send
    
    Set json = JsonConverter.ParseJSON(http.responseText)
    MsgBox (json("status"))
    Call SaveSettingValue("login", "")
    Call StopUpdate(False)
End Sub

Function PerformLogin(apikey As String, secret As String) As String
    Dim http As Object
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim response As String
    Dim json As Object
    Dim ws As Worksheet
    Dim sheetName As String
    Dim uid As String
    Dim symbolRange As Range
    Dim symbol As String
    
    url = "http://127.0.0.1:8080/login"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "apikey", apikey
    jsonRequest.Add "secret", secret
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
        response = .responseText
    End With
    
    Set json = JsonConverter.ParseJSON(response)
    
    uid = json("uid")
    
    PerformLogin = uid
    
    sheetName = uid
    
    '  浪琩琌
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = sheetName
    
    Set comboBoxEventHandler = New ClsComboBoxEvent
    Set buttonEventHandler = New ClsBottonEvent
    
    Call SetBalance(ws, json("total"), json("free"), json("used"))
    
    Dim JsonArray As Variant
    
    JsonArray = FetchAssets(uid)
    
    Call SetAssets(ws, JsonArray)
    
    Dim data As String
    
    data = FetchBalance(uid, "future")
    
    Call SetFutureBalance(ws, data)
    
    Call SetQuotesSymbol(ws, uid)
    
    data = FetchMaxLeverage(uid, ws.OLEObjects("symbolCbo").Object.value)
    Call SetTrade(ws, data)
    
    Call SetPositions(ws, uid)
    
    Call SetRealTimeQuotes(ws)
    
    Call StartUpdate
End Function

Sub SetPositions(ws As Worksheet, uid As String)
    With ws.Range("I38:AB39")
        .Merge
        .value = ""
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(248, 203, 173)
    End With
    
    positionsCount = 0
    Call SetAllPositions(ws, uid)
End Sub

Sub SetTrade(ws As Worksheet, leverageResponse As String)
    Dim tradeActionObj As OLEObject
    Dim marginModeObj As OLEObject
    Dim longObj As OLEObject
    Dim shortObj As OLEObject
    
    Dim leverage As Object
    Dim maxLeverage As String
    
    Set leverage = JsonConverter.ParseJSON(leverageResponse)
    maxLeverage = leverage("maxLeverage")
    
    With ws.Range("I21:P22")
        .Merge
        .Font.Bold = True
        .value = "ユ"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(248, 203, 173)
    End With
    
    With ws.Range("I23:J24")
        .Merge
        .value = "ユ笆"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I25:J26")
        .Merge
        .value = "玂靡家Α"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I27:J28")
        .Merge
        .value = "篵膘计"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I29:P29")
        .Merge
        .Interior.Color = RGB(248, 203, 173)
    End With
    
    With ws.Range("K27:L28")
        .Merge
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .NumberFormatLocal = "0""x"""
        .value = "1"
    End With
    
    With ws.Range("M23:P28")
        .Merge
        .Font.Bold = True
        .NumberFormatLocal = "0""x"""
        .value = maxLeverage
        .HorizontalAlignment = xlLeft
        .NumberFormatLocal = """*程篵膘计 ""0"
    End With
    
    Set tradeActionObj = ws.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, _
        DisplayAsIcon:=False, Left:=479.25, Top:=352, Width:=96.75, Height:=21.75)
        
    tradeActionObj.Name = "tradeActionCbo"
    
    With tradeActionObj.Object
        .TextAlign = 2
        .AddItem "秨"
        .AddItem "キ"
    End With
    
    tradeActionObj.Object = "秨"
    
    Set comboBoxEventHandler.tradeActionCbo = tradeActionObj.Object
    
    Set marginModeObj = ws.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, _
        DisplayAsIcon:=False, Left:=479.25, Top:=379.5, Width:=96.75, Height:=21.75)
        
    marginModeObj.Name = "marginModeCbo"
        
    With marginModeObj.Object
        .TextAlign = 2
        .AddItem "硋"
        .AddItem ""
    End With
    
    marginModeObj.Object = "硋"
    
    Set comboBoxEventHandler.marginModeCbo = marginModeObj.Object
    
    With ws.Range("I30:J30")
        .Merge
        .value = "基"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I31:J31")
        .Merge
        .value = "ゎ穕基"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I32:J32")
        .Merge
        .value = "ゎ基"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("I33:J33")
        .Merge
        .value = "计秖 (USDT)"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("K30:P30")
        .Merge
    End With
    
    With ws.Range("K31:P31")
        .Merge
    End With
    
    With ws.Range("K32:P32")
        .Merge
    End With
    
    With ws.Range("K33:P33")
        .Merge
    End With
    
    Set longObj = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
        , DisplayAsIcon:=False, Left:=395.25, Top:=521.25, Width:=175.5, Height _
        :=27.75)
    
    longObj.Name = "longbtn"
    
    With longObj.Object
        .Caption = "Open Long"
        .BackColor = RGB(169, 208, 142)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    Set buttonEventHandler.longbtn = longObj.Object
    
    Set shortObj = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
        , DisplayAsIcon:=False, Left:=577.5, Top:=521.25, Width:=175.5, Height _
        :=27.75)
        
    shortObj.Name = "shortbtn"
    
    With shortObj.Object
        .Caption = "Open Short"
        .BackColor = RGB(247, 106, 91)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    Set buttonEventHandler.shortbtn = shortObj.Object
    
    With Range("I36:P36")
        .Merge
        .value = "*狦基玥カ基はぇ玥基"
        .Font.Bold = True
    End With
End Sub

Sub SetAssets(ws As Worksheet, assets As Variant)
    Dim startRow As Integer
    Dim json As Object
    Dim i As Integer
    
    startRow = 7
    
    With ws.Range("A5:F6")
        .Merge
        .value = "Assets"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 40
        .Font.Bold = True
    End With
    
    Dim tmpRange As String

    For i = LBound(assets) To UBound(assets)
        Set json = JsonConverter.ParseJSON(assets(i))
        
        tmpRange = "A" & startRow & ":C" & startRow
        With ws.Range(tmpRange)
            .Merge
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .value = json("asset")
            .Interior.ColorIndex = 35
        End With
        
        tmpRange = "D" & startRow & ":F" & startRow
        With ws.Range(tmpRange)
            .Merge
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .value = json("free")
            .Interior.ColorIndex = 35
        End With
        startRow = startRow + 1
    Next i
End Sub

Sub SetQuotesSymbol(ws As Worksheet, uid As String)
    Dim symbolCbo As OLEObject
    
    With ws.Range("I12:J13")
        .Merge
        .value = "ユ癸"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    Set symbolCbo = ws.OLEObjects.Add(ClassType:="Forms.ComboBox.1", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=480, _
        Top:=182.25, _
        Width:=191.25, _
        Height:=15.75)
    
    symbolCbo.Name = "symbolCbo"
    
    With symbolCbo.Object
            .AddItem "BTCUSDT"
            .AddItem "ETHUSDT"
            .AddItem "LTCUSDT"
            .AddItem "XRPUSDT"
            .AddItem "SOLUSDT"
            .AddItem "DOGEUSDT"
            .AddItem "STXUSDT"
    End With
    
    symbolCbo.Object = "BTCUSDT"
    
    Set comboBoxEventHandler.symbolCbo = symbolCbo.Object
    
    Call SetMarginMode(uid, "BTCUSDT", "Isolated")
    Call SetLeverage(uid, "BTCUSDT", 1)
End Sub

Sub SetRealTimeQuotes(ws As Worksheet)
    With ws.Range("I14:J15")
        .Merge
        .value = "讽玡Θユ基:"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("O14:P15")
        .Merge
        .value = "Funding Rate:"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("I16:J16")
        .Merge
        .value = "秨絃基"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("I17:J17")
        .Merge
        .value = "程蔼基"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("I18:J18")
        .Merge
        .value = "Θユ秖"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("O16:P16")
        .Merge
        .value = "程基"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("O17:P17")
        .Merge
        .value = "Μ絃基"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("K14:N15")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("Q14:T15")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormatLocal = "0.00000000%"
    End With
    
    With ws.Range("K16:N16")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("K17:N17")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("K18:N18")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("Q16:T16")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("Q17:T17")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("I19:T19")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .value = "*厨基5穝Ω狦稱璶氨ゎ厨基叫翴匡よStopRealTimeQuotes秙"
        .Font.Color = RGB(192, 0, 0)
    End With
    
End Sub

Sub SetFutureBalance(ws As Worksheet, balance As String)
    Dim json As Object
    
    Set json = JsonConverter.ParseJSON(balance)
    
      With ws.Range("I4:N5")
        .Merge
        .value = "Future Balance"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(248, 203, 173)
    End With
    
    With ws.Range("I6:J7")
        .Merge
        .value = "Total"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(219, 219, 219)
    End With
    
    With ws.Range("K6:L7")
        .Merge
        .value = "Free"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(219, 219, 219)
    End With
    
    With ws.Range("M6:N7")
        .Merge
        .value = "Used"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(219, 219, 219)
    End With
    
    With ws.Range("I8:J9")
        .Merge
        .value = json("total")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    
    With ws.Range("K8:L9")
        .Merge
        .value = json("free")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With
    
    With ws.Range("M8:N9")
        .Merge
        .value = json("used")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(198, 224, 180)
    End With

End Sub

Sub SetBalance(ws As Worksheet, total As String, free As String, used As String)
    With ws.Range("A1:F2")
        .Merge
        .value = "Balance"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 22
        .Font.Bold = True
    End With
    
    With ws.Range("A3:B3")
        .Merge
        .value = "Total"
        .Interior.ColorIndex = 37
        .Font.Bold = True
    End With
    
    With ws.Range("C3:D3")
        .Merge
        .value = "Free"
        .Interior.ColorIndex = 37
        .Font.Bold = True
    End With
    
    With ws.Range("E3:F3")
        .Merge
        .value = "Used"
        .Interior.ColorIndex = 37
        .Font.Bold = True
    End With
    
    With ws.Range("A4:B4")
        .Merge
        .value = total
        .Interior.ColorIndex = 24
    End With
    
    With ws.Range("C4:D4")
        .Merge
        .value = free
        .Interior.ColorIndex = 24
    End With
    
    With ws.Range("E4:F4")
        .Merge
        .value = used
        .Interior.ColorIndex = 24
    End With
End Sub
