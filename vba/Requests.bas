Attribute VB_Name = "Requests"
Option Explicit

Function FetchUrl(url As String) As String
    Dim http As Object
    Dim JsonString As String
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.send
    
    JsonString = http.responseText
    
    FetchUrl = JsonString
End Function

Function FetchPositions(uid As String) As Variant
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid
    
    url = "http://127.0.0.1:8080/future/positions?" & params
    
    FetchPositions = FetchUrlToVariant(url)
End Function

Function OpenMarketOrder(uid As String, _
    symbol As String, _
    side As String, _
    amount As String, _
    SLPrice As String, _
    TPPrice As String, _
    marginMode As String, _
    leverage As String) As String
    
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    Dim response As String
    Dim json As Object
    
    url = "http://127.0.0.1:8080/future/marketOrder/open"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "side", side
    jsonRequest.Add "amount", amount
    jsonRequest.Add "SLPrice", SLPrice
    jsonRequest.Add "TPPrice", TPPrice
    jsonRequest.Add "marginMode", marginMode
    jsonRequest.Add "leverage", leverage
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
        response = .responseText
    End With
    
    Set json = JsonConverter.ParseJSON(response)
    
    OpenMarketOrder = json("msg")
End Function

Function OpenLimitOrder(uid As String, _
    symbol As String, _
    side As String, _
    amount As String, _
    price As String, _
    SLPrice As String, _
    TPPrice As String, _
    marginMode As String, _
    leverage As String) As String
    
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    Dim response As String
    Dim json As Object
    
    url = "http://127.0.0.1:8080/future/limitOrder/open"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "side", side
    jsonRequest.Add "amount", amount
    jsonRequest.Add "price", price
    jsonRequest.Add "SLPrice", SLPrice
    jsonRequest.Add "TPPrice", TPPrice
    jsonRequest.Add "marginMode", marginMode
    jsonRequest.Add "leverage", leverage
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
        response = .responseText
    End With
    
    Set json = JsonConverter.ParseJSON(response)
    
    OpenLimitOrder = json("msg")
End Function

Function CloseMarketOrder(uid As String, _
    symbol As String, _
    side As String, _
    amount As String) As String
    
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    Dim response As String
    Dim json As Object
    
    url = "http://127.0.0.1:8080/future/marketOrder/close"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "side", side
    jsonRequest.Add "amount", amount
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
        response = .responseText
    End With
    
    Set json = JsonConverter.ParseJSON(response)
    
    CloseMarketOrder = json("msg")
End Function

Function CloseLimitOrder(uid As String, _
    symbol As String, _
    side As String, _
    amount As String, _
    price As String) As String
    
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    Dim response As String
    Dim json As Object
    
    url = "http://127.0.0.1:8080/future/limitOrder/close"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "side", side
    jsonRequest.Add "amount", amount
    jsonRequest.Add "price", price
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
        response = .responseText
    End With
    
    Set json = JsonConverter.ParseJSON(response)
    
    CloseLimitOrder = json("msg")
End Function

Sub SetLeverage(uid As String, symbol As String, leverage As String)
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    
    url = "http://127.0.0.1:8080/future/leverage"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "leverage", leverage
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
    End With
    
End Sub

Sub SetMarginMode(uid As String, symbol As String, marginMode As String)
    Dim url As String
    Dim jsonRequest As Object
    Dim jsonRequestStr As String
    Dim http As Object
    
    url = "http://127.0.0.1:8080/future/marginMode"
    
    Set jsonRequest = CreateObject("Scripting.Dictionary")
    jsonRequest.Add "uid", uid
    jsonRequest.Add "symbol", symbol
    jsonRequest.Add "marginMode", marginMode
    
    jsonRequestStr = JsonConverter.ConvertToJson(jsonRequest)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send jsonRequestStr
    End With
End Sub

Function FetchUrlToVariant(url As String) As Variant
    Dim http As Object
    Dim JsonString As String
    Dim JsonArray As Variant

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.Open "GET", url, False
    http.send
    
    JsonString = http.responseText
    
    JsonString = Mid(JsonString, 2, Len(JsonString) - 2)
    JsonString = Replace(JsonString, "[", "")
    JsonString = Replace(JsonString, "]", "")
    
    JsonString = Replace(JsonString, "\", "")
    JsonString = Replace(JsonString, "},{", "}/{")
    
    JsonArray = Split(JsonString, "/")
    FetchUrlToVariant = JsonArray
End Function

Function FetchMaxLeverage(uid As String, symbol As String) As String
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid & "&symbol=" & symbol
    
    url = "http://127.0.0.1:8080/future/maxLeverage?" & params
    
    FetchMaxLeverage = FetchUrl(url)
End Function

Function FetchNowFundingRate(uid As String, symbol As String) As String
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid & "&symbol=" & symbol
    
    url = "http://127.0.0.1:8080/fetchNow/fundingRate?" & params
    
    FetchNowFundingRate = FetchUrl(url)
End Function

Function FetchNowKines(uid As String, symbol As String, market As String) As String
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid & "&symbol=" & symbol & "&market=" & market
    
    url = "http://127.0.0.1:8080/fetchNow/klines?" & params
    
    FetchNowKines = FetchUrl(url)
End Function

Function FetchBalance(uid As String, market As String) As String
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid & "&market=" & market
    
    url = "http://127.0.0.1:8080/balance?" & params
    
    FetchBalance = FetchUrl(url)
End Function

Function FetchAssets(uid As String) As Variant
    Dim params As String
    Dim url As String
    
    params = "uid=" & uid
    
    url = "http://127.0.0.1:8080/assets?" & params
    
    FetchAssets = FetchUrlToVariant(url)
End Function

Sub FetchKlines(uid As String, market_type As String, symbol As String, timeframe As String, start_time As String, end_time As String)
    Dim params As String
    Dim JsonArray As Variant
    Dim json As Object
    Dim url As String
    Dim ws As Worksheet
    Dim sheetName As String
    Dim i As Long, row As Long
    
    params = "uid=" & uid & "&market=" & market_type & "&symbol=" & symbol & "&timeframe=" & timeframe & "&start=" & start_time & "&end=" & end_time
    
    ' 設置URL
    url = "http://127.0.0.1:8080/fetch/klines?" & params

    JsonArray = FetchUrlToVariant(url)
    
    ' 建立新的工作表
    sheetName = Left(UCase(market_type), 1) & "_" & symbol & "_" & timeframe & "_" & start_time
    
    '  檢查是否存在
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    '設定標頭
    ws.Cells(1, 1).value = "DateTime"
    ws.Cells(1, 2).value = "Unix"
    ws.Cells(1, 3).value = "Open"
    ws.Cells(1, 4).value = "High"
    ws.Cells(1, 5).value = "Low"
    ws.Cells(1, 6).value = "Close"
    ws.Cells(1, 7).value = "Volume"
    
    row = 2

    For i = LBound(JsonArray) To UBound(JsonArray)
        ' 解析每個JSON對象字符串
        On Error GoTo JsonError
        Set json = JsonConverter.ParseJSON(JsonArray(i))

        ' 寫入每個JSON對象的內容
        ws.Cells(row, 1).value = json("datetime")
        ws.Cells(row, 2).value = json("unix")
        ws.Cells(row, 3).value = json("open")
        ws.Cells(row, 4).value = json("high")
        ws.Cells(row, 5).value = json("low")
        ws.Cells(row, 6).value = json("close")
        ws.Cells(row, 7).value = json("volume")
        
        row = row + 1
    Next i
    Exit Sub
    
JsonError:
    MsgBox "Error parsing JSON: " & Err.Description
End Sub

