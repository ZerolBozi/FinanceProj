Attribute VB_Name = "DataFetch"
Option Explicit

Function FetchAssets(uid As String) As Variant
    Dim params As String
    Dim http As Object
    Dim JsonString As String
    Dim JsonArray As Variant
    Dim url As String
    
    params = "uid=" & uid
    
    url = "http://127.0.0.1:8080/assets?" & params
    
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
    FetchAssets = JsonArray
End Function

Sub FetchAndParseJSON(uid As String, market_type As String, symbol As String, timeframe As String, start_time As String, end_time As String)
    Dim params As String
    Dim http As Object
    Dim JsonString As String
    Dim JsonArray As Variant
    Dim Json As Object
    Dim url As String
    Dim ws As Worksheet
    Dim sheetName As String
    Dim i As Long, row As Long
    
    params = "uid=" & uid & "&market=" & market_type & "&symbol=" & symbol & "&timeframe=" & timeframe & "&start=" & start_time & "&end=" & end_time
    
    ' �]�mURL
    url = "http://127.0.0.1:8080/fetch?" & params

    ' �Ы�WinHttpRequest��H
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' �o�eHTTP GET�ШD
    http.Open "GET", url, False
    http.send

    ' �����^��JSON�r�Ŧ�
    JsonString = http.responseText
    
    JsonString = Mid(JsonString, 2, Len(JsonString) - 2)
    JsonString = Replace(JsonString, "[", "")
    JsonString = Replace(JsonString, "]", "")
    JsonString = Replace(JsonString, "\", "")
    JsonString = Replace(JsonString, "},{", "}/{")

    JsonArray = Split(JsonString, "/")
    
    ' �إ߷s���u�@��
    sheetName = Left(UCase(market_type), 1) & "_" & symbol & "_" & timeframe & "_" & start_time & "_" & end_time
    
    '  �ˬd�O�_�s�b
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    '�]�w���Y
    ws.Cells(1, 1).value = "DateTime"
    ws.Cells(1, 2).value = "Unix"
    ws.Cells(1, 3).value = "Open"
    ws.Cells(1, 4).value = "High"
    ws.Cells(1, 5).value = "Low"
    ws.Cells(1, 6).value = "Close"
    ws.Cells(1, 7).value = "Volume"
    
    row = 2

    For i = LBound(JsonArray) To UBound(JsonArray)
        ' �ѪR�C��JSON��H�r�Ŧ�
        On Error GoTo JsonError
        Set Json = JsonConverter.ParseJSON(JsonArray(i))

        ' �g�J�C��JSON��H�����e
        ws.Cells(row, 1).value = Json("datetime")
        ws.Cells(row, 2).value = Json("unix")
        ws.Cells(row, 3).value = Json("open")
        ws.Cells(row, 4).value = Json("high")
        ws.Cells(row, 5).value = Json("low")
        ws.Cells(row, 6).value = Json("close")
        ws.Cells(row, 7).value = Json("volume")
        
        row = row + 1
    Next i
    Exit Sub
    
JsonError:
    MsgBox "Error parsing JSON: " & Err.Description
End Sub

