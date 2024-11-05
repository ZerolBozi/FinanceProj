Attribute VB_Name = "UpdateData"
Option Explicit

Dim NextTick As Double

Sub StartUpdate()
    NextTick = Now + TimeValue("00:00:05")
    Application.OnTime EarliestTime:=NextTick, Procedure:="Update", _
        Schedule:=True
End Sub

Sub StopUpdate(msg As Boolean)
    On Error Resume Next
    Application.OnTime EarliestTime:=NextTick, Procedure:="Update", _
        Schedule:=False
    On Error GoTo 0
    
    If msg Then
        MsgBox "Stop Real Time Quotes Successed"
    End If
End Sub

Sub Update()
    Dim uid As String
    Dim ws As Worksheet
    Dim symbol  As String
    Dim ohlcv As String
    Dim rate As String
    Dim ohlcvJson As Object
    Dim rateJson As Object
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)
    symbol = ws.OLEObjects("symbolCbo").Object.value
    
    ohlcv = FetchNowKines(uid, symbol, "future")
    rate = FetchNowFundingRate(uid, symbol)
    
    Set ohlcvJson = JsonConverter.ParseJSON(ohlcv)
    Set rateJson = JsonConverter.ParseJSON(rate)
    
    ws.Range("K14").value = ohlcvJson("close")
    ws.Range("K16").value = ohlcvJson("open")
    ws.Range("K17").value = ohlcvJson("high")
    ws.Range("K18").value = ohlcvJson("volume")
    ws.Range("Q14").value = rateJson("fundingRate")
    ws.Range("Q16").value = ohlcvJson("low")
    ws.Range("Q17").value = ohlcvJson("close")
    
    Call SetAllPositions(ws, uid)
    
    StartUpdate
End Sub
