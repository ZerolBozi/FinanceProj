VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsBottonEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents longbtn As CommandButton
Attribute longbtn.VB_VarHelpID = -1
Public WithEvents shortbtn As CommandButton
Attribute shortbtn.VB_VarHelpID = -1

Private Sub longbtn_Click()
    Dim uid As String
    Dim ws As Worksheet
    Dim symbol As String
    Dim marginMode As String
    Dim leverage As String
    Dim price As String
    Dim SLPrice As String
    Dim TPPrice As String
    Dim amount As String
    Dim tradeAction As String
    
    Dim msg As String
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)

    symbol = ws.OLEObjects("symbolCbo").Object.value
    marginMode = ws.OLEObjects("marginModeCbo").Object.value
    
    If marginMode = "逐倉" Then
        marginMode = "Isolated"
    ElseIf marginMode = "全倉" Then
        marginMode = "cross"
    End If
    
    leverage = ws.Range("K27").value
    price = ws.Range("K30").value
    SLPrice = ws.Range("K31").value
    TPPrice = ws.Range("K32").value
    amount = ws.Range("K33").value
    
    If amount = "" Then
        MsgBox ("未輸入數量")
        Exit Sub
    End If
    
    tradeAction = ws.OLEObjects("tradeActionCbo").Object.value
    If tradeAction = "開倉" Then
        If price = "" Then
            msg = OpenMarketOrder(uid, symbol, "buy", amount, SLPrice, TPPrice, marginMode, leverage)
        Else
            msg = OpenLimitOrder(uid, symbol, "buy", amount, price, SLPrice, TPPrice, marginMode, leverage)
        End If
    ElseIf tradeAction = "平倉" Then
        If price = "" Then
            msg = CloseMarketOrder(uid, symbol, "sell", amount)
        Else
            msg = CloseLimitOrder(uid, symbol, "sell", amount, price)
        End If
    End If
    MsgBox (msg)
End Sub

Private Sub shortbtn_Click()
    Dim uid As String
    Dim ws As Worksheet
    Dim symbol As String
    Dim marginMode As String
    Dim leverage As String
    Dim price As String
    Dim SLPrice As String
    Dim TPPrice As String
    Dim amount As String
    Dim tradeAction As String
    
    Dim msg As String
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)

    symbol = ws.OLEObjects("symbolCbo").Object.value
    marginMode = ws.OLEObjects("marginModeCbo").Object.value
    
    If marginMode = "逐倉" Then
        marginMode = "Isolated"
    ElseIf marginMode = "全倉" Then
        marginMode = "cross"
    End If
    
    leverage = ws.Range("K27").value
    price = ws.Range("K30").value
    SLPrice = ws.Range("K31").value
    TPPrice = ws.Range("K32").value
    amount = ws.Range("K33").value
    
    If amount = "" Then
        MsgBox ("未輸入數量")
        Exit Sub
    End If
    
    tradeAction = ws.OLEObjects("tradeActionCbo").Object.value
    If tradeAction = "開倉" Then
        If price = "" Then
            msg = OpenMarketOrder(uid, symbol, "sell", amount, SLPrice, TPPrice, marginMode, leverage)
        Else
            msg = OpenLimitOrder(uid, symbol, "sell", amount, price, SLPrice, TPPrice, marginMode, leverage)
        End If
    ElseIf tradeAction = "平倉" Then
        If price = "" Then
            msg = CloseMarketOrder(uid, symbol, "buy", amount)
        Else
            msg = CloseLimitOrder(uid, symbol, "buy", amount, price)
        End If
    End If
    MsgBox (msg)
End Sub
