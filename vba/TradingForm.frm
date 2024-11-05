VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TradingForm 
   Caption         =   "TradingForm"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "TradingForm.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "TradingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub leverageTb_Change()
    Dim uid As String
    Dim symbol As String
    Dim leverage As String
    
    uid = GetSettingValue("login")
    symbol = Me.symbolCbo.value
    leverage = leverageTb.value
    Call SetLeverage(uid, symbol, leverage)
End Sub

Private Sub marginModeCbo_Change()
    Dim uid As String
    Dim symbol As String
    Dim marginMode As String
    
    uid = GetSettingValue("login")
    symbol = Me.symbolCbo.value
    
    marginMode = marginModeCbo.value
    If marginMode = "�v��" Then
        marginMode = "Isolated"
    ElseIf marginMode = "����" Then
        marginMode = "cross"
    End If
        
    Call SetMarginMode(uid, symbol, marginMode)
End Sub

Private Sub longbtn_Click()
    Dim uid As String
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

    symbol = Me.symbolCbo.value
    leverage = leverageTb.value
    marginMode = marginModeCbo.value
    
    If marginMode = "�v��" Then
        marginMode = "Isolated"
    ElseIf marginMode = "����" Then
        marginMode = "cross"
    End If
    
    price = priceTb.value
    SLPrice = slPriceTb.value
    TPPrice = tpPriceTb.value
    amount = amountTb.value
    
    tradeAction = Me.tradeActionCbo.value
    If tradeAction = "�}��" Then
        If price = "" Then
            msg = OpenMarketOrder(uid, symbol, "buy", amount, SLPrice, TPPrice, marginMode, leverage)
        Else
            msg = OpenLimitOrder(uid, symbol, "buy", amount, price, SLPrice, TPPrice, marginMode, leverage)
        End If
    ElseIf tradeAction = "����" Then
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

    symbol = Me.symbolCbo.value
    leverage = leverageTb.value
    marginMode = marginModeCbo.value
    
    If marginMode = "�v��" Then
        marginMode = "Isolated"
    ElseIf marginMode = "����" Then
        marginMode = "cross"
    End If
    
    price = priceTb.value
    SLPrice = slPriceTb.value
    TPPrice = tpPriceTb.value
    amount = amountTb.value
    
    tradeAction = Me.tradeActionCbo.value
    If tradeAction = "�}��" Then
        If price = "" Then
            msg = OpenMarketOrder(uid, symbol, "sell", amount, SLPrice, TPPrice, marginMode, leverage)
        Else
            msg = OpenLimitOrder(uid, symbol, "sell", amount, price, SLPrice, TPPrice, marginMode, leverage)
        End If
    ElseIf tradeAction = "����" Then
        If price = "" Then
            msg = CloseMarketOrder(uid, symbol, "buy", amount)
        Else
            msg = CloseLimitOrder(uid, symbol, "buy", amount, price)
        End If
    End If
    MsgBox (msg)
End Sub

Private Sub symbolCbo_Change()
    Dim uid As String
    Dim symbol As String
    Dim leverageResponse As String
    Dim leverageJson As Object
    
    uid = GetSettingValue("login")
    symbol = Me.symbolCbo.value
    leverageResponse = FetchMaxLeverage(uid, symbol)
    Set leverageJson = JsonConverter.ParseJSON(leverageResponse)

     Me.maxLeverageLabel.Caption = "*�̤j�b�쭿��: " & leverageJson("maxLeverage")
End Sub

Private Sub tradeActionCbo_Change()
    If Me.tradeActionCbo.value = "�}��" Then
        Me.longbtn.Caption = "Open Long"
        Me.shortbtn.Caption = "Open Short"
    ElseIf Me.tradeActionCbo.value = "����" Then
        Me.longbtn.Caption = "Close Long"
        Me.shortbtn.Caption = "Close Short"
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.longbtn.BackColor = RGB(169, 208, 142)
    Me.shortbtn.BackColor = RGB(247, 106, 91)
    
    With Me.symbolCbo
            .AddItem "BTCUSDT"
            .AddItem "ETHUSDT"
            .AddItem "LTCUSDT"
            .AddItem "XRPUSDT"
            .AddItem "SOLUSDT"
            .AddItem "DOGEUSDT"
            .AddItem "STXUSDT"
    End With
    Me.symbolCbo.value = "BTCUSDT"
    Call symbolCbo_Change
    
    With Me.tradeActionCbo
        .AddItem "�}��"
        .AddItem "����"
    End With
    Me.tradeActionCbo.value = "�}��"
    
    With Me.marginModeCbo
        .AddItem "�v��"
        .AddItem "����"
    End With
    Me.marginModeCbo.value = "�v��"
    Call marginModeCbo_Change
    Call leverageTb_Change
End Sub
