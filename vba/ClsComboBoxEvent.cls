VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsComboBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents symbolCbo As MSForms.ComboBox
Attribute symbolCbo.VB_VarHelpID = -1
Public WithEvents tradeActionCbo As MSForms.ComboBox
Attribute tradeActionCbo.VB_VarHelpID = -1
Public WithEvents marginModeCbo As MSForms.ComboBox
Attribute marginModeCbo.VB_VarHelpID = -1

Private Sub symbolCbo_Change()
    Dim uid As String
    Dim ws As Worksheet
    Dim leverageResponse As String
    Dim leverageJson As Object
    Dim symbol As String
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)
    
    symbol = ws.OLEObjects("symbolCbo").Object.value
    
    leverageResponse = FetchMaxLeverage(uid, symbol)
    Set leverageJson = JsonConverter.ParseJSON(leverageResponse)
    
    ws.Range("K27").value = "1"
    ws.Range("M23").value = leverageJson("maxLeverage")
End Sub

Private Sub tradeActionCbo_Change()
    Dim uid As String
    Dim ws As Worksheet
    Dim leverageResponse As String
    Dim leverageJson As Object
    Dim tradeAction As String
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)
    
    tradeAction = ws.OLEObjects("tradeActionCbo").Object.value
    
    If tradeAction = "¶}­Ü" Then
        ws.OLEObjects("longbtn").Object.Caption = "Open Long"
        ws.OLEObjects("shortbtn").Object.Caption = "Open Short"
    ElseIf tradeAction = "¥­­Ü" Then
        ws.OLEObjects("longbtn").Object.Caption = "Close Long"
        ws.OLEObjects("shortbtn").Object.Caption = "Close Short"
    End If
    
End Sub

Private Sub marginModeCbo_Change()
    Dim uid As String
    Dim ws As Worksheet
    Dim leverageResponse As String
    Dim leverageJson As Object
    Dim symbol As String
    Dim marginMode As String
    
    uid = GetSettingValue("login")
    Set ws = ThisWorkbook.Sheets(uid)
    
    symbol = ws.OLEObjects("symbolCbo").Object.value
    marginMode = ws.OLEObjects("marginModeCbo").Object.value
    
    If marginMode = "³v­Ü" Then
        marginMode = "Isolated"
    ElseIf marginMode = "¥þ­Ü" Then
        marginMode = "cross"
    End If
        
    Call SetMarginMode(uid, symbol, marginMode)
End Sub
