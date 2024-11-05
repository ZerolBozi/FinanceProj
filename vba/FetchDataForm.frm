VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FetchDataForm 
   Caption         =   "FetchData"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   OleObjectBlob   =   "FetchDataForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "FetchDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FetchBtn_Click()
    If GetSettingValue("login", "") = "" Then
            MsgBox ("尚未登入，無法獲取資料")
            Exit Sub
    End If
    
    If Me.MarketTypeCbo.value = "" Or Me.SymbolComboBox.value = "" Or Me.TimeframeComboBox = "" Or Me.StartYearCbo.value = "" Or Me.StartMonthCbo.value = "" Or Me.StartDayCbo.value = "" Or Me.EndYearCbo.value = "" Or Me.EndMonthCbo.value = "" Or Me.EndDayCbo.value = "" Then
        MsgBox "沒有選擇完整，請重新選擇"
        Exit Sub
    End If
    
    Dim start_date As String
    Dim end_date As String
    Dim uid  As String
    start_date = Me.StartYearCbo.value & "-" & Me.StartMonthCbo.value & "-" & Me.StartDayCbo.value
    end_date = Me.EndYearCbo.value & "-" & Me.EndMonthCbo.value & "-" & Me.EndDayCbo.value
    
    uid = GetSettingValue("login", "")
    Call FetchKlines(uid, Me.MarketTypeCbo.value, Me.SymbolComboBox.value, Me.TimeframeComboBox.value, start_date, end_date)
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me.MarketTypeCbo
            .AddItem "spot"
            .AddItem "future"
    End With
    Me.MarketTypeCbo.value = "spot"

    With Me.SymbolComboBox
            .AddItem "BTCUSDT"
            .AddItem "ETHUSDT"
            .AddItem "LTCUSDT"
            .AddItem "XRPUSDT"
            .AddItem "SOLUSDT"
            .AddItem "DOGEUSDT"
            .AddItem "STXUSDT"
    End With
    Me.SymbolComboBox.value = "BTCUSDT"
            
    With Me.TimeframeComboBox
        .AddItem "1m"
        .AddItem "5m"
        .AddItem "15m"
        .AddItem "30m"
        .AddItem "1h"
        .AddItem "4h"
        .AddItem "1d"
    End With
    Me.TimeframeComboBox.value = "1d"
    
    Dim i As Integer
    
    For i = Year(Date) - 10 To Year(Date) + 10
        Me.StartYearCbo.AddItem i
    Next i
    Me.StartYearCbo.value = Year(Date)
    
    For i = 1 To 12
        Me.StartMonthCbo.AddItem i
    Next i
    Me.StartMonthCbo.value = Month(Date)
    
    For i = Year(Date) - 10 To Year(Date) + 10
        Me.EndYearCbo.AddItem i
    Next i
    Me.EndYearCbo.value = Year(Date)
    
    For i = 1 To 12
        Me.EndMonthCbo.AddItem i
    Next i
    Me.EndMonthCbo.value = Month(Date)
    
End Sub

Private Sub StartMonthCbo_Change()
    UpdateStartDays
End Sub

Private Sub EndMonthCbo_Change()
    UpdateEndDays
End Sub

Private Sub UpdateStartDays()
    Dim i As Integer
    Dim selectedYear As Integer
    Dim selectedMonth As Integer

    selectedYear = CInt(Me.StartYearCbo.value)
    selectedMonth = CInt(Me.StartMonthCbo.value)

    Me.StartDayCbo.Clear
    For i = 1 To Day(DateSerial(selectedYear, selectedMonth + 1, 0))
        Me.StartDayCbo.AddItem i
    Next i
    Me.StartDayCbo.value = Day(Date)
End Sub

Private Sub UpdateEndDays()
    Dim i As Integer
    Dim selectedYear As Integer
    Dim selectedMonth As Integer

    selectedYear = CInt(Me.EndYearCbo.value)
    selectedMonth = CInt(Me.EndMonthCbo.value)

    Me.EndDayCbo.Clear
    For i = 1 To Day(DateSerial(selectedYear, selectedMonth + 1, 0))
        Me.EndDayCbo.AddItem i
    Next i
    Me.EndDayCbo.value = Day(Date)
End Sub
