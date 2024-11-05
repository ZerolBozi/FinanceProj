VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalcIndicators 
   Caption         =   "Calculate Indicators"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "CalcIndicators.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "CalcIndicators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CleanIndicators(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim startCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    startCol = ws.Range("H1").Column
    lastCol = ws.Columns.Count
    
    ws.Range(ws.Cells(1, startCol), ws.Cells(lastRow, lastCol)).Clear
End Sub

Private Sub CalculateBtn_Click()
    Dim sheetName As String
    Dim ws As Worksheet
    
    sheetName = Me.WorksheetCbo.value
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Call CleanIndicators(ws)
    
    If Me.RSIChb.value Then
        Call CalculateRSI(ws, Me.RSIParam)
    End If
    
    If Me.MAChb1.value Then
        Call CalculateMA(ws, Me.MA1Param)
    End If
    
    If Me.MAChb2.value Then
        Call CalculateMA(ws, Me.MA2Param)
    End If
    
    If Me.MAChb3.value Then
        Call CalculateMA(ws, Me.MA3Param)
    End If
    
    If Me.EMAChb1.value Then
        Call CalculateEMA(ws, Me.EMA1Param)
    End If
    
    If Me.EMAChb2.value Then
        Call CalculateEMA(ws, Me.EMA2Param)
    End If
    
    If Me.EMAChb3.value Then
        Call CalculateEMA(ws, Me.EMA3Param)
    End If
    
    If Me.MACDChb.value Then
        Call CalculateMACD(ws, Me.MACDParam1, Me.MACDParam2, 9)
    End If
    
    If Me.BBChb.value Then
        Call CalculateBB(ws, Me.BBParam1, Me.BBParam2)
    End If
    
    If Me.VOChb.value Then
        Call CalculateVO(ws, Me.VOParam)
    End If
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    If Left(ActiveSheet.Name, 1) = "S" Or Left(ActiveSheet.Name, 1) = "F" Then
        Me.WorksheetCbo.value = ActiveSheet.Name
    End If
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) = "S" Or Left(ws.Name, 1) = "F" Then
            Me.WorksheetCbo.AddItem ws.Name
        End If
    Next ws

End Sub
