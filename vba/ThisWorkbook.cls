VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call SaveSettingValue("login", "")
End Sub

Private Sub Workbook_Open()
    AddReferences
    ShowLoginForm
End Sub

Private Sub AddReferences()
    Dim vbProj As Object
    Dim chkRef As Object
    Dim bFound As Boolean

    Set vbProj = ThisWorkbook.VBProject

    bFound = False
    For Each chkRef In vbProj.References
        If chkRef.Name = "Scripting" Then
            bFound = True
            Exit For
        End If
    Next chkRef

    If Not bFound Then
        vbProj.References.AddFromFile "C:\Windows\System32\scrrun.dll"
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim uid As String
    
    uid = GetSettingValue("login", "")

    If uid = "" Then
        Exit Sub
    End If
    
    If Sh.Name <> uid Then
        Exit Sub
    End If
    
    If Intersect(Target, Sh.Range("K27")) Is Nothing Then
        Exit Sub
    End If
    
    Dim symbol As String
    Dim leverage As String
    
    symbol = Sh.OLEObjects("symbolCbo").Object.value
    leverage = Target.value
    Call SetLeverage(uid, symbol, leverage)
End Sub
