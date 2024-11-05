Attribute VB_Name = "Setting"
Option Explicit

Sub SaveSettingValue(key As String, value As String)
    SaveSetting "LoginForm", "Login", key, value
End Sub

Function GetSettingValue(key As String, Optional defaultValue As String = "") As String
    GetSettingValue = GetSetting("LoginForm", "Login", key, defaultValue)
End Function

Function KeyExists(key As String) As Boolean
    Dim tempValue As String
    On Error Resume Next
    tempValue = GetSetting("LoginForm", "Login", key, "KeyNotFound")
    KeyExists = (tempValue <> "KeyNotFound")
    On Error GoTo 0
End Function

Sub DeleteSettingValue(key As String)
    If KeyExists(key) Then
        DeleteSetting "LoginForm", "Login", key
    End If
End Sub
