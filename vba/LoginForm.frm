VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoginBtn_Click()
    Dim apikey As String
    Dim secret As String
    Dim response As String
    
    If GetSettingValue("login", "") <> "" Then
        MsgBox ("Already login")
        Exit Sub
    End If
    
    apikey = Me.ApikeyText.Text
    secret = Me.SecretText.Text
    
    If Me.RememberCbo.value = True Then
        SaveSettingValue "apiKey", apikey
        SaveSettingValue "secret", secret
    Else
        DeleteSettingValue "apiKey"
        DeleteSettingValue "secret"
    End If
    
    response = PerformLogin(apikey, secret)
    If response <> "" Then
        SaveSettingValue "login", response
        MsgBox ("login successed, uid=" & response)
        Unload Me
    Else
        MsgBox ("login failed")
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.ApikeyText.Text = GetSettingValue("apiKey", "")
    Me.SecretText.Text = GetSettingValue("secret", "")
    Me.RememberCbo.value = (Me.ApikeyText.Text <> "" And Me.SecretText.Text <> "")
End Sub
