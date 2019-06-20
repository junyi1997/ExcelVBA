VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPassword 
   Caption         =   "UserForm1"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPassword.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    txtPassWord.IMEMode = 2     '關閉中文輸入
    txtPassWord.PasswordChar = "*"  '替代字元為*
    txtPassWord.MaxLength = 6   '最多6個字元
    txtUser.SetFocus            '駐停焦點移到使用者
End Sub

Private Sub txtUser_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtPassWord.SetFocus  '若按Enter鍵就跳到密碼
End Sub

Private Sub txtPassWord_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then cmdOK.SetFocus    '若按Enter鍵就跳到確定鈕
    Select Case KeyAscii
        Case 48 To 57       '數字
        Case 65 To 90       '大寫英文字母
        Case 97 To 122      '小寫英文字母
        Case Else
            KeyAscii = 0    '清除輸入的字元
    End Select
End Sub

Private Sub cmdOK_Click()
    Static num As Integer   'num要累計所以要宣告成靜態變數
    num = num + 1     '輸入次數加1
    If num < 3 Then   '若輸入次數小於3
        If txtUser.Text = "root" And txtPassWord.Text = "123456" Then    '若正確
            Unload Me   '關閉表單
            MainForm.Show
        Else
            MsgBox "錯誤！"
            txtUser.SetFocus    '駐停焦點移到使用者
        End If
    Else
        MsgBox "錯誤三次！"
        ThisWorkbook.Close      '關閉活頁簿
    End If
End Sub

Private Sub cmdUnload_Click()
    ThisWorkbook.Close          '關閉活頁簿
End Sub

