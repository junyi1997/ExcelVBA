VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} inDataForm1 
   Caption         =   "inDataForm1"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "inDataForm1.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "inDataForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'按下在輸入資料中的回主選單按鈕
    Unload Me   '關閉表單
    MainForm.Show
End Sub

