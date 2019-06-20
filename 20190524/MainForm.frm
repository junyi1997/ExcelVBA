VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "MainForm"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'按下在主選單中的輸入資料按鈕
    Unload Me   '關閉表單
    inDataForm1.Show
End Sub

Private Sub CommandButton2_Click()
'按下在主選單中的查詢資料按鈕
    Unload Me   '關閉表單
    SearchForm1.Show
End Sub

Private Sub CommandButton3_Click()
'按下在主選單中的關閉按鈕
    Unload Me   '關閉表單
    ThisWorkbook.Close '關閉活頁簿
    
End Sub

Private Sub UserForm_Click()
    CommandButton2.SetFocus '駐停焦點移到使用者
End Sub
