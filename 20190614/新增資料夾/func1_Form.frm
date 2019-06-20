VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} func1_Form 
   Caption         =   "篩出班級成績"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "func1_Form.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "func1_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sheet_name1 As String
Dim file_fullname, filepath As String
Dim rNum, cNum As Integer
Dim row, col As Integer
Dim course_col, class_row As Integer
Dim std_num As Integer
Dim class_data_arr() As String
Private Sub CommandButton1_Click()
    sheet_name1 = "班級-" & class.Text
    class_data
End Sub

Private Sub CommandButton3_Click()
    sheet_name1 = "班級-" & class.Text
    filepath = ActiveWorkbook.Path
    Create_class_data_file
End Sub



Private Sub func1_quit_Click()
    Unload func1_Form
    menu.Show
End Sub


Sub class_data()
'
' 產出 "班級n" 工作表
'
    Dim Sheet_add_flag As Boolean
    Sheet_add_flag = True
    '判別 sheet_name1工作表是否存在
    For i = 1 To Sheets.Count
        If (Sheets(i).Name = sheet_name1) Then
           Sheet_add_flag = False
           Exit For
        End If
    Next
    If Sheet_add_flag = True Then
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets((Sheets.Count)).Name = sheet_name1
        Create_class_data
        OK_本檔
    Else
        MsgBox (sheet_name1 & "-工作表 已存在 !!! ")
    End If
End Sub

Public Sub Create_class_data()
    Sheets("成績資料表").Select
    course_col = 4
    rNum = Range("A1").End(xlDown).row
    cNum = Range("A1").End(xlToRight).Column
    
    std_num = 0
    For row = 2 To rNum
    '請完成--計算 class.Text 班級人
     ???

    Next
    
    ReDim class_data_arr(std_num, 1 To cNum) As String    '重新宣告陣列大小
    
    class_row = 0
    For col = 1 To cNum
    '請完成--複製表頭
     ???
     
    Next

    For row = 2 To rNum
    '請完成--複製班級資料
     ???
     
    Next

End Sub
Public Sub OK_本檔()
    Sheets(sheet_name1).Select
    Range(Cells(1, 1), Cells(std_num + 1, cNum)).FormulaArray = class_data_arr()
End Sub

Public Sub Create_class_data_file()
    Create_class_data
    
    file_fullname = filepath & "\" & filename1.Text
    Workbooks.Add
    Workbooks(Workbooks.Count).Activate
    Sheets(1).Name = sheet_name1
    Sheets(1).Select
    Range(Cells(1, 1), Cells(std_num + 1, cNum)).FormulaArray = class_data_arr()
    ActiveWorkbook.SaveAs file_fullname
    ActiveWorkbook.Close
    
End Sub

