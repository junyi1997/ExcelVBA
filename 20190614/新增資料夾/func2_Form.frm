VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} func2_Form 
   Caption         =   "篩出輔導名單"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   OleObjectBlob   =   "func2_Form.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "func2_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num2 As Integer
Dim sheet_name As String
Dim file_fullname, filepath As String
Dim rNum, cNum As Integer
Dim row, col As Integer
Dim course_col, need_help_row As Integer
Dim fail_num, need_help_num As Integer
Dim score() As String
Private Sub CommandButton1_Click()
    Unload func2_Form
    menu.Show
End Sub
Private Sub CommandButton3_Click()
    sheet_name = "輔導名單" & unpass_num.Text
    filepath = ActiveWorkbook.Path
    Create_need_help_file
End Sub
Private Sub CommandButton2_Click()
    num2 = Val(unpass_num.Text)
    sheet_name = "輔導名單" & unpass_num.Text
    need_help
End Sub

Sub need_help()
'
' 產出 "輔導名單" 工作表，2課程不及格之名單，成績不及格之儲存格，底色標為紅色
'
    Dim Sheet_add_flag As Boolean
    Sheet_add_flag = True
    For i = 1 To Sheets.Count
        If (Sheets(i).Name = sheet_name) Then
           Sheet_add_flag = False
           Exit For
        End If
    Next
    If Sheet_add_flag = True Then
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets((Sheets.Count)).Name = sheet_name
        Create_need_help
        OK_本檔
        mark_data
    Else
        MsgBox (sheet_name & "-工作表 已存在 !!! ")
    End If
End Sub

Public Sub Create_need_help()
    Sheets("成績資料表").Select
    course_col = 4
    rNum = Range("A1").End(xlDown).row
    cNum = Range("A1").End(xlToRight).Column

    For row = 2 To rNum
        fail_num = 0
 
        For col = course_col To cNum
        '請完成--計算 不及格科目數
        ???
        
        Next
        
        '判別是否需輔導
        If fail_num >= num2 Then
           need_help_num = need_help_num + 1
        End If
    Next
    ReDim score(need_help_num, 1 To cNum) As String    '重新宣告陣列大小
    
    need_help_row = 0
    '複製表頭
    For col = 1 To cNum
        score(need_help_row, col) = Cells(1, col).Value
    Next
    
    For row = 2 To rNum
        fail_num = 0

        For col = course_col To cNum
        '請完成--計算 不及格科目數
        ???
        
        Next
        
        '複製需輔導資料
        If fail_num >= num2 Then
           need_help_row = need_help_row + 1
           For col = 1 To cNum
               score(need_help_row, col) = Cells(row, col).Value
           Next
        End If
    Next
End Sub
Public Sub OK_本檔()
    Sheets(sheet_name).Select
    Range(Cells(1, 1), Cells(need_help_num + 1, cNum)).FormulaArray = score()
End Sub

Public Sub mark_data()
    Sheets(sheet_name).Select
    course_col = 4
    rNum = Range("A1").End(xlDown).row
    cNum = Range("A1").End(xlToRight).Column
    For row = 2 To rNum
        For col = course_col To cNum
            If Cells(row, col).Value < 60 Then
                Cells(row, col).Interior.Pattern = xlSolid
                Cells(row, col).Interior.PatternColorIndex = xlAutomatic
                Cells(row, col).Interior.Color = 255
            End If
        Next
    Next
    Range(Cells(1, 1), Cells(rNum, cNum)).Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous

End Sub
Public Sub Create_need_help_file()
    Create_need_help
    
    file_fullname = filepath & "\" & filename2.Text
    Workbooks.Add
    Workbooks(Workbooks.Count).Activate
    Sheets(1).Name = sheet_name
    Sheets(1).Select
    Range(Cells(1, 1), Cells(need_help_num + 1, cNum)).FormulaArray = score()
    
    mark_data
    
    ActiveWorkbook.SaveAs file_fullname
    ActiveWorkbook.Close

End Sub



