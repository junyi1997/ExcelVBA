VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} func1_Form 
   Caption         =   "�z�X�Z�Ŧ��Z"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "func1_Form.frx":0000
   StartUpPosition =   1  '���ݵ�������
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
    sheet_name1 = "�Z��-" & class.Text
    class_data
End Sub

Private Sub CommandButton3_Click()
    sheet_name1 = "�Z��-" & class.Text
    filepath = ActiveWorkbook.Path
    Create_class_data_file
End Sub



Private Sub func1_quit_Click()
    Unload func1_Form
    menu.Show
End Sub


Sub class_data()
'
' ���X "�Z��n" �u�@��
'
    Dim Sheet_add_flag As Boolean
    Sheet_add_flag = True
    '�P�O sheet_name1�u�@��O�_�s�b
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
        OK_����
    Else
        MsgBox (sheet_name1 & "-�u�@�� �w�s�b !!! ")
    End If
End Sub

Public Sub Create_class_data()
    Sheets("���Z��ƪ�").Select
    course_col = 4
    rNum = Range("A1").End(xlDown).row
    cNum = Range("A1").End(xlToRight).Column
    
    std_num = 0
    For row = 2 To rNum
    '�Ч���--�p�� class.Text �Z�ŤH
     ???

    Next
    
    ReDim class_data_arr(std_num, 1 To cNum) As String    '���s�ŧi�}�C�j�p
    
    class_row = 0
    For col = 1 To cNum
    '�Ч���--�ƻs���Y
     ???
     
    Next

    For row = 2 To rNum
    '�Ч���--�ƻs�Z�Ÿ��
     ???
     
    Next

End Sub
Public Sub OK_����()
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

