VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPassword 
   Caption         =   "UserForm1"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPassword.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    txtPassWord.IMEMode = 2     '���������J
    txtPassWord.PasswordChar = "*"  '���N�r����*
    txtPassWord.MaxLength = 6   '�̦h6�Ӧr��
    txtUser.SetFocus            '�n���J�I����ϥΪ�
End Sub

Private Sub txtUser_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then txtPassWord.SetFocus  '�Y��Enter��N����K�X
End Sub

Private Sub txtPassWord_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then cmdOK.SetFocus    '�Y��Enter��N����T�w�s
    Select Case KeyAscii
        Case 48 To 57       '�Ʀr
        Case 65 To 90       '�j�g�^��r��
        Case 97 To 122      '�p�g�^��r��
        Case Else
            KeyAscii = 0    '�M����J���r��
    End Select
End Sub

Private Sub cmdOK_Click()
    Static num As Integer   'num�n�֭p�ҥH�n�ŧi���R�A�ܼ�
    num = num + 1     '��J���ƥ[1
    If num < 3 Then   '�Y��J���Ƥp��3
        If txtUser.Text = "root" And txtPassWord.Text = "123456" Then    '�Y���T
            Unload Me   '�������
            MainForm.Show
        Else
            MsgBox "���~�I"
            txtUser.SetFocus    '�n���J�I����ϥΪ�
        End If
    Else
        MsgBox "���~�T���I"
        ThisWorkbook.Close      '��������ï
    End If
End Sub

Private Sub cmdUnload_Click()
    ThisWorkbook.Close          '��������ï
End Sub

