VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "MainForm"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'���U�b�D��椤����J��ƫ��s
    Unload Me   '�������
    inDataForm1.Show
End Sub

Private Sub CommandButton2_Click()
'���U�b�D��椤���d�߸�ƫ��s
    Unload Me   '�������
    SearchForm1.Show
End Sub

Private Sub CommandButton3_Click()
'���U�b�D��椤���������s
    Unload Me   '�������
    ThisWorkbook.Close '��������ï
    
End Sub

Private Sub UserForm_Click()
    CommandButton2.SetFocus '�n���J�I����ϥΪ�
End Sub
