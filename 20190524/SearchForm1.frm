VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm1 
   Caption         =   "SearchForm1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "SearchForm1.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "SearchForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'���U�b�d�߸�Ƥ����^�D�����s
    Unload Me   '�������
    MainForm.Show
End Sub
