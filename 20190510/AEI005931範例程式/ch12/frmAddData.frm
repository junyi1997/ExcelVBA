VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddData 
   Caption         =   "UserForm1"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "frmAddData.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frmAddData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = Sheets("���")
    With lstData
        .List = ws.UsedRange.Value  '�[�J�M�涵��
        .ColumnCount = 6       '6��
        .ColumnWidths = "30,50,30,30,30,30" '�]�w��e
    End With
End Sub

Private Sub cmdAdd_Click()
    Dim ws As Worksheet
    Set ws = Sheets(1)
    Dim last As Integer
    last = ws.UsedRange.Rows.Count + 1  '�g�J���C��
    Dim sel As Integer
    sel = lstData.ListIndex '��ܶ��ت����ޭ�
    For c = 1 To 6  '�v��g�J���
        ws.Cells(last, c).Value = lstData.List(sel, c - 1)
    Next
End Sub
