VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   1160
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   1760
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0 ' �蓮�ʒu�w��
    Me.Top = 100           ' �ォ��100�s�N�Z��
    Me.Left = 200          ' ������200�s�N�Z��
    Me.Caption = "�I��"
    Label1.Caption = "��b���̂�"
    Label2.Caption = "�n�̕��̂�"
End Sub


Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        UserForm2.Show vbModeless
    Else
        UserForm2.Hide
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        UserForm3.Show vbModeless
    Else
        UserForm3.Hide
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    if CloseMode = vbFormControlMenu Then
        If MsgBox("����Ƃ�����x�J�E���g���K�v�ɂȂ�܂��B" & vbCr & "���Ă���낵���ł����H", vbCritical + vbYesNo, "����m�F") = vbYes Then
            Unload UserForm2
            Unload UserForm3
            Cancel = False
        Else
            Cancel = True
        End If
    end if
End Sub
