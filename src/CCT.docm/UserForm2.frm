VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6050
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9960.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
    Dim txt As String
    Dim pos As Long
    Dim startPos As Long, endPos As Long
    Dim result As String

    ' �e�L�X�g�{�b�N�X���̑S�e�L�X�g
    txt = TextBox1.Text
    ' �J�[�\���̈ʒu
    pos = TextBox1.SelStart

    ' �O�̉��s�̈ʒu��T���i�ŏ�����J�[�\���̈ʒu�܂Łj
    startPos = InStrRev(txt, vbCrLf, pos)
    If startPos = 0 Then
        startPos = 1
    Else
        startPos = startPos + 2 ' ���s�̒���
    End If

    ' ���̉��s�̈ʒu��T���i�J�[�\���ʒu������j
    endPos = InStr(pos + 1, txt, vbCrLf)
    If endPos = 0 Then
        endPos = Len(txt) + 1
    End If

    ' �Y���s�𒊏o
    result = Mid(txt, startPos, endPos - startPos)

    MsgBox "���݂̍s�̃e�L�X�g: " & vbCrLf & result
    
    result = Replace(result, vbCr, "^p")
    
    With ActiveDocument.Content.Find
        .ClearFormatting
        .text = result
        .Forward = True
        .Wrap = wdFindStop
        .Execute

        If .Found Then
            ' �����ʒu�ɃW�����v
            Dim rng As Range
            Set rng = .Parent
            rng.Select
            ' ��ʂɕ\�������悤�X�N���[��
            ActiveWindow.ScrollIntoView rng, True
            Me.Hide
            UserForm4.CheckBox1.Value = False
        Else
            MsgBox "������u" & result & "�v�͖{�����Ɍ�����܂���ł����B", vbInformation
        End If
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    if CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        UserForm4.CheckBox1.Value = False
    end if
End Sub
