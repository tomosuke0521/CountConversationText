VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   1160
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   1760
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0 ' 手動位置指定
    Me.Top = 100           ' 上から100ピクセル
    Me.Left = 200          ' 左から200ピクセル
    Me.Caption = "選択"
    Label1.Caption = "会話文のみ"
    Label2.Caption = "地の文のみ"
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
        If MsgBox("閉じるともう一度カウントが必要になります。" & vbCr & "閉じてもよろしいですか？", vbCritical + vbYesNo, "動作確認") = vbYes Then
            Unload UserForm2
            Unload UserForm3
            Cancel = False
        Else
            Cancel = True
        End If
    end if
End Sub
