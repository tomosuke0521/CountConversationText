VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6050
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9960.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
    Dim pos As Long
    Dim totalLen As Long
    Dim result As String

    TextBox1.SetFocus
    DoEvents

    With TextBox1
        pos = .SelStart
        totalLen = Len(.text)

        ' 選択されてないなら強制的に10文字選択
        If .SelLength = 0 Then
            If pos >= totalLen Then
                MsgBox "カーソルが末尾にあります。", vbExclamation
                Exit Sub
            End If
            
            
            Dim AfterPos As Long: AfterPos = pos
            Dim PrePos As Long: PrePos = pos
            Dim SelLen As Long: SelLen = 0
            Do While SelLen < 100
                DoEvents
                .SelStart = PrePos
                .SelLength = 1
                If .SelText = vbCr And Not (SelLen = 0) Then
                    AfterPos = PrePos + 1
                    Do While SelLen < 200
                        DoEvents
                        .SelStart = AfterPos
                        .SelLength = 1
                        If .SelText = vbCr Then
                            .SelStart = PrePos + 1
                            .SelLength = AfterPos - PrePos - 1
                            result = .SelText
                            Exit Do
                        End If
                        AfterPos = AfterPos + 1
                        SelLen = SelLen + 1
                    Loop
                    Exit Do
                End If
                
                PrePos = PrePos - 1
                SelLen = SelLen + 1
            Loop
            
        Else
            result = .SelText
        End If
    End With
    
    result = Replace(result, vbCr, "^p")
    
    With ActiveDocument.Content.Find
        .ClearFormatting
        .text = result
        .Forward = True
        .Wrap = wdFindStop
        .Execute

        If .Found Then
            ' 検索位置にジャンプ
            Dim rng As Range
            Set rng = .Parent
            rng.Select
            ' 画面に表示されるようスクロール
            ActiveWindow.ScrollIntoView rng, True
            Me.Hide
            UserForm4.CheckBox1.Value = False
        Else
            MsgBox "文字列「" & result & "」は本文内に見つかりませんでした。", vbInformation
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
