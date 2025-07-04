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
    Dim txt As String
    Dim pos As Long
    Dim startPos As Long, endPos As Long
    Dim result As String

    ' テキストボックス内の全テキスト
    txt = TextBox1.Text
    ' カーソルの位置
    pos = TextBox1.SelStart

    ' 前の改行の位置を探す（最初からカーソルの位置まで）
    startPos = InStrRev(txt, vbCrLf, pos)
    If startPos = 0 Then
        startPos = 1
    Else
        startPos = startPos + 2 ' 改行の直後
    End If

    ' 次の改行の位置を探す（カーソル位置から後ろ）
    endPos = InStr(pos + 1, txt, vbCrLf)
    If endPos = 0 Then
        endPos = Len(txt) + 1
    End If

    ' 該当行を抽出
    result = Mid(txt, startPos, endPos - startPos)

    MsgBox "現在の行のテキスト: " & vbCrLf & result
    
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
