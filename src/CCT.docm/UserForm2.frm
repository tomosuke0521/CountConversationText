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
    Dim selPos As Long
    Dim lines() As String
    Dim i As Long
    Dim charIndex As Long
    Dim result As String

    txt = TextBox1.Text
    selPos = TextBox1.SelStart
    lines = Split(txt, vbCr)

    charIndex = 0
    For i = 0 To UBound(lines)
        ' selPos がこの行の中にあるか？
        If selPos >= charIndex And selPos <= charIndex + Len(lines(i)) Then
            result = lines(i)
            Exit For
        End If
        charIndex = charIndex + Len(lines(i)) + 2  ' vbCrLf = 2文字
    Next i

    MsgBox "カーソルがある行の内容：" & vbCrLf & result
    
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
