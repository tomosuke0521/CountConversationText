Attribute VB_Name = "Module1"
Option Explicit

Public SHOULD_STOP As Integer
Public SHOULD_OPEN_UF As Integer
Private non As Variant

Private Function GetWithoutLB(ByVal inp_text As String) As Long
    Dim count_lb As Long
    count_lb = UBound(Split(inp_text, vbCr))
    Dim count_tb As Long
    count_tb = UBound(Split(inp_text, vbTab))
    Dim count_ec As Long
    count_ec = UBound(Split(inp_text, "？"))
    Dim clean_text As String
    clean_text = Replace(inp_text, "　", "")
    clean_text = Replace(clean_text, "「", "")
    clean_text = Replace(clean_text, "」", "")
    GetWithoutLB = Len(clean_text) - count_lb - count_tb
End Function

Private Sub ShowDetailResult(ByRef inp_text As String)

    Dim count_qm As Long: count_qm = UBound(Split(inp_text, "？"))
    Dim count_em As Long: count_em = UBound(Split(inp_text, "！"))
    Dim count_as As Long: count_as = UBound(Split(inp_text, "＊"))
    Dim count_el As Long: count_el = UBound(Split(inp_text, "……"))
    Dim count_p As Long: count_p = UBound(Split(inp_text, "。"))
    Dim count_c As Long: count_c = UBound(Split(inp_text, "、"))
    Dim test As Long: test = UBound(Split(inp_text, "」"))
    On Error Resume Next
        MsgBox "------------------------" & vbCrLf & _
                "疑問符(？)  : " & count_qm & vbCrLf & _
                "感嘆符(！)  : " & count_em & vbCrLf & _
                "三点(……)    : " & count_el & vbCrLf & _
                "＊区切り     : " & count_as / 3 & vbCrLf & _
                "句点、読点　　　 : " & count_c & " , " & count_p & vbCrLf & _
                "平均一文文字数　: " & Format(Round((GetWithoutLB(inp_text) - count_p) / (count_p + test), 2), "#.00") & vbCrLf & _
                "------------------------" _
                , vbOKOnly, "詳細情報"
        Err.Clear
    On Error GoTo 0

End Sub

Sub CountCharactersInBrackets()

    SHOULD_STOP = 0
    SHOULD_OPEN_UF = 1

    Dim rng As Range
    Dim startPos As Long
    Dim endPos As Long: endPos = 0
    Dim dialogue_count As Long
    Dim selection_text As String
    Dim sumstr As Long
    
    ' 選択範囲を取得
    Set rng = Selection.Range
    selection_text = rng.text
    If selection_text = "" Then
        MsgBox "カウントする範囲を選択してください", vbExclamation + vbOKOnly, "エラー"
        Exit Sub
    End If
    
    Dim all_text_count As Long: all_text_count = GetWithoutLB(selection_text)
    
    dialogue_count = 0
    startPos = InStr(1, selection_text, "「") ' 開始文字を検索
    
    Dim frm As New UserFormPB
    frm.PBInit ("会話文検索中")
    Call frm.ProgressBar(startPos, Len(selection_text))
    
    Dim SText() As Variant
    Dim JText() As Variant
    Dim i As Long: i = 0
    
    ' 開始文字が見つかる限りループ
    Do While startPos > 0
        DoEvents
        ReDim Preserve JText(i + 1)
        JText(i) = Mid(selection_text, endPos + 1, startPos - endPos - 1) '字の文を配列に格納
        endPos = InStr(startPos, selection_text, "」") ' 終了文字を検索
        
        If endPos > 0 Then  ' 終了文字が見つかった場合
            dialogue_count = dialogue_count + (endPos - startPos - 1) ' 中の文字数を加算
            ReDim Preserve SText(i + 1)
            SText(i) = Mid(selection_text, startPos, endPos - startPos + 1) '会話文を配列に格納
            startPos = InStr(endPos + 1, selection_text, "「") ' 次の開始文字を検索
            
            Call frm.ProgressBar(startPos, Len(selection_text))
            If frm.SHOULD_STOP_MAIN Then Exit Sub
            
        Else  ' 終了文字が見つからない場合
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
    ReDim Preserve JText(i + 1)
    JText(i) = Right(selection_text, Len(selection_text) - endPos) '最後の字の文を配列に格納
    
    
    SHOULD_OPEN_UF = 0
    Unload frm
    
    ' 結果を表示
    MsgBox "会話の文字数: " & dialogue_count & vbCrLf & _
            "全体の文字数: " & all_text_count & vbCrLf & _
            "会話文の割合: " & Format(Round(dialogue_count / all_text_count * 100, 2), "00.00") & "%" _
            , vbOKOnly, "検索結果"
    
    ShowDetailResult (selection_text)
    
    Application.ScreenUpdating = False
    Dim l As Long: l = 0
    Dim show_str As String
    ReDim Preserve SText(i + 1)
    For l = 0 To i
        show_str = show_str & SText(l) & vbCr & vbCr
    Next l
    With UserForm2
        .Show vbModeless
        .Caption = "会話文のみ"
        .CommandButton1.Caption = "ジャンプ"
        .TextBox1.Value = show_str
        DoEvents
        .TextBox1.SelStart = 0
        .Hide
    End With
    
    show_str = ""
    For l = 0 To i
        show_str = show_str & JText(l) & vbCr & vbCr
    Next l
    With UserForm3
        .Show vbModeless
        .Caption = "地の文のみ"
        .CommandButton1.Caption = "ジャンプ"
        .TextBox1.Value = show_str
        DoEvents
        .TextBox1.SelStart = 0
        .Hide
    End With
    Application.ScreenUpdating = True
    
    UserForm4.Show vbModeless
    
End Sub





