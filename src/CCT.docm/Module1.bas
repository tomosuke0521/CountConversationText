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
    count_ec = UBound(Split(inp_text, "�H"))
    Dim clean_text As String
    clean_text = Replace(inp_text, "�@", "")
    clean_text = Replace(clean_text, "�u", "")
    clean_text = Replace(clean_text, "�v", "")
    GetWithoutLB = Len(clean_text) - count_lb - count_tb
End Function

Private Sub ShowDetailResult(ByRef inp_text As String)

    Dim count_qm As Long: count_qm = UBound(Split(inp_text, "�H"))
    Dim count_em As Long: count_em = UBound(Split(inp_text, "�I"))
    Dim count_as As Long: count_as = UBound(Split(inp_text, "��"))
    Dim count_el As Long: count_el = UBound(Split(inp_text, "�c�c"))
    Dim count_p As Long: count_p = UBound(Split(inp_text, "�B"))
    Dim count_c As Long: count_c = UBound(Split(inp_text, "�A"))
    Dim test As Long: test = UBound(Split(inp_text, "�v"))
    On Error Resume Next
        MsgBox "------------------------" & vbCrLf & _
                "�^�╄(�H)  : " & count_qm & vbCrLf & _
                "���Q��(�I)  : " & count_em & vbCrLf & _
                "�O�_(�c�c)    : " & count_el & vbCrLf & _
                "����؂�     : " & count_as / 3 & vbCrLf & _
                "��_�A�Ǔ_�@�@�@ : " & count_c & " , " & count_p & vbCrLf & _
                "���ψꕶ�������@: " & Format(Round((GetWithoutLB(inp_text) - count_p) / (count_p + test), 2), "#.00") & vbCrLf & _
                "------------------------" _
                , vbOKOnly, "�ڍ׏��"
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
    
    ' �I��͈͂��擾
    Set rng = Selection.Range
    selection_text = rng.text
    If selection_text = "" Then
        MsgBox "�J�E���g����͈͂�I�����Ă�������", vbExclamation + vbOKOnly, "�G���["
        Exit Sub
    End If
    
    Dim all_text_count As Long: all_text_count = GetWithoutLB(selection_text)
    
    dialogue_count = 0
    startPos = InStr(1, selection_text, "�u") ' �J�n����������
    
    Dim frm As New UserFormPB
    frm.PBInit ("��b��������")
    Call frm.ProgressBar(startPos, Len(selection_text))
    
    Dim SText() As Variant
    Dim JText() As Variant
    Dim i As Long: i = 0
    
    ' �J�n��������������胋�[�v
    Do While startPos > 0
        DoEvents
        ReDim Preserve JText(i + 1)
        JText(i) = Mid(selection_text, endPos + 1, startPos - endPos - 1) '���̕���z��Ɋi�[
        endPos = InStr(startPos, selection_text, "�v") ' �I������������
        
        If endPos > 0 Then  ' �I�����������������ꍇ
            dialogue_count = dialogue_count + (endPos - startPos - 1) ' ���̕����������Z
            ReDim Preserve SText(i + 1)
            SText(i) = Mid(selection_text, startPos, endPos - startPos + 1) '��b����z��Ɋi�[
            startPos = InStr(endPos + 1, selection_text, "�u") ' ���̊J�n����������
            
            Call frm.ProgressBar(startPos, Len(selection_text))
            If frm.SHOULD_STOP_MAIN Then Exit Sub
            
        Else  ' �I��������������Ȃ��ꍇ
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
    ReDim Preserve JText(i + 1)
    JText(i) = Right(selection_text, Len(selection_text) - endPos) '�Ō�̎��̕���z��Ɋi�[
    
    
    SHOULD_OPEN_UF = 0
    Unload frm
    
    ' ���ʂ�\��
    MsgBox "��b�̕�����: " & dialogue_count & vbCrLf & _
            "�S�̂̕�����: " & all_text_count & vbCrLf & _
            "��b���̊���: " & Format(Round(dialogue_count / all_text_count * 100, 2), "00.00") & "%" _
            , vbOKOnly, "��������"
    
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
        .Caption = "��b���̂�"
        .CommandButton1.Caption = "�W�����v"
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
        .Caption = "�n�̕��̂�"
        .CommandButton1.Caption = "�W�����v"
        .TextBox1.Value = show_str
        DoEvents
        .TextBox1.SelStart = 0
        .Hide
    End With
    Application.ScreenUpdating = True
    
    UserForm4.Show vbModeless
    
End Sub





