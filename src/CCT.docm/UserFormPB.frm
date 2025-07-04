VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPB 
   Caption         =   "UserForm1"
   ClientHeight    =   1000
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9860.001
   OleObjectBlob   =   "UserFormPB.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserFormPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public SHOULD_STOP_MAIN As Boolean

Private m_max_width As Double
Private myFrame1 As MSForms.Frame
Private myLabel1 As MSForms.Label
Private myLabel2 As MSForms.Label

Private Sub UserForm_Initialize()

    Set myFrame1 = Me.Controls.Add("Forms.Frame.1", "Frame1", True)
    Set myLabel1 = myFrame1.Controls.Add("Forms.Label.1", "Label1", True)
    Set myLabel2 = Me.Controls.Add("Forms.Label.1", "Label2", True)
    
    With myFrame1
        .Caption = ""
        .Left = 6
        .Top = 6
        .Width = 480
        .Height = 26
    End With

    With myLabel1
        .Caption = ""
        .Left = 0
        .Top = 0
        .Width = 0
        .Height = 22
        .BackColor = RGB(0, 0, 204)
    End With

    With myLabel2
        .Left = 426
        .Top = 36
        .Width = 60
        .Height = 12
        .TextAlign = fmTextAlignRight
    End With
    
    SHOULD_STOP_MAIN = False

    Me.Show vbModeless
        
End Sub


Public Sub PBInit(ByVal title As String)
    Me.Caption = title
    m_max_width = myFrame1.Width
End Sub


Public Sub ProgressBar(ByVal par_number As Double, ByVal max_index As Double)
    
    Dim pre_number As Double
    pre_number = m_max_width * (par_number / max_index)
    
    myLabel1.Width = pre_number
    myLabel2.Caption = Format(Round(100 * par_number / max_index, 2), "###.00") & "%"
    
    If par_number = max_index Then Unload Me
    
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If MsgBox("処理を中断しますか？", vbCritical + vbYesNo, "動作確認") = vbYes Then
            SHOULD_STOP_MAIN = True
            Cancel = True
            Me.Hide
        Else
            SHOULD_STOP_MAIN = False
            Cancel = True
        End If
    End If
End Sub






