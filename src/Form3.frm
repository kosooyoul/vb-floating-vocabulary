VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "단어찾기"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5385
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "찾기(&F)"
      Height          =   400
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "단어목록에서 위 단어가 있는지 찾기"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "닫기(&C)"
      Height          =   400
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "창 닫기"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5175
      Begin VB.CheckBox Check1 
         Caption         =   "ひらがな"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         ToolTipText     =   "입력시 히라가나를 자동입력할것인지 여부"
         Top             =   240
         Value           =   1  '확인
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         TabIndex        =   0
         ToolTipText     =   "검색할 단어 / 히라가나만 자동입력가능"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ことば :"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   660
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "찾기버튼 누름"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Check1.Value Then
If KeyAscii = 13 Then
    Call Command1_Click
ElseIf KeyAscii = 27 Then
    Call Command2_Click
ElseIf KeyAscii <> 8 Then
If Len(Text1.Text) > 0 Then
    Hira = Mid(Text1.Text, Text1.SelStart, 1) & Chr(KeyAscii)
    Text3.Text = Hira
    Call toHira

    Text1.SelStart = Text1.SelStart - 1
    Text1.SelLength = 1
    Text1.SelText = Hira
    KeyAscii = 0
End If
End If
End If
End Sub
