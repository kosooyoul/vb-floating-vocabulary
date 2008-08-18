VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   Icon            =   "Form133.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture0 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "Form133.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   8670
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1920
         Top             =   0
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   660
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Height          =   480
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   200
      End
      Begin VB.Label Picture4 
         BackStyle       =   0  '투명
         Height          =   300
         Left            =   7920
         TabIndex        =   5
         Top             =   90
         Width           =   300
      End
      Begin VB.Label Picture3 
         BackStyle       =   0  '투명
         Height          =   300
         Left            =   8280
         TabIndex        =   4
         Top             =   90
         Width           =   300
      End
      Begin VB.Label Picture2 
         BackStyle       =   0  '투명
         Height          =   300
         Left            =   7480
         TabIndex        =   3
         Top             =   90
         Width           =   300
      End
      Begin VB.Label Picture1 
         BackStyle       =   0  '투명
         Height          =   300
         Left            =   270
         TabIndex        =   2
         Top             =   90
         Width           =   300
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WordCounter As Integer

Private Sub Form_Load()
    Me.Picture0.Move 0, 0
    Me.Left = 0
    Me.Top = 0
    WordCounter = 0

End Sub

Private Sub Label1_Click()
    End
End Sub

Private Sub Picture0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton                                               '폼이동
            Call ReleaseCapture
            Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End Select
End Sub
Private Sub Picture1_Click()
    Timer1.Enabled = False
    WordCounter = WordCounter - 2
    Call Timer1_Timer
    Timer1.Enabled = True
End Sub

Private Sub Picture2_Click()
    Timer1.Enabled = False
    Call Timer1_Timer
    Timer1.Enabled = True
End Sub

Private Sub Picture3_Click()
    Dim Temp As Integer
    
    On Error Resume Next
    
    Temp = InputBox("시간간격", "시간간격 설정", Timer1.Interval)
    
    Timer1.Interval = Val(Temp)
    
End Sub

Private Sub Picture4_Click()
    Timer1.Enabled = Not (Timer1.Enabled)
End Sub

Private Sub Text1_DblClick()
    Form11.Show
End Sub

Private Sub Timer1_Timer()
    Text1.Text = Form11.List1.List(WordCounter)
    
    If WordCounter = Form11.List1.ListCount Then
        WordCounter = 0
    Else
        WordCounter = WordCounter + 1
    End If
    
End Sub
