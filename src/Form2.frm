VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "프로그램 정보"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인(&O)"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "창 닫기"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "2007년 8월 20일"
      Height          =   180
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1290
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   480
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   300
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   720
      Picture         =   "Form2.frx":0395
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "첫번째 타입 단어장"
      Height          =   180
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1560
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3960
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "Made by Ahyane"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
