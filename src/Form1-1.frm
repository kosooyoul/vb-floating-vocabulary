VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "일본어 단어장"
   ClientHeight    =   3495
   ClientLeft      =   6330
   ClientTop       =   3510
   ClientWidth     =   6735
   Icon            =   "Form1-1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6735
   StartUpPosition =   1  '소유자 가운데
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "삭제(&D)"
         Height          =   400
         Left            =   5340
         TabIndex        =   9
         ToolTipText     =   "위 단어 목록에서 선택된 단어 삭제"
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "읽기(&O)"
         Height          =   400
         Left            =   1860
         TabIndex        =   7
         ToolTipText     =   "위 파일에 오른쪽 단어 목록을 저장"
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "저장(&S)"
         Height          =   400
         Left            =   720
         TabIndex        =   6
         ToolTipText     =   "오른쪽의 단어목록을 위 파일이름으로 저장"
         Top             =   2715
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "추가(&A)"
         Height          =   400
         Left            =   1860
         TabIndex        =   5
         ToolTipText     =   "오른쪽 목록에 추가"
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "見(1)"
         Height          =   200
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "일어단어 보이기"
         Top             =   2080
         Value           =   1  '확인
         Width           =   795
      End
      Begin VB.CheckBox Check2 
         Caption         =   "見(2)"
         Height          =   200
         Left            =   1020
         TabIndex        =   4
         ToolTipText     =   "설명 보이기"
         Top             =   2085
         Value           =   1  '확인
         Width           =   795
      End
      Begin VB.ListBox List1 
         Height          =   2040
         ItemData        =   "Form1-1.frx":058A
         Left            =   3240
         List            =   "Form1-1.frx":058C
         TabIndex        =   8
         ToolTipText     =   "저장할 단어 목록"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   680
         Left            =   120
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   0
         Text            =   "Form1-1.frx":058E
         ToolTipText     =   "히라가나만 자동입력가능"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   120
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Text            =   "Form1-1.frx":05A0
         ToolTipText     =   "해당 단어 설명"
         Top             =   1215
         Width           =   2895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   3120
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   3120
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "ことば LIST :"
         Height          =   180
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   3135
         X2              =   3135
         Y1              =   120
         Y2              =   3290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   3120
         X2              =   3120
         Y1              =   105
         Y2              =   3300
      End
      Begin VB.Label Label1 
         Caption         =   "ことば :"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Menu File 
      Caption         =   "파일(&F)"
      Begin VB.Menu CountinueWords 
         Caption         =   "읽어오기(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveWords 
         Caption         =   "저장하기(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProgram 
         Caption         =   "끝내기"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu EditWords 
      Caption         =   "편집(&E)"
      Begin VB.Menu AddWords 
         Caption         =   "삽입하기(&A)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu DeleteWords 
         Caption         =   "삭제하기(&X)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu Wfind 
         Caption         =   "단어찾기(&F)"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Intro 
      Caption         =   "정보(&I)"
      Begin VB.Menu Introduction 
         Caption         =   "프로그램정보(&I)"
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AddWords_Click()
Call Command1_Click
End Sub

Private Sub Save()
    Dim filenumber As Integer '파일번호를 위한 변수
    Dim filename As String '파일이름을 위한 변수
    Dim i As Integer
    
    Hira = ""
    
    For i = 0 To List1.ListCount    '저장할 단어들을 하나의 텍스트로 만듬
    Hira = Hira + List1.List(i) + vbCrLf
    
    Next i
    filename = App.Path & "\word.ahyane"
    filenumber = FreeFile '사용가능한 파일 번호를 구하고
    '저장 모드로 파일을 읽어 온다.
    Open filename For Output As filenumber
    '텍스트 박스의 내용으로 파일을 덮어씌운다.
    Print #filenumber, Hira
    Close filenumber '파일을 닫는다.
End Sub

Private Sub Load()
    Dim filenumber As Integer '파일번호
    Dim filename As String '파일이름
    Dim ftemp As String '파일내용
    On Error Resume Next
    filename = App.Path & "\word.ahyane"
    filenumber = FreeFile '사용가능한 파일번호를 구하고
    '파일을 Input 모드(읽기 전용)로 연다.
    Open filename For Input As filenumber

    Do Until EOF(filenumber)
        '줄단위로 파일 끝가지 ftemp 라는 변수로 읽어 들인다.
        Line Input #filenumber, ftemp
        '다시 텍스트 박스로 줄 단위(vbcrlf)로 변수의 내용을 읽어 들인다.
        If ftemp <> "" Then List1.AddItem (ftemp)   '줄단위로 로딩하면서 한줄씩 리스트에 삽입
    Loop

    Close filenumber '파일을 닫든다.
End Sub

Private Sub Check1_Click()
If Check1.Value Then
Text1.ForeColor = 0
Else
Text1.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Check2_Click()
If Check2.Value Then
Text6.ForeColor = 0
Else
Text6.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Command1_Click()
List1.AddItem (Text1.Text + "     " + Text6.Text)
End Sub

Private Sub Command2_Click()
Call Save
File1.Refresh
End Sub

Private Sub Command3_Click()
Call Load
End Sub

Private Sub Command4_Click()
On Error Resume Next
If List1.ListCount > 0 Then
List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub CountinueWords_Click()
Call Load
End Sub


Private Sub DeleteWords_Click()
Call Command4_Click
End Sub

Private Sub ExitProgram_Click()
End
End Sub


Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
List1.Clear
Text5.Text = File1.filename
Call Load
End Sub

Private Sub File1_Scroll()
List1.Clear
Text5.Text = File1.filename
Call Load
End Sub

Private Sub Form_Load()
'Text5.Text = "Word_" & Year(Now) & "_" & Month(Now) & "_" & Day(Now) & ".ajw"
Call Load
'File1.Path = CurDir()
'Text5.Text = File1.Path & "\" & Text5.Text
End Sub

Private Sub Form_Resize()
'List1.Height = Abs(Me.Height - 850)
'Text5.Top = Abs(List1.Height - 300)
'File1.Top = Text5.Top - 2300
'Text1.Width = Abs(Me.Width - 3100)
'Text6.Width = Abs(Text1.Width)
End Sub

Private Sub Introduction_Click()
Form2.Show vbModal
End Sub

Private Sub List1_Click()
Dim i As Integer
If List1.ListIndex >= 0 Then
Text1.Text = List1.List(List1.ListIndex)
If Len(Text1.Text) > 5 Then
For i = 1 To Len(Text1.Text) - 5
If Mid(Text1.Text, i, 5) = "     " Then
Text1.SelStart = i - 1
Text1.SelLength = Len(Text1.Text)
Text6.Text = Trim(Text1.SelText)
Text1.SelText = ""
End If
Next i
End If
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Call List1_Click
End Sub

Private Sub List1_Scroll()
Call List1_Click
End Sub

Private Sub NewWords_Click()
List1.Clear
End Sub


Private Sub SaveWords_Click()
Call Command2_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
If Len(Text1.Text) > 0 Then
    Hira = Mid(Text1.Text, Text1.SelStart, 1) & Chr(KeyAscii)
    Call toHira

    Text1.SelStart = Text1.SelStart - 1
    Text1.SelLength = 1
    Text1.SelText = Hira
    KeyAscii = 0
End If
End If
End Sub

Private Sub Wfind_Click()
Form3.Show vbModal
End Sub
