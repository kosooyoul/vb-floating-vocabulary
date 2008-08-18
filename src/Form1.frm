VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "일본어 단어장"
   ClientHeight    =   6150
   ClientLeft      =   6450
   ClientTop       =   6930
   ClientWidth     =   7935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7935
   StartUpPosition =   1  '소유자 가운데
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox Check2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2880
      TabIndex        =   4
      Top             =   2250
      Value           =   1  '확인
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   2880
      TabIndex        =   3
      Top             =   860
      Value           =   1  '확인
      Width           =   200
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   5520
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   2880
      MaxLength       =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Text            =   "Form1.frx":08CA
      Top             =   1040
      Width           =   5055
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  '평면
      Height          =   2190
      Left            =   2880
      Pattern         =   "*.ajw"
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   2760
      TabIndex        =   6
      Text            =   "Word_2007_5_21.ajw"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   2880
      MaxLength       =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Text            =   "Form1.frx":08D2
      Top             =   0
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '평면
      Height          =   6150
      ItemData        =   "Form1.frx":08D9
      Left            =   0
      List            =   "Form1.frx":08DB
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu File 
      Caption         =   "파일(&F)"
      Begin VB.Menu NewWords 
         Caption         =   "새로쓰기(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu CountinueWords 
         Caption         =   "이어쓰기(&A)"
         Shortcut        =   ^A
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
   End
   Begin VB.Menu FocusO 
      Caption         =   "단축(&F)"
      Begin VB.Menu ListF 
         Caption         =   "단어목록 포커스(&L)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu WordsF 
         Caption         =   "단어쓰기 포커스(&W)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu De 
         Caption         =   "설명쓰기 포커스(&D)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu b4 
         Caption         =   "-"
      End
      Begin VB.Menu Addee 
         Caption         =   "삽입하기 단축(&A)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Delee 
         Caption         =   "삭제하기 단축(&D)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu Saveee 
         Caption         =   "저장하기 단축(&S)"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Intro 
      Caption         =   "정보(&I)"
      Begin VB.Menu Introduction 
         Caption         =   "프로그램정보(&I)"
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu Site 
         Caption         =   "http://www.ahyane.net"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Addee_Click()
Call AddWords_Click
End Sub

Private Sub AddWords_Click()
List1.AddItem (Text1.Text + "     " + Text6.Text)
End Sub

Private Sub Save()
    Dim filenumber As Integer '파일번호를 위한 변수
    Dim filename As String '파일이름을 위한 변수
    Dim i As Integer
    
    Text3.Text = ""
    
    For i = 0 To List1.ListCount
    Text3.Text = Text3.Text + List1.List(i) + vbCrLf
    
    Next i
    filename = File1.Path & "\" & Text5.Text
    filenumber = FreeFile '사용가능한 파일 번호를 구하고
    '저장 모드로 파일을 읽어 온다.
    Open filename For Output As filenumber
    '텍스트 박스의 내용으로 파일을 덮어씌운다.
    Print #filenumber, Text3.Text
    Close filenumber '파일을 닫는다.
End Sub

Private Sub Load()
    Dim filenumber As Integer '파일번호
    Dim filename As String '파일이름
    Dim ftemp As String '파일내용
    On Error Resume Next
    filename = File1.Path & "\" & Text5.Text
    filenumber = FreeFile '사용가능한 파일번호를 구하고
    '파일을 Input 모드(읽기 전용)로 연다.
    Open filename For Input As filenumber

    Do Until EOF(filenumber)
        '줄단위로 파일 끝가지 ftemp 라는 변수로 읽어 들인다.
        Line Input #filenumber, ftemp
        '다시 텍스트 박스로 줄 단위(vbcrlf)로 변수의 내용을 읽어 들인다.
        If ftemp <> "" Then List1.AddItem (ftemp)
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

File1.Refresh
End Sub

Private Sub CountinueWords_Click()
Call Load
End Sub

Private Sub De_Click()
Text6.SetFocus
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Delee_Click()
Call DeleteWords_Click
End Sub

Private Sub DeleteWords_Click()
On Error Resume Next
If List1.ListCount > 0 Then
List1.RemoveItem (List1.ListIndex)
End If
End Sub

Private Sub ExitProgram_Click()
End
End Sub

Private Sub File1_Click()
List1.Clear
Text5.Text = File1.filename
Call Load

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
Text5.Text = "Word_" & Year(Now) & "_" & Month(Now) & "_" & Day(Now) & ".ajw"
'Call Load
File1.Path = CurDir()
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

Private Sub ListF_Click()
List1.SetFocus
End Sub

Private Sub NewWords_Click()
List1.Clear
End Sub

Private Sub Saveee_Click()
Call Save
End Sub

Private Sub SaveWords_Click()
Call Save
File1.Refresh
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 1 Then
Text3.Text = Mid(Text1.Text, Text1.SelStart - 1, 2)
If Left(Text3.Text, 1) = "k" Then
    If Text3.Text = "ka" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "か"
    ElseIf Text3.Text = "ki" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "き"
    ElseIf Text3.Text = "ku" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "く"
    ElseIf Text3.Text = "ke" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "け"
    ElseIf Text3.Text = "ko" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "こ"
    End If
ElseIf Left(Text3.Text, 1) = "g" Then
    If Text3.Text = "ga" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "が"
    ElseIf Text3.Text = "gi" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぎ"
    ElseIf Text3.Text = "gu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぐ"
    ElseIf Text3.Text = "ge" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "げ"
    ElseIf Text3.Text = "go" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ご"
    End If
ElseIf Left(Text3.Text, 1) = "s" Then
    If Text3.Text = "sa" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "さ"
    ElseIf Text3.Text = "si" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "し"
    ElseIf Text3.Text = "su" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "す"
    ElseIf Text3.Text = "se" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "せ"
    ElseIf Text3.Text = "so" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "そ"
    End If
ElseIf Left(Text3.Text, 1) = "j" Then
    If Text3.Text = "ja" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ざ"
    ElseIf Text3.Text = "ji" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "じ"
    ElseIf Text3.Text = "ju" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ず"
    ElseIf Text3.Text = "je" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぜ"
    ElseIf Text3.Text = "jo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぞ"
    End If
ElseIf Left(Text3.Text, 1) = "t" Then
    If Text3.Text = "ta" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "た"
    ElseIf Text3.Text = "ti" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ち"
    ElseIf Text3.Text = "tu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "つ"
    ElseIf Text3.Text = "te" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "て"
    ElseIf Text3.Text = "to" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "と"
    ElseIf Text3.Text = "tt" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "っ"
    End If
ElseIf Left(Text3.Text, 1) = "d" Then
    If Text3.Text = "da" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "だ"
    ElseIf Text3.Text = "di" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぢ"
    ElseIf Text3.Text = "du" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "づ"
    ElseIf Text3.Text = "de" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "で"
    ElseIf Text3.Text = "do" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ど"
    End If
ElseIf Left(Text3.Text, 1) = "n" Then ''''''''''''''''''''''''''''''''''''
    If Text3.Text = "na" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "な"
    ElseIf Text3.Text = "ni" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "に"
    ElseIf Text3.Text = "nu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぬ"
    ElseIf Text3.Text = "ne" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ね"
    ElseIf Text3.Text = "no" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "の"
    ElseIf Text3.Text = "ng" Or Text3.Text = "nn" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ん"
    End If
ElseIf Left(Text3.Text, 1) = "h" Then
    If Text3.Text = "ha" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "は"
    ElseIf Text3.Text = "hi" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ひ"
    ElseIf Text3.Text = "hu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ふ"
    ElseIf Text3.Text = "he" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "へ"
    ElseIf Text3.Text = "ho" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ほ"
    End If
ElseIf Left(Text3.Text, 1) = "b" Then
    If Text3.Text = "ba" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ば"
    ElseIf Text3.Text = "bi" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "び"
    ElseIf Text3.Text = "bu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぶ"
    ElseIf Text3.Text = "be" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "べ"
    ElseIf Text3.Text = "bo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぼ"
    End If
ElseIf Left(Text3.Text, 1) = "p" Or Left(Text3.Text, 1) = "f" Then
    If Text3.Text = "pa" Or Text3.Text = "fa" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぱ"
    ElseIf Text3.Text = "pi" Or Text3.Text = "fi" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぴ"
    ElseIf Text3.Text = "pu" Or Text3.Text = "fu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぷ"
    ElseIf Text3.Text = "pe" Or Text3.Text = "fe" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぺ"
    ElseIf Text3.Text = "po" Or Text3.Text = "fo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ぽ"
    End If
ElseIf Left(Text3.Text, 1) = "m" Then
    If Text3.Text = "ma" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ま"
    ElseIf Text3.Text = "mi" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "み"
    ElseIf Text3.Text = "mu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "む"
    ElseIf Text3.Text = "me" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "め"
    ElseIf Text3.Text = "mo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "も"
    End If
ElseIf Left(Text3.Text, 1) = "r" Then
    If Text3.Text = "ra" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ら"
    ElseIf Text3.Text = "ri" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "り"
    ElseIf Text3.Text = "ru" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "る"
    ElseIf Text3.Text = "re" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "れ"
    ElseIf Text3.Text = "ro" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ろ"
    End If
ElseIf Left(Text3.Text, 1) = "y" Then
    If Text3.Text = "ya" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "や"
    ElseIf Text3.Text = "yu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ゆ"
    ElseIf Text3.Text = "yo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "よ"
    End If
ElseIf Left(Text3.Text, 1) = "i" Then
    If Text3.Text = "ia" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ゃ"
    ElseIf Text3.Text = "iu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ゅ"
    ElseIf Text3.Text = "io" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ょ"
    ElseIf Text3.Text = "ii" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "い"
    End If
ElseIf Left(Text3.Text, 1) = "w" Then
    If Text3.Text = "wa" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "わ"
    ElseIf Text3.Text = "ww" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "ゎ"
    ElseIf Text3.Text = "wo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "を"
    End If
Else
    If Text3.Text = "aa" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "あ"
    ElseIf Text3.Text = "uu" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "う"
    ElseIf Text3.Text = "ee" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "え"
    ElseIf Text3.Text = "oo" Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "お"
    ElseIf Text3.Text = ".." Then
    Text1.SelStart = Text1.SelStart - 2
    Text1.SelLength = 2
    Text1.SelText = "。"
    End If
End If
End If




End Sub



Private Sub WordsF_Click()
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
