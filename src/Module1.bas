Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long          '폼이동에 대한 API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                                                                        '그림창에 대한 API
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1 '-1
Public Const HWND_NOTOPMOST = -2 '-2 창 보이기 순서
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const HTCAPTION = 2                                              '폼이동에 대한 변수
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const RGN_OR = 2

Public Hira As String
Dim Ing As Boolean

Sub Main()
    Load Form1
    
    SetWindowRgn Form1.hwnd, lGetRegion(Form1.Picture0, RGB(255, 0, 255)), True
    DeleteObject lGetRegion(Form1.Picture0, RGB(255, 0, 255))
    
    Form1.Show
    SetFormPosition Form1.hwnd, True                                    '인덱스 폼을 항상 위에 표시
End Sub


Public Sub SetFormPosition(hwnd As Long, TopPosition As Boolean)        '폼위치(Layer) 대한 함수
    If TopPosition Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
     Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
     End If
End Sub

'그림으로 폼을 표시 : 인덱스 폼
Public Function lGetRegion(pic As PictureBox, lBackColor As Long) As Long
    Dim lRgn As Long
    Dim lSkinRgn As Long
    Dim lStart As Long
    Dim lX As Long, lY As Long
    Dim lHeight As Long, lWidth As Long
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With pic
        lHeight = .Height / Screen.TwipsPerPixelY
        lWidth = .Width / Screen.TwipsPerPixelX
        For lX = 0 To lHeight - 1
            lY = 0
            Do While lY < lWidth
                Do While lY < lWidth And GetPixel(.hdc, lY, lX) = lBackColor
                    lY = lY + 1
                Loop
                If lY < lWidth Then
                    lStart = lY
                    Do While lY < lWidth And GetPixel(.hdc, lY, lX) <> lBackColor
                        lY = lY + 1
                    Loop
                    If lY > lWidth Then lY = lWidth
                    lRgn = CreateRectRgn(lStart, lX, lY, lX + 1)
                    CombineRgn lSkinRgn, lSkinRgn, lRgn, RGN_OR
                    DeleteObject lRgn
                End If
            Loop
        Next
    End With
    lGetRegion = lSkinRgn
End Function

Public Sub AlwaysOnTop(TheForm As Form, Toggle As Boolean) '창 표시 순서
    If Toggle = True Then
        SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
End Sub

Public Sub toHira()
If Left(Hira, 1) = "k" Then
    If Hira = "ka" Then
    Hira = "か"
    ElseIf Hira = "ki" Then
    Hira = "き"
    ElseIf Hira = "ku" Then
    Hira = "く"
    ElseIf Hira = "ke" Then
    Hira = "け"
    ElseIf Hira = "ko" Then
    Hira = "こ"
    End If
ElseIf Left(Hira, 1) = "g" Then
    If Hira = "ga" Then
    Hira = "が"
    ElseIf Hira = "gi" Then
    Hira = "ぎ"
    ElseIf Hira = "gu" Then
    Hira = "ぐ"
    ElseIf Hira = "ge" Then
    Hira = "げ"
    ElseIf Hira = "go" Then
    Hira = "ご"
    End If
ElseIf Left(Hira, 1) = "s" Then
    If Hira = "sa" Then
    Hira = "さ"
    ElseIf Hira = "si" Then
    Hira = "し"
    ElseIf Hira = "su" Then
    Hira = "す"
    ElseIf Hira = "se" Then
    Hira = "せ"
    ElseIf Hira = "so" Then
    Hira = "そ"
    End If
ElseIf Left(Hira, 1) = "j" Then
    If Hira = "ja" Then
    Hira = "ざ"
    ElseIf Hira = "ji" Then
    Hira = "じ"
    ElseIf Hira = "ju" Then
    Hira = "ず"
    ElseIf Hira = "je" Then
    Hira = "ぜ"
    ElseIf Hira = "jo" Then
    Hira = "ぞ"
    End If
ElseIf Left(Hira, 1) = "t" Then
    If Hira = "ta" Then
    Hira = "た"
    ElseIf Hira = "ti" Then
    Hira = "ち"
    ElseIf Hira = "tu" Then
    Hira = "つ"
    ElseIf Hira = "te" Then
    Hira = "て"
    ElseIf Hira = "to" Then
    Hira = "と"
    ElseIf Hira = "tt" Then
    Hira = "っ"
    End If
ElseIf Left(Hira, 1) = "d" Then
    If Hira = "da" Then
    Hira = "だ"
    ElseIf Hira = "di" Then
    Hira = "ぢ"
    ElseIf Hira = "du" Then
    Hira = "づ"
    ElseIf Hira = "de" Then
    Hira = "で"
    ElseIf Hira = "do" Then
    Hira = "ど"
    End If
ElseIf Left(Hira, 1) = "n" Then ''''''''''''''''''''''''''''''''''''
    If Hira = "na" Then
    Hira = "な"
    ElseIf Hira = "ni" Then
    Hira = "に"
    ElseIf Hira = "nu" Then
    Hira = "ぬ"
    ElseIf Hira = "ne" Then
    Hira = "ね"
    ElseIf Hira = "no" Then
    Hira = "の"
    ElseIf Hira = "ng" Or Hira = "nn" Then
    Hira = "ん"
    End If
ElseIf Left(Hira, 1) = "h" Then
    If Hira = "ha" Then
    Hira = "は"
    ElseIf Hira = "hi" Then
    Hira = "ひ"
    ElseIf Hira = "hu" Then
    Hira = "ふ"
    ElseIf Hira = "he" Then
    Hira = "へ"
    ElseIf Hira = "ho" Then
    Hira = "ほ"
    End If
ElseIf Left(Hira, 1) = "b" Then
    If Hira = "ba" Then
    Hira = "ば"
    ElseIf Hira = "bi" Then
    Hira = "び"
    ElseIf Hira = "bu" Then
    Hira = "ぶ"
    ElseIf Hira = "be" Then
    Hira = "べ"
    ElseIf Hira = "bo" Then
    Hira = "ぼ"
    End If
ElseIf Left(Hira, 1) = "p" Or Left(Hira, 1) = "f" Then
    If Hira = "pa" Or Hira = "fa" Then
    Hira = "ぱ"
    ElseIf Hira = "pi" Or Hira = "fi" Then
    Hira = "ぴ"
    ElseIf Hira = "pu" Or Hira = "fu" Then
    Hira = "ぷ"
    ElseIf Hira = "pe" Or Hira = "fe" Then
    Hira = "ぺ"
    ElseIf Hira = "po" Or Hira = "fo" Then
    Hira = "ぽ"
    End If
ElseIf Left(Hira, 1) = "m" Then
    If Hira = "ma" Then
    Hira = "ま"
    ElseIf Hira = "mi" Then
    Hira = "み"
    ElseIf Hira = "mu" Then
    Hira = "む"
    ElseIf Hira = "me" Then
    Hira = "め"
    ElseIf Hira = "mo" Then
    Hira = "も"
    End If
ElseIf Left(Hira, 1) = "r" Then
    If Hira = "ra" Then
    Hira = "ら"
    ElseIf Hira = "ri" Then
    Hira = "り"
    ElseIf Hira = "ru" Then
    Hira = "る"
    ElseIf Hira = "re" Then
    Hira = "れ"
    ElseIf Hira = "ro" Then
    Hira = "ろ"
    End If
ElseIf Left(Hira, 1) = "y" Then
    If Hira = "ya" Then
    Hira = "や"
    ElseIf Hira = "yu" Then
    Hira = "ゆ"
    ElseIf Hira = "yo" Then
    Hira = "よ"
    End If
ElseIf Left(Hira, 1) = "x" Then
    If Hira = "xa" Then
    Hira = "ゃ"
    ElseIf Hira = "xu" Then
    Hira = "ゅ"
    ElseIf Hira = "xo" Then
    Hira = "ょ"
    ElseIf Hira = "xw" Then
    Hira = "ゎ"
    ElseIf Hira = "xt" Then
    Hira = "っ"
    End If
ElseIf Left(Hira, 1) = "w" Then
    If Hira = "wa" Then
    Hira = "わ"
    ElseIf Hira = "wo" Then
    Hira = "を"
    End If
Else
    If Hira = "aa" Then
    Hira = "あ"
    ElseIf Hira = "ii" Then
    Hira = "い"
    ElseIf Hira = "uu" Then
    Hira = "う"
    ElseIf Hira = "ee" Then
    Hira = "え"
    ElseIf Hira = "oo" Then
    Hira = "お"
    ElseIf Hira = ".." Then
    Hira = "。"
    End If
End If
End Sub
