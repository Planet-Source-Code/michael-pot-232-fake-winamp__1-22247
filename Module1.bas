Attribute VB_Name = "Module1"
'       -------------
'        FAKE WINAMP
'       -------------
'     a skinning project...
'
' This project demonstrates the use of Bitblt to draw
' Pictures from a winamp skin directory to pre-defined places
' on the window.
'
' To Use with other winamp skins: Goto your winamp directory
' select a wsz file and open it with Winzip, extract it to a
' directory then in Form_load type: LoadSkin "C:\MySkinDir\"



Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


'Global variables
Public Volume As Long, Pan As Long ' Volume Max = 27, Pan Centre = 27; Max = 54
Public Pos As Single, Kbps As String, kHz As String, CurrentSong As String
Public Shuffle As Boolean, Repeat As Boolean, Pl As Boolean, Eq As Boolean, Stereo As Integer '0 - none, 1 - mono, 2 - stereo
Public PlayState As Integer '0 - Play, 1 - Pause, 2 - Stop

Public CutS As String, CutI As Long 'Current song string

'Integer switches: 0 - off, 1 - on.

'Each variable with a 'b' prefix can either be 0 or 1 or in special cases 2 or above.
'These variables are linked to buttons and sliders that make up the winamp-like display
'for example, BPlay is linked to the play button, this can be 0 - unpressed, or 1 - pressed.
'BPos is linked to the main slider and determins weather the slider button is down or not.

Public BShuf As Integer, BRep As Integer, BEq As Integer, BPl As Integer
Public BBack As Integer, BPlay As Integer, BPause As Integer, BStop As Integer, BNext As Integer, BEject As Integer
Public BPos As Integer, BVol As Integer, BPan As Integer, BTitle As Integer
Public BMin As Integer, BUp As Integer, BExit As Integer, BTit As Integer 'titlebar buttons
Public BSideBar As Integer, Numbers As Integer
Public Function LoadSkin(Path As String)
On Error Resume Next
'Load bitmaps

If Path = "" Then GoTo SKip

frmMain.picBack.Picture = LoadPicture(Path & "MAIN.BMP")
frmMain.picPP.Picture = LoadPicture(Path & "PLAYPAUS.BMP")
frmMain.picButs.Picture = LoadPicture(Path & "CBUTTONS.BMP")
frmMain.picTitle.Picture = LoadPicture(Path & "TITLEBAR.BMP")
If Dir(Path & "NUMBERS.BMP") = "" Then
    frmMain.picNums.Picture = LoadPicture(Path & "NUMS_EX.BMP")
Else
    frmMain.picNums.Picture = LoadPicture(Path & "NUMBERS.BMP")
End If
frmMain.picBals.Picture = LoadPicture(Path & "BALANCE.BMP")
frmMain.picMS.Picture = LoadPicture(Path & "MONOSTER.BMP")
frmMain.picBar.Picture = LoadPicture(Path & "POSBAR.BMP")
frmMain.picSR.Picture = LoadPicture(Path & "SHUFREP.BMP")
frmMain.picText.Picture = LoadPicture(Path & "FONT.BMP")
frmMain.picVol.Picture = LoadPicture(Path & "VOLUME.BMP")

SKip:

'Resize to winamp size
frmMain.Width = frmMain.picBack.Width * 15
frmMain.Height = frmMain.picBack.Height * 15

' Song Title
' Must be 31 characters long...

'from here                           to here
'      |                               |
CutS = "Fake Winamp by Michael Pote ** "
Kbps = "128"
kHz = "44"

PaintWinamp
End Function

Public Sub PaintWinamp()
Dim N As String
'This sub actually draws the skin onto the form

N = Format(Numbers, "0000") 'Give the numbers leading zeros

With frmMain

    'Start with the background...
    BitBlt .hdc, 0, 0, .picBack.Width, .picBack.Height, .picBack.hdc, 0, 0, SRCCOPY
    
    'Next the Title Bar...
    BitBlt .hdc, 0, 0, .picBack.Width, 14, .picTitle.hdc, 27, BTitle * 15, SRCCOPY
    
    'And it's buttons
    BitBlt .hdc, 6, 3, 9, 9, .picTitle.hdc, 0, BTit * 9, SRCCOPY
    BitBlt .hdc, 245, 3, 9, 9, .picTitle.hdc, 9, BMin * 9, SRCCOPY
    BitBlt .hdc, 254, 3, 9, 9, .picTitle.hdc, BUp * 9, 18, SRCCOPY
    BitBlt .hdc, 265, 3, 9, 9, .picTitle.hdc, 18, BExit * 9, SRCCOPY
    
    'Sidebar
    If BSideBar <= 0 Then
        BitBlt .hdc, 10, 22, 8, 43, .picTitle.hdc, 304, 0, SRCCOPY
    Else
        BitBlt .hdc, 10, 22, 8, 43, .picTitle.hdc, 304 + ((BSideBar - 1) * 8), 44, SRCCOPY
    End If
    
    
    'Numbers
    BitBlt .hdc, 48, 26, 9, 13, .picNums.hdc, Int(Mid(N, 1, 1)) * 9, 0, SRCCOPY
    BitBlt .hdc, 60, 26, 9, 13, .picNums.hdc, Int(Mid(N, 2, 1)) * 9, 0, SRCCOPY
    BitBlt .hdc, 78, 26, 9, 13, .picNums.hdc, Int(Mid(N, 3, 1)) * 9, 0, SRCCOPY
    BitBlt .hdc, 90, 26, 9, 13, .picNums.hdc, Int(Mid(N, 4, 1)) * 9, 0, SRCCOPY
    
    'Volume
    BitBlt .hdc, 107, 57, 68, 13, .picVol.hdc, 0, Volume * 15, SRCCOPY
    BitBlt .hdc, 107 + ((Volume / 27) * 50), 58, 14, 11, .picVol.hdc, (1 - BVol) * 15, 422, SRCCOPY
    
    'Panning control
    If Pan <= 27 Then
        BitBlt .hdc, 177, 57, 38, 13, .picBals.hdc, 9, (27 - Pan) * 15, SRCCOPY
    Else
        BitBlt .hdc, 177, 57, 38, 13, .picBals.hdc, 9, (Pan - 27) * 15, SRCCOPY
    End If
    'Pan button
    BitBlt .hdc, 177 + ((Pan / 54) * 20), 58, 14, 11, .picBals.hdc, (1 - BPan) * 15, 422, SRCCOPY
    
    'Play Pause
    BitBlt .hdc, 27, 28, 9, 9, .picPP.hdc, PlayState * 9, 0, SRCCOPY
    If PlayState = 0 Then
        If Pos >= 0.9 Then 'If song is near end, show small stop sign..
            BitBlt .hdc, 24, 28, 3, 3, .picPP.hdc, 39, 0, SRCCOPY
            BitBlt .hdc, 24, 34, 3, 3, .picPP.hdc, 39, 6, SRCCOPY
        Else
            BitBlt .hdc, 24, 28, 3, 3, .picPP.hdc, 36, 0, SRCCOPY
            BitBlt .hdc, 24, 34, 3, 3, .picPP.hdc, 36, 6, SRCCOPY
        End If
    End If
    
    'Shuffle Repeat, Equliser & Play-list
    BitBlt .hdc, 210, 89, 29, 15, .picSR.hdc, 0, BRep * 15 + (30 * IIf(Repeat = True, 1, 0)), SRCCOPY
    BitBlt .hdc, 164, 89, 46, 15, .picSR.hdc, 29, BShuf * 15 + (30 * IIf(Shuffle = True, 1, 0)), SRCCOPY
        
    BitBlt .hdc, 219, 58, 23, 12, .picSR.hdc, BEq * 46, IIf(Eq = False, 61, 73), SRCCOPY
    BitBlt .hdc, 242, 58, 23, 12, .picSR.hdc, 23 + BPl * 46, IIf(Pl = False, 61, 73), SRCCOPY
        
    'Mono/Stereo
    BitBlt .hdc, 239, 41, 29, 12, .picMS.hdc, 0, IIf(Stereo = 2, 0, 12), SRCCOPY
    BitBlt .hdc, 212, 41, 29, 12, .picMS.hdc, 29, IIf(Stereo = 1, 0, 12), SRCCOPY
    
    
    'Buttons
    'Back
    BitBlt .hdc, 16, 88, 23, 18, .picButs.hdc, 0, BBack * 18, SRCCOPY
    'Play
    BitBlt .hdc, 39, 88, 23, 18, .picButs.hdc, 23, BPlay * 18, SRCCOPY
    'Pause
    BitBlt .hdc, 62, 88, 23, 18, .picButs.hdc, 46, BPause * 18, SRCCOPY
    'Stop
    BitBlt .hdc, 85, 88, 23, 18, .picButs.hdc, 69, BStop * 18, SRCCOPY
    'Next
    BitBlt .hdc, 108, 88, 22, 18, .picButs.hdc, 92, BNext * 18, SRCCOPY
    'Eject
    BitBlt .hdc, 136, 89, 22, 16, .picButs.hdc, 114, BEject * 16, SRCCOPY

    'Play Bar
    BitBlt .hdc, 16, 72, 249, 10, .picBar.hdc, 0, 0, SRCCOPY
    'Play bar button
    BitBlt .hdc, 16 + (223 * Pos), 72, 29, 10, .picBar.hdc, 249 + (BPos * 29), 0, SRCCOPY

    WriteText 111, 43, Kbps 'Draw text on the window
    WriteText 156, 43, kHz
    WriteText 111, 27, CutS

.Refresh
End With

End Sub

Public Function WriteText(X As Long, Y As Long, Text As String)
Dim I As Long, L As String, LX As Long
Text = UCase(Text)
With frmMain
    LX = X
    For I = 1 To Len(Text)
        L = Mid(UCase(Text), I, 1)
        If L = " " Then GoTo SKip
        BitBlt .hdc, LX, Y, 5, 6, .picText.hdc, FindTextX(L), FindTextY(L), SRCCOPY
SKip:
        LX = LX + 5
    Next I
End With
End Function

Function FindTextX(L As String) As Long
If L Like "[A-Z]" Then FindTextX = (Asc(L) - 65) * 5: Exit Function
If L Like "#" Then FindTextX = (Asc(L) - 48) * 5: Exit Function

'Hard coded character position
Select Case L
Case """"
FindTextX = 130
Case "@"
FindTextX = 136
Case "_"
FindTextX = 90
Case "-"
FindTextX = 75
Case "+"
FindTextX = 95
Case "="
FindTextX = 60
Case "("
FindTextX = 65
Case ")"
FindTextX = 70
Case "'"
FindTextX = 80
Case "!"
FindTextX = 85
Case "'"
FindTextX = 80
Case "\"
FindTextX = 100
Case "/"
FindTextX = 105
Case "["
FindTextX = 110
Case "]"
FindTextX = 115
Case "^"
FindTextX = 120
Case "&"
FindTextX = 125
Case "%"
FindTextX = 130
Case ","
FindTextX = 135
Case "$"
FindTextX = 145
Case "#"
FindTextX = 150
Case "*"
FindTextX = 20

End Select
End Function

Function FindTextY(L As String) As Long
If L Like "[A-Z]" Or L = """" Or L = "@" Then FindTextY = 0: Exit Function
If L Like "#" Then FindTextY = 6: Exit Function
Select Case L
Case "*"
FindTextY = 12
Case Else
FindTextY = 6
End Select
End Function

Public Sub CutBar()
CutS = Mid(CutS, 2) & Mid(CutS, 1, 1)
End Sub
