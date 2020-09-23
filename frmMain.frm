VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   810
   ClientTop       =   1815
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCol 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   60
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   11
      Top             =   3675
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.PictureBox picVol 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   6540
      Picture         =   "frmMain.frx":17772
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   10
      Top             =   -375
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picBals 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   5970
      Picture         =   "frmMain.frx":1A49A
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   405
      Top             =   705
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   1245
      Picture         =   "frmMain.frx":1CA5C
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   9
      Top             =   510
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox picSR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   3690
      Picture         =   "frmMain.frx":1E946
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   1140
      Picture         =   "frmMain.frx":2042E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   7
      Top             =   5460
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.PictureBox picMS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   60
      Picture         =   "frmMain.frx":20BEA
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   6
      Top             =   3045
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox picNums 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4230
      Picture         =   "frmMain.frx":214AC
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   4
      Top             =   3855
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   105
      Picture         =   "frmMain.frx":21CA6
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   3
      Top             =   1635
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.PictureBox picButs 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4230
      Picture         =   "frmMain.frx":258CC
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   2
      Top             =   4815
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.PictureBox picPP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   135
      Picture         =   "frmMain.frx":26E58
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   1
      Top             =   3465
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   0
      Picture         =   "frmMain.frx":273D0
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   3645
      Visible         =   0   'False
      Width           =   4125
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Down As Boolean, Mx, My, L, T, Col As Long
Private P As POINTAPI

Private Sub Form_Load()
'Read module comments on how to load your own winamp skins into the program.
LoadSkin "" 'Loads the default winzip skin
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = True
Col = ColDec(X, Y)
If Col = 0 Then
Mx = X
My = Y
Else
ButtonDown (Col)
Mx = -100
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim D
If Down Then

D = Col
Col = ColDec(X, Y) 'If the mouse is over a different button, depress
If Col <> D Then ButtonUp (0): ButtonDown (Col) 'the current one and press the new one.

If Col = 0 And Y < 24 And Not Mx = -100 Then 'User clicked the title bar and no buttons are selected.
 GetCursorPos P
 L = (P.X - Mx)
 T = (P.Y - My)

'Snapping routine, only works with the top and left corners of the screen.

 If L >= ((Screen.Width / 15) + picBack.Width) - 10 Then L = (Screen.Width / 15) + picBack.Width
 If L <= 10 Then L = 0
 If T <= 10 Then T = 0
 Left = L * 15
 Top = T * 15
 End If

Else
ButtonOver (Col)
End If
PaintWinamp
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = False
ButtonUp (Col)

End Sub

Private Sub Timer1_Timer()
'pos is expressed in percent: 0 - 0%, 1 - 100%, 0.5 = 50%
Pos = Pos + 0.01
If Pos >= 1 Then Pos = 0

Volume = Volume + 1
If Volume >= 28 Then Volume = 0

If Pos * 100 Mod 5 = 0 Then CutBar 'the Cutbar routine scrolls the song title

Pan = Pan + 1
If Pan >= 55 Then Pan = 0

Numbers = Numbers + 1 'The main song title numbers

If Pl = False Then GoTo SkipLights

Stereo = Stereo + 1 'None, Mono, Stereo
If Stereo = 3 Then Stereo = 0

BEq = Int(Rnd * 2)
BPl = Int(Rnd * 2)
BRep = Int(Rnd * 2)
BShuf = Int(Rnd * 2)

PlayState = Int(Rnd * 3)

BBack = Int(Rnd * 2)
BPlay = Int(Rnd * 2)
BPause = Int(Rnd * 2)
BStop = Int(Rnd * 2)
BNext = Int(Rnd * 2)
BEject = Int(Rnd * 2)

BVol = Int(Rnd * 2)
BPan = Int(Rnd * 2)
BPos = Int(Rnd * 2)

BTitle = Int(Rnd * 2)
BTit = Int(Rnd * 2)
BMin = Int(Rnd * 2)
BUp = Int(Rnd * 2)
BExit = Int(Rnd * 2)

BSideBar = Int(Rnd * 6)

SkipLights:

PaintWinamp 'Draw winamp on to the window

End Sub

Function ColDec(X As Single, Y As Single) As Long
Dim C As Long
' Looks at the current pixel in picCol
' The color is a shade of red:
' The play button is RGB(50, 0, 0)
' This is then divided by 10 to get button number 5

C = GetPixel(picCol.hdc, X, Y)

If C > 200 Then C = 0 'Not a shade of red therefore ignore it.
ColDec = Int(C / 10)

End Function

Sub ButtonDown(Col As Long)
Select Case Col

Case 1
BBack = 1

Case 2
BPlay = 1

Case 3
BPause = 1

Case 4
BStop = 1

Case 5
BNext = 1

Case 6
BEject = 1

Case 7
BShuf = 1

Case 8
BRep = 1

Case 9
BPos = 1

Case 10
BVol = 1

Case 11
BPan = 1

Case 12
BEq = 1

Case 13
BPl = 1

Case 14
BExit = 1

Case 15
BUp = 1

Case 16
BMin = 1

Case 17
BTit = 1

End Select
PaintWinamp
End Sub

Sub ButtonOver(Col As Long)
If Col = 13 And Not CurrentSong = "Fake Winamp by Michael Pote ** " Then CurrentSong = "Clicks the buttons randomly * " Else CurrentSong = "Fake Winamp by Michael Pote ** ": CutS = CurrentSong
End Sub

Sub ButtonUp(Col As Long)
BBack = 0
BPlay = 0
BPause = 0
BStop = 0
BNext = 0
BRep = 0
BShuf = 0
BPos = 0
BEq = 0
BPl = 0
BExit = 0
BUp = 0
BMin = 0
BTit = 0
BEject = 0
BVol = 0
BPan = 0
PaintWinamp

If Col = 14 Then Unload Me 'Exit button pressed
If Col = 4 Then Timer1.Enabled = False 'stop animation when stop is pressed
If Col = 2 Then Timer1.Enabled = True 'play animation when play is pressed

If Col = 5 Then 'Jump forward
For I = 0 To 10
Timer1_Timer
Next
End If

If Col = 1 Then 'Jump Back
Pos = Pos - 0.1
Volume = 0
Pan = 0
End If


If Col = 7 Then Shuffle = (Shuffle = False) 'Toggle lights
If Col = 8 Then Repeat = (Repeat = False) 'Toggle repeat button
If Col = 12 Then Eq = (Eq = False)
If Col = 13 Then Pl = (Pl = False)

End Sub

