VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPECIAL FX BY SIMON PRICE"
   ClientHeight    =   5616
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6252
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   816
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   1092
   End
   Begin VB.PictureBox PB2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   1920
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   2760
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   6096
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Do It !"
      Height          =   372
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.ListBox List1 
      Height          =   816
      ItemData        =   "Form1.frx":030A
      Left            =   720
      List            =   "Form1.frx":031D
      TabIndex        =   1
      Top             =   720
      Width           =   1932
   End
   Begin VB.PictureBox Display 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   720
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   1800
      Width           =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "Choose an effect, choose a speed and then click the button to see it in action!"
      Height          =   372
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   4692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
  x As Byte
  y As Byte
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Const HWAVE = 0
Const VWAVE = 1
Const HFLIP = 2
Const VFLIP = 3
Const ZOOM = 4

Const PI = 3.1415
Const PIdiv180 = PI / 180

Dim PBWidth, PBHeight, HalfWidth, HalfHeight As Integer
Dim i, x, y, x2, y2  As Integer
Dim Color As Long
Dim lpPoint As POINTAPI
Dim yy, inc As Single

Private Sub cmdDoIt_Click()
PB.Cls
Display.Cls

Select Case List1.ListIndex
Case HFLIP
  DoFlip HFLIP
Case VFLIP
  DoFlip VFLIP
Case ZOOM
  DoZoom
Case HWAVE
  DoWave HWAVE
Case VWAVE
  DoWave VWAVE
End Select
End Sub

Sub DoFlip(WhichWay As Byte)

Select Case WhichWay

Case HFLIP
  For x = 0 To PBWidth Step List2.ListIndex + 1
    PB.Cls
    PB.Line (0, 0)-(PBWidth, PBHeight), vbWhite, BF
    StretchBlt PB.hdc, x, 0, PBWidth - 2 * x, PBHeight, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next
  For x = PBWidth To 0 Step -List2.ListIndex
    PB.Cls
    StretchBlt PB.hdc, x, 0, PBWidth - 2 * x, PBHeight, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next

Case VFLIP
  
  For y = 0 To PBHeight Step List2.ListIndex + 1
    PB.Cls
    PB.Line (0, 0)-(PBWidth, PBHeight), vbWhite, BF
    StretchBlt PB.hdc, 0, y, PBWidth, PBHeight - 2 * y, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next
  For y = PBHeight To 0 Step -List2.ListIndex
    PB.Cls
    StretchBlt PB.hdc, 0, y, PBWidth, PBHeight - 2 * y, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next

End Select

End Sub

Sub DoZoom()
inc = PBHeight / PBWidth * (List2.ListIndex + 1)
yy = 0
  For x = 0 To HalfWidth Step List2.ListIndex + 1
    yy = yy + inc
    StretchBlt PB.hdc, x, y, PBWidth - x * 2, PBHeight - 2 * yy, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next
yy = HalfHeight
  For x = HalfWidth To 0 Step -List2.ListIndex - 1
    yy = yy - inc
    StretchBlt PB.hdc, x, y, PBWidth - x * 2, PBHeight - 2 * yy, PB2.hdc, 0, 0, PBWidth, PBHeight, vbSrcCopy
    BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
  Next
End Sub

Sub DoWave(WhichWay As Byte)

Select Case WhichWay

Case HWAVE
For y = 0 To (HalfHeight \ 3) * (List2.ListIndex + 1) Step List2.ListIndex + 1
  For y2 = 0 To PBHeight
    BitBlt PB.hdc, Sin((y2 + y) * PIdiv180) * 30, y2, PBWidth, 1, PB2.hdc, 0, y2, vbSrcCopy
  Next
  BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
Next

Case VWAVE
For x = 0 To (HalfWidth \ 4) * (List2.ListIndex + 1) Step List2.ListIndex + 1
  For x2 = 0 To PBWidth
    BitBlt PB.hdc, x2, Sin((x2 + x) * PIdiv180) * 30, 1, PBHeight, PB2.hdc, x2, 0, vbSrcCopy
  Next
  BitBlt Display.hdc, 0, 0, PBWidth, PBHeight, PB.hdc, 0, 0, vbSrcCopy
Next

End Select

End Sub

Private Sub Form_Load()
'load picture, you can change this if you want
PB2 = LoadPicture(App.Path & "\SpecialFX.jpg")
Show
DoEvents
'copy pic into invisible pic
PB = PB2
'remember size of pic
PBWidth = PB.Width
PBHeight = PB.Height
HalfWidth = PBWidth \ 2
HalfHeight = PBHeight \ 2
'fill speed listbox
For i = 0 To 39
List2.AddItem i + 1, i
Next
'select defaults
List1.ListIndex = HWAVE
List2.ListIndex = 19
'press button
'cmdDoIt.Value = True
End Sub
