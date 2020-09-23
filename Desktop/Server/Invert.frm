VERSION 5.00
Begin VB.Form Invert 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H0000FFFF&
   Icon            =   "Invert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Invert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************
'Copyright © 2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'http://www.i.com.ua/~aka
'*************************************

Dim pict As Picture
Dim j, p

Private Declare Function BitBlt _
    Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long

Private Declare Function GetDesktopWindow _
    Lib "user32" () As Long

Private Declare Function GetDC _
    Lib "user32" ( _
    ByVal hwnd As Long _
    ) As Long

Private Declare Function ReleaseDC _
    Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hdc As Long _
    ) As Long
    'XXXXXXXXXXX
    'XXXXXXXXXXX   OnTop = true Or false
    'XXXXXXXXXXX
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos _
    Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hwndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal CX As Long, _
    ByVal CY As Long, _
    ByVal wFlags As Long _
    ) As Long
Private mbOnTop As Boolean

Private Property Let OnTop(Setting As Boolean)
    If Setting Then
        SetWindowPos hwnd, HWND_TOPMOST, _
            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, _
            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    mbOnTop = Setting
End Property

Private Property Get OnTop() As Boolean
    OnTop = mbOnTop
End Property
'XXXXXXXXXXX
'XXXXXXXXXXX
'XXXXXXXXXXX


Private Sub Form_Activate()
    OnTop = True
    BeginPlaySound 5
ex:
For j = 0 To Picture1.ScaleWidth - 1
For p = 0 To Picture1.ScaleHeight - 1
Picture1.PSet (j, p), 16777215 - Picture1.Point(j, p)
Next p
BeginPlaySound 7
Next j
GoTo ex
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeyEscape Then
    EndPlaySound
    Unload Me
End If

End Sub

Private Sub Form_Load()
    Dim x     As Long, y As Long
    Dim xSrc  As Long, ySrc As Long
    Dim dwRop As Long, hwndSrc As Long, hSrcDC As Long
    Dim Res   As Long
    Dim m1, m2
    Dim n1, n2
    Dim PixelColor, PixelCount
    If App.PrevInstance = True Then
        Unload Me
        Exit Sub
    End If
    Width = Screen.Width
    Height = Screen.Height
    Randomize
    ScaleMode = vbPixels
    Move 0, 0, Screen.Width + 1, Screen.Height + 1
    dwRop = &HCC0020
    hwndSrc = GetDesktopWindow()
    hSrcDC = GetDC(hwndSrc)
    Res = BitBlt(hdc, 0, 0, ScaleWidth, _
        ScaleHeight, hSrcDC, 0, 0, dwRop)
    Res = ReleaseDC(hwndSrc, hSrcDC)
    Show
    Set pict = Image
    WindowState = vbMaximized
    Picture1.Width = Screen.Width \ 15
    Picture1.Height = Screen.Height \ 15
    Picture1 = pict
End Sub
