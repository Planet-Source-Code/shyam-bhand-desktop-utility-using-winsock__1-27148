VERSION 5.00
Begin VB.Form Erase1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   9000
      Left            =   840
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Erase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************
'Copyright © 2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'For more my code samples visit:
'http://www.i.com.ua/~aka
'*************************************
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

Dim pict As Picture

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
Dim h
Dim w

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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    End
End If
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long
   Dim xSrc As Long, ySrc As Long
   Dim dwRop As Long, hwndSrc As Long, hSrcDC As Long
   Dim Res As Long
   Dim PixelColor, PixelCount
   If App.PrevInstance = True Then
      Unload Me
      Exit Sub
   End If
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
      Picture = pict
      WindowState = vbMaximized
      Picture1.Width = Screen.Width \ Screen.TwipsPerPixelX
      Picture1.Height = Screen.Height \ Screen.TwipsPerPixelY
      w = Screen.Width \ 300 + 1
      h = Screen.Height \ 225 + 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
End
End Sub



Private Sub Timer1_Timer()
xx = 0
yy = 0
For e = 0 To 299
 PaintPicture Picture1.Image, w * xx, h * yy, w, h, w * xx, h * yy, w, h, vbSrcCopy
 xx = xx + 1: If xx = 20 Then xx = 0: yy = yy + 1
 DoEvents
Next e
xx = 0
yy = 0
For e = 0 To 299
 PaintPicture Form1.Picture, w * xx, h * yy, w, h, w * xx, h * yy, w, h, vbSrcCopy
 xx = xx + 1: If xx = 20 Then xx = 0: yy = yy + 1
 DoEvents
Next e
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
xxx = 0
yyy = 0
For e = 0 To 159
 PaintPicture Picture1.Image, w * xxx, h * yyy, w, h, w * xxx, h * yyy, w, h, vbSrcCopy
 xxx = xxx + 1: If xxx = 20 Then xxx = 0: yyy = yyy + 2
 DoEvents
Next e
xxx = 0
yyy = 1
For e = 0 To 139
 PaintPicture Picture1.Image, w * xxx, h * yyy, w, h, w * xxx, h * yyy, w, h, vbSrcCopy
 xxx = xxx + 1: If xxx = 20 Then xxx = 0: yyy = yyy + 2
 DoEvents
Next e
xxx = 0
yyy = 0
For e = 0 To 159
 PaintPicture Form1.Picture, w * xxx, h * yyy, w, h, w * xxx, h * yyy, w, h, vbSrcCopy
 xxx = xxx + 1: If xxx = 20 Then xxx = 0: yyy = yyy + 2
 DoEvents
Next e
xxx = 0
yyy = 1
For e = 0 To 139
 PaintPicture Form1.Picture, w * xxx, h * yyy, w, h, w * xxx, h * yyy, w, h, vbSrcCopy
 xxx = xxx + 1: If xxx = 20 Then xxx = 0: yyy = yyy + 2
 DoEvents
Next e
Timer1.Enabled = True
Timer2.Enabled = False

End Sub
