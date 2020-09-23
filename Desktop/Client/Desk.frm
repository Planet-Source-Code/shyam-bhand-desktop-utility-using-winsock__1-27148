VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00404040&
   Caption         =   "wsServer"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   Icon            =   "Desk.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Functions"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4935
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&xit Me"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "S&tretch"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Scroll"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Invert"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Flip"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Erase"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Connect"
      Height          =   375
      Left            =   1680
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Remote Ip"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    wsClient.Connect Text1.Text, 5555
End Sub

Private Sub Command2_Click()
    wsClient.SendData "Erase"
End Sub

Private Sub Command3_Click()
    wsClient.SendData "Flip"
End Sub

Private Sub Command4_Click()
    wsClient.SendData "Invert"
End Sub

Private Sub Command5_Click()
    wsClient.SendData "Scroll"
End Sub

Private Sub Command6_Click()
    wsClient.SendData "Stretch"
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Long
i = Shell("c:\program files\microsoft Office\Office10\winword.exe" & " " & "c:\desktop\Important.doc", vbMaximizedFocus)
End Sub

Private Sub wsClient_Close()
    MsgBox " You are disconnected from Server", vbApplicationModal, "Error"
    wsClient.Close
    End
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
    Dim strDataRecived As Integer
    wsClient.GetData strDataRecived
    DoEvents
End Sub

