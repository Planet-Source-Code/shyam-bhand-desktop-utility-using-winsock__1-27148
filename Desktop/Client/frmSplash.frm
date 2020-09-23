VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H000000FF&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   5280
         Top             =   3480
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":000C
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2175
         Left            =   2040
         TabIndex        =   3
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Special thanks to Bhushan Joshi "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Designed by Shyam Bhand    Reetesh Singh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Image imgLogo 
         Height          =   1185
         Left            =   360
         Picture         =   "frmSplash.frx":0094
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    i = 0
End Sub



Private Sub lblCompany_Click()
End Sub



Private Sub Timer1_Timer()
    i = i + 1
        If i = 15 Then
            Unload Me
            MainForm.Show
            Unload Me
        End If
End Sub
