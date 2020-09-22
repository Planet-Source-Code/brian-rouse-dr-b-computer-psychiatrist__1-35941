VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ""Dr. B.!"""
   ClientHeight    =   1770
   ClientLeft      =   4380
   ClientTop       =   3585
   ClientWidth     =   4110
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1221.686
   ScaleMode       =   0  'User
   ScaleWidth      =   3859.502
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Copyright (c) 2002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Original ""Dr. B.!"" "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Version 1.01.01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "The Computer Psychiatrist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  CloseWindow
End Sub

Private Sub Image1_Click()
  CloseWindow
End Sub

Private Sub Label1_Click()
  CloseWindow
End Sub

Private Sub Label2_Click()
  CloseWindow
End Sub

Private Sub Label3_Click()
  CloseWindow
End Sub

Private Sub Label4_Click()
  CloseWindow
End Sub

Private Sub Timer1_Timer()
  CloseWindow
End Sub
Public Sub CloseWindow()
  Unload frmAbout
  Load frmMain
  frmMain.Show
End Sub

