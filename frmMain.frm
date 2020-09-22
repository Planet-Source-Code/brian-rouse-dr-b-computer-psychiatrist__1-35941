VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Talk to ""DoctorB."""
   ClientHeight    =   5895
   ClientLeft      =   1815
   ClientTop       =   1680
   ClientWidth     =   8625
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      DownPicture     =   "frmMain.frx":0442
      Height          =   1455
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Respond"
      DownPicture     =   "frmMain.frx":5A64
      Height          =   1455
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtQuestion 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   3360
      Width           =   6015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "frmMain.frx":B132
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Type here:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblReply 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   """Dr. B."" replies:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblConversation 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Your conversation with ""Doctor Brian"" :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer
Private Sub cmdExit_Click()
MsgBox "BRouse@AAANP.ORG", vbInformation, "Reply to:"
  Unload Me
End Sub




Private Sub Command1_Click()

    
  If txtQuestion.Text = "" Then
lblReply.Caption = "Cut the games and type something in the box, you are not that crazy!"
  Else
    HandleReply
  txtQuestion.SetFocus
  End If
  

End Sub

Private Sub Command2_Click()
MsgBox "BRouse@AAANP.ORG or WWW.AAANP.ORG    ", vbInformation, "Reply to:"
  Unload Me
End Sub

Private Sub Form_Load()
  lblReply.Caption = Greeting
    Label4.Caption = "Please Type your first name below:"
End Sub

Private Sub Image1_Click()
Unload Me
frmAbout.Show
End Sub

Private Sub Label1_Click()
txtQuestion = ""
End Sub



Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    HandleReply
  End If
End Sub

Public Sub HandleReply()
  Const LOWER = 1
  Const UPPER = 10
  Static TalkArray(LOWER To UPPER) As String
  Dim OldReply As String
  Dim TempString As String
  
  NL = Chr(10) & Chr(13)
  OldReply = lblReply.Caption
  If lblConversation.Caption <> "" Then
    For i = LOWER To UPPER - 2
      TalkArray(i) = TalkArray(i + 2)
    Next i
  End If
  TalkArray(9) = "Dr. B. : " & lblReply.Caption
  TalkArray(10) = Text1.Text & " : " & txtQuestion.Text
  TempString = ""
  For i = LOWER To UPPER
    TempString = TempString & TalkArray(i) & NL
  Next i
  lblConversation.Caption = TempString
  lblReply.Caption = NewReply(OldReply, txtQuestion.Text)
  txtQuestion.Text = ""
End Sub

