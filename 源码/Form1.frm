VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "交通灯"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdstart 
      Caption         =   "开启红绿灯"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "破坏交通线路网"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2040
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      Height          =   1575
      Left            =   480
      Shape           =   3  'Circle
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g As Integer
Dim y As Integer
Dim r As Integer
Private Sub cmdstart_Click()
cmdstart.Visible = False
cmdstop.Visible = True
Timer1.Enabled = True
Label1.Visible = True
Shape1.Visible = True
End Sub
Private Sub cmdstop_Click()
Form1.Cls
Form1.Visible = False
Form2.Show
End Sub
Private Sub Form_Load()
Form1.BackColor = vbBlack
Label1.BackColor = vbBlack
Shape1.FillColor = vbGreen
Label1.ForeColor = vbGreen
Label1.Caption = g
Shape1.Visible = False
Label1.Visible = False
cmdstop.Visible = False
g = 10
y = 0
r = 0
Timer1.Enabled = False
End Sub
Private Sub Timer1_Timer()
Static n As Integer
   n = n + 1
   Shape1.BackStyle = 1
   Select Case n
      Case 1 To 10
         Shape1.BackColor = vbGreen
         Label1.Caption = g
         Label1.ForeColor = vbGreen
         g = g - 1
         y = 3
      Case 11 To 16
         If n Mod 2 = 1 Then
         Timer1.Interval = 500
            Shape1.BackColor = vbYellow
            Label1.Caption = y
            Label1.ForeColor = vbYellow
            y = y - 1
            r = 10
         Else
            Shape1.BackStyle = 0
         End If
      Case 17 To 26
      Timer1.Interval = 1000
         Shape1.BackColor = vbRed
         Label1.Caption = r
         Label1.ForeColor = vbRed
         r = r - 1
         g = 10
        If n = 26 Then n = 0
   End Select
End Sub
