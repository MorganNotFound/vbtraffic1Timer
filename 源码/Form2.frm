VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer timer2 
      Left            =   2160
      Top             =   1920
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents timer2 As Timer
Attribute timer2.VB_VarHelpID = -1
Dim r&, r1&, t&, a1!, a2!, xb!, yb!, s!, b#
Private Sub Form_Load()
Me.Width = 4500: Me.Height = 4500
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Me.AutoRedraw = True
Me.Caption = "Azis¿ñÎè"
Set timer2 = Controls.Add("vb.timer", "Timer1")
timer2.Interval = 10
End Sub
Private Sub Timer2_Timer()
Randomize
r = 340 * Rnd
If r <> 0 Then
r1 = 500
s = r * Rnd
b = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
For t = 1 To 10000
a1 = t * 3.1415926 / 180
a2 = (r1 / r) * a1
xb = 500 + (-(r1 - r) * Cos(a1) - s * Cos(a2 - a1) + 420) * 4
yb = 500 + ((r1 - r) * Sin(a1) - s * Sin(a2 - a1) + 380) * 4
Me.PSet (xb, yb), b
Next t
End If
End Sub
