VERSION 5.00
Begin VB.Form Progress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress..."
   ClientHeight    =   330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BarPic 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefLng C
DefLng R
DefLng G

Public Sub Update(Percentage)
If Progress.Visible = False Then Progress.Visible = True
Progress.SetFocus

BarLength = Int(Percentage * (BarPic.Width / 100))

RedAmount = 255 - Int(BarLength * (255 / BarPic.Width))
GreenAmount = 255 - RedAmount
Progress.BarPic.ForeColor = RGB(RedAmount, GreenAmount, 0)

Progress.BarPic.Line (0, 0)-(BarLength, BarPic.Height), , BF

Progress.BarPic.Line (BarLength, 0)-(BarPic.Width, BarPic.Height), 0, BF
Progress.BarPic.Refresh
Progress.Refresh

End Sub


Private Sub Picture1_Click()

End Sub


