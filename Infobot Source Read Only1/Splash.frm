VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   FillColor       =   &H0000C000&
   ForeColor       =   &H00008000&
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":030A
   ScaleHeight     =   3570
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Version 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label CopyRight 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © David Whalley 1999-2001. All Rights Reserved"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label Okay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing, please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label OkayShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub


Private Sub Label2_Click()
Unload Splash

End Sub


Private Sub Form_Click()
If Splash.Label1.Caption Like "*initialzing*" Then Exit Sub
Unload Splash

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Okay.Visible Then Unload Splash

End Sub

Private Sub Okay_Click()
Unload Splash
End Sub


Private Sub Okay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
OkayShadow.Top = 2440
OkayShadow.Left = 2320


End Sub


