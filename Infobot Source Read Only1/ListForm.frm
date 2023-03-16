VERSION 5.00
Begin VB.Form ListForm 
   Caption         =   "Speech Engines:"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "ListForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListWindow 
      Height          =   5520
      ItemData        =   "ListForm.frx":0442
      Left            =   120
      List            =   "ListForm.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "ListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedItem
Private Sub Form_Load()
SelectedItem = -1
End Sub

Private Sub ListWindow_Click()
SelectedItem = ListWindow.ListIndex
End Sub


