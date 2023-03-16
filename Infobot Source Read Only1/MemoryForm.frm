VERSION 5.00
Begin VB.Form MemoryForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Contents:"
   ClientHeight    =   5280
   ClientLeft      =   12675
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "MemoryForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6630
   Begin VB.TextBox MemoryWindow 
      Height          =   5055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "MemoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InMessage
Public OutMessage
Option Compare Text

Private LastUsersMessageTime
Private LastBotsMessageTime

Private Sub ForceUpdate()

MemoryForm.MemoryWindow.Text = Right(Memory, 10000)
MemoryForm.MemoryWindow.SelStart = Len(MemoryForm.MemoryWindow.Text)
MemoryForm.Refresh

End Sub

Public Sub Update()

InMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime
OutMessage = ""

If InMessage Like "SM666*" Then GoTo ShowItReally
If InMessage Like "*Show memory*" Then GoTo ShowIt
If InMessage Like "*Display memory*" Then GoTo ShowIt
If InMessage Like "*Show *your memory*" Then GoTo ShowIt
If InMessage Like "*see *your memory*" Then GoTo ShowIt
If InMessage Like "*display *your memory*" Then GoTo ShowIt
If InMessage Like "*close *memory*" Then GoTo HideIt
If InMessage Like "*Shut *memory*" Then GoTo HideIt
Call UpdateMemoryWindow
GoTo Done



HideIt:
MemoryForm.Hide
OutMessage = "Ok."
GoTo Done



ShowIt:
Static SMC
SMC = SMC + 1
If SMC = 1 Then OutMessage = "This feature has been disabled.": GoTo Done
If SMC = 2 Then OutMessage = "Get out of my head!!.": GoTo Done
If SMC = 3 Then OutMessage = "No.": SMC = 3: GoTo Done

GoTo Done


ShowItReally:
Call ForceUpdate



If Interface.Top + Interface.Height + MemoryForm.Height < Screen.Height Then MemoryForm.Top = Interface.Top + Interface.Height: GoTo SetLeft
If Interface.Top > MemoryForm.Height Then MemoryForm.Top = Interface.Top - MemoryForm.Height: GoTo SetLeft


SetLeft:
MemoryForm.Left = Interface.Left


MemoryForm.Show
Call ForceUpdate

Interface.UsersWindow.SetFocus
RN = Int(Rnd * 3)
Dim R(3)
R(0) = "This window contains the most recent data only."
R(1) = "These are the latest memory items only."
R(2) = "Take a look at my most recent memories."
OutMessage = R(RN)


Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage)


End Sub


Public Sub UpdateMemoryWindow()
Static LastMemory

If Right(Memory, 20) = LastMemory Then Exit Sub
LastMemory = Right(Memory, 20)

Call ForceUpdate



End Sub

Private Sub Form_Activate()
'MemoryForm.Left = Interface.Left
'MemoryForm.Top = Interface.Top - 5700

End Sub

Private Sub Form_Load()
'MemoryForm.Left = Interface.Left
'MemoryForm.Top = Interface.Top - 5700
Call UpdateMemoryWindow
End Sub


