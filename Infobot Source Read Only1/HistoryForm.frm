VERSION 5.00
Begin VB.Form HistoryForm 
   Caption         =   "Message History:"
   ClientHeight    =   5280
   ClientLeft      =   5475
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "HistoryForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6630
   Begin VB.TextBox HistoryWindow 
      Height          =   5295
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "HistoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InMessage
Public OutMessage
Private Const MessageHistorySize = 5000
Public MessageHistory
Private LastMessageLength
Private LastMessageTime

Option Compare Text

Private LastBotsMessageTime
Private LastUsersMessageTime
Public Sub HideView()
HistoryForm.Hide
'OutMessage = "Ok."

End Sub

Public Sub ShowView()
HistoryForm.Width = Interface.Width
If Interface.Top > HistoryForm.Height Then HistoryForm.Top = Interface.Top - HistoryForm.Height: GoTo SetLeft
If Interface.Top + Interface.Height + HistoryForm.Height < Screen.Height Then HistoryForm.Top = Interface.Top + Interface.Height


SetLeft:
HistoryForm.Left = Interface.Left

HistoryForm.Show
'OutMessage = "Ok."

Interface.UsersWindow.SetFocus


End Sub

Public Sub Update()

InMessage = ""
OutMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime

Msg1 = ""
Msg2 = ""

If InMessage <> "" Then Msg1 = UsersName + ": " + InMessage
If BotsMessageTime <> LastBotsMessageTime Then Msg2 = "Apad: " + BotsMessage: LastBotsMessageTime = BotsMessageTime

If Msg1 <> "" Then If Msg2 <> "" Then If UsersMessageTime > BotsMessageTime Then Temp = Msg1: Msg1 = Msg2: Msg2 = Temp

If Msg1 <> "" Then Call AddMessageToHistory(Msg1)
If Msg2 <> "" Then Call AddMessageToHistory(Msg2)


If InMessage Like "Show History[?.]*" Then GoTo ShowIt
If InMessage Like "*recent messages*" Then GoTo CheckMore
If InMessage Like "*history window*" Then GoTo CheckMore
If InMessage Like "*message history*" Then GoTo CheckMore

Call UpdateHistoryWindow
GoTo Done





CheckMore:
If InMessage Like "*Do not*" Then GoTo HideIt
If InMessage Like "*Show*" Then GoTo ShowIt
If InMessage Like "*open*" Then GoTo ShowIt
If InMessage Like "*display*" Then GoTo ShowIt
If InMessage Like "*View*" Then GoTo ShowIt
If InMessage Like "*Bring up*" Then GoTo ShowIt
If InMessage Like "*see*" Then GoTo ShowIt
If InMessage Like "*Close*" Then GoTo HideIt
If InMessage Like "*Remove*" Then GoTo HideIt
If InMessage Like "*Hide*" Then GoTo HideIt
GoTo Done


ShowIt:
Call ShowView
GoTo Done


HideIt:
Call HideView



Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage): LastBotsMessageTime = BotsMessageTime: Call AddMessageToHistory("Apad: " + BotsMessage)


End Sub

Public Sub AddMessageToHistory(Message)
If Message = "" Then Exit Sub
If Len(Message) > MessageHistorySize Then Exit Sub
Loop1:
If (Len(MessageHistory) + Len(Message)) < MessageHistorySize Then GoTo AddIt
Loop2:
Position = InStr(1, MessageHistory, Chr(13) + Chr(10) + Chr(13) + Chr(10))
If Position = 0 Then MessageHistory = Right(MessageHistory, Len(MessageHistory) - Len(Message)): GoTo Loop1
MessageHistory = Right(MessageHistory, Len(MessageHistory) - (Position + 1))
GoTo Loop1

AddIt:
MessageHistory = MessageHistory + (Message + Chr(13) + Chr(10) + Chr(13) + Chr(10))
LastMessageLength = Len(Message)
LastMessageTime = Timer

End Sub


Public Sub UpdateHistoryWindow()
Static PreviousData
If LastMessageTime + (LastMessageLength * 0.05) < Timer Then GoTo Update
Exit Sub

Update:
If Right(MessageHistory, 20) = PreviousData Then Exit Sub
PreviousData = Right(MessageHistory, 20)
HistoryForm.HistoryWindow.Text = MessageHistory
HistoryForm.HistoryWindow.SelStart = Len(HistoryForm.HistoryWindow.Text)


End Sub

Private Sub Form_Activate()
'HistoryForm.HistoryWindow.Text = MessageHistory
'HistoryForm.Left = Interface.Left
'HistoryForm.Top = Interface.Top - 5700
HistoryForm.HistoryWindow.SelStart = Len(HistoryForm.HistoryWindow.Text)

End Sub

Private Sub Form_Load()
'HistoryForm.Left = Interface.Left
'HistoryForm.Top = Interface.Top - 5700

End Sub


Private Sub Form_Resize()

HistoryForm.HistoryWindow.Width = HistoryForm.ScaleWidth
'If Interface.ScaleHeight < 50 Then Exit Sub
HistoryForm.HistoryWindow.Height = HistoryForm.ScaleHeight


End Sub


