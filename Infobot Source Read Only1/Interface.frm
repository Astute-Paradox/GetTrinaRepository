VERSION 5.00
Begin VB.Form Interface 
   AutoRedraw      =   -1  'True
   Caption         =   "Answerpad - The Intelligent Data Bank."
   ClientHeight    =   1110
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7950
   Icon            =   "Interface.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   Begin VB.CommandButton Command1 
      Caption         =   "Don't say that!"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox UsersWindow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
   Begin VB.PictureBox BotsWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   0
      ScaleHeight     =   12.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   393.75
      TabIndex        =   1
      Top             =   0
      Width           =   7935
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu LoadKnow 
         Caption         =   "Read Facts."
      End
      Begin VB.Menu SaveMem 
         Caption         =   "Save memory"
      End
      Begin VB.Menu Exit 
         Caption         =   "Save memory and exit."
         Index           =   0
      End
      Begin VB.Menu ExitNoSave 
         Caption         =   "Exit without saving memory."
      End
   End
   Begin VB.Menu QuickMenu 
      Caption         =   "Commands"
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Cmd 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu QuickGuide 
         Caption         =   "Quick Guide"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AutoReturnTime   ' This is for the Auto-return feature, this holds the time that it should wait before taking the user's message.

Public LastPressedKey   ' Contains the keycode of the last pressed key
Public KeyPressed       ' This will be set to true if a key is pressed.
Public UsersInput       ' This holds the last message the user entered. It is unprocessed, so it may contain extra return codes.
Public CurrentSentence  ' The bot's messages are displayed one sentence at a time. CurrentSentence contains the sentence currently shown.
Public WholeMessage     ' This contains the whole of the bot's current message.
Public AutoMessage      ' This message is automatically typed into the User's window. It can be used to type out the message the user has spoken.

Dim UserMsgHistory(10)  ' This holds the most recent messages the User entered into the User's window.
Public UserMsgNumber    ' The number of the last message the user entered.
Const UserMsgHistorySize = 10 ' Number of messages that can be stored


Private LastUsersMessageTime
Private UserMessage

Private Heat

Option Compare Text





' This will speak <message> out loud using either an MSAgent or a speech engine
Private Sub BotSays(Message)

If Message = "" Then Message = "?"
SpokenMessage = Message

If Message Like "*year*" Then GoTo SayIt
If Message Like "*in ####*" Then GoTo SayIt
If Message Like "*###*" Then SpokenMessage = NumbersToWords(Message)
Call ReplaceWords("Whalley", "Warley", SpokenMessage)
Call ReplaceWords("The Sun", "the sun", SpokenMessage)  ' Sun is pronounced as Sunday unless it is lowercase.


SayIt:
If MSAgent.AgentActive Then MSAgent.MessageToSay = SpokenMessage: Exit Sub


SpeechHandler.Speak SpokenMessage
End Sub





Public Sub RecordUserMessage()
UserMsgHistory(UserMsgNumber) = UserMessage
UserMsgNumber = UserMsgNumber + 1
If UserMsgNumber = UserMsgHistorySize Then UserMsgNumber = 0
End Sub

' This will type the users previous message into the users window.

Public Sub RecallPreviousMessage()
N = 0
Loop1:
UserMsgNumber = UserMsgNumber - 1: If UserMsgNumber < 0 Then UserMsgNumber = UserMsgHistorySize - 1
If UserMsgHistory(UserMsgNumber) <> "" Then UsersWindow = UserMsgHistory(UserMsgNumber): Exit Sub
N = N + 1
If N < UserMsgHistorySize Then GoTo Loop1
End Sub


Public Sub RecallNextMessage()
N = 0
Loop1:
UserMsgNumber = UserMsgNumber + 1: If UserMsgNumber > UserMsgHistorySize - 1 Then UserMsgNumber = 0
If UserMsgHistory(UserMsgNumber) <> "" Then UsersWindow = UserMsgHistory(UserMsgNumber): Exit Sub
N = N + 1
If N < UserMsgHistorySize Then GoTo Loop1

End Sub






' This will set the Auto-return to the value specified by the user.
Private Sub SetAutoReturn()
If UserMessage Like "*auto*return*#*" Then GoTo GetValue
If UserMessage Like "*auto*enter*#*" Then GoTo GetValue
If UserMessage Like "*enable auto*return*" Then Value = "4": GoTo SetIt
If UserMessage Like "*enable auto*enter*" Then Value = "4": GoTo SetIt
If UserMessage Like "*auto*return*on*" Then Value = "4": GoTo SetIt
If UserMessage Like "*auto*enter*on*" Then Value = "4": GoTo SetIt
If UserMessage Like "*disable*auto*return*" Then Value = "80": GoTo SetIt
If UserMessage Like "*disable*auto*enter*" Then Value = "80": GoTo SetIt
If UserMessage Like "*auto*return*off*" Then Value = "80": GoTo SetIt
If UserMessage Like "*auto*enter*off*" Then Value = "80": GoTo SetIt
Exit Sub

GetValue:
Message = UserMessage
Value = SuperSum(Message)

SetIt:
Interface.AutoReturnTime = Val(Value)
BotsMessage = "Auto-return has been set..": BotsMessageTime = Timer

End Sub



Private Sub SetControlsPositions()

'GoSub SizeReducer
Command1.Visible = True


Interface.BotsWindow.Width = Interface.ScaleWidth
Interface.UsersWindow.Width = Interface.ScaleWidth

If Interface.ScaleHeight < 40 Then BotsWindow.Visible = False


If BotsWindow.Visible = False Then
 Interface.UsersWindow.Top = 0
Else
 Interface.UsersWindow.Top = BotsWindow.Top + BotsWindow.ScaleHeight + 10
End If



Command1.Left = (Interface.ScaleWidth / 2) - (Command1.Width / 2)

Frame1.Top = UsersWindow.Top + UsersWindow.Height
Frame2.Top = UsersWindow.Top + UsersWindow.Height
Frame1.Left = Interface.ScaleLeft + 10
Frame2.Left = Command1.Left + Command1.Width + 10
If Interface.ScaleWidth > 180 Then
Frame1.Width = Command1.Left - 20
Frame2.Width = Interface.ScaleWidth - Frame2.Left - 10
End If

If BotsWindow.Visible And Interface.ScaleHeight < 70 Then Command1.Visible = False: Exit Sub
If BotsWindow.Visible = False And Interface.ScaleHeight < 48 Then Command1.Visible = False: Exit Sub

Command1.Top = Interface.ScaleHeight - Command1.Height - 2

Interface.UsersWindow.Height = (Interface.ScaleHeight - UsersWindow.Top) - Command1.Height - 3




Exit Sub

SizeReducer:
If BotsWindow.Visible = False Then GoTo ReduceSize
If UsersWindow.Top > 0 Then Return
UsersWindow.Top = (BotsWindow.Top + BotsWindow.Height) + 1
Return
ReduceSize:
If UsersWindow.Top = 0 Then Return
UsersWindow.Top = 0
Return


End Sub

Public Sub Update()

If UsersMessageTime <> LastUsersMessageTime Then LastUsersMessageTime = UsersMessageTime: GoTo SetNewMessage

Refresh:
Call UpdateInterfaceWindows
Call SetAutoReturn


If UserMessage <> "" Then Call TidyMessage(UserMessage):  Call SetUsersMessage(UserMessage): LastUsersMessageTime = UsersMessageTime: UserMessage = ""




Exit Sub



SetNewMessage:
AutoMessage = UsersMessage
UsersWindow = UsersMessage
UsersWindow.SelStart = Len(UsersWindow)
GoTo Refresh


End Sub




Public Sub UpdateInterfaceWindows()

Static LastBotsMessageTime
Static TypePosition
Static CurrentMessage
Static MessagePart
Static CurrentMessagePosition
Static TypePause
Static BobFinishedTyping
Static WholeMessageTyped
Static BotsText
Static BotsTextPosition
Static LastMessagePartTime

GoSub ActOnBotsMessage

GoSub DoNextMessagePart

GoSub TypeMessagePart

GoSub DoAutoReturn

GoSub FadeBotsMessage

GoSub SetInterfaceSize

GoSub ClearUsersAutoMessages

Exit Sub




SetInterfaceSize:
' Resize the interface if the bots window isn't active.
If Interface.WindowState = 1 Then Return

Call SetControlsPositions
Return




FadeBotsMessage:
If BobFinishedTyping = False Then Return
If BotsWindow.ForeColor >= &HFFFFFF Then Return
BotsWindow.ForeColor = BotsWindow.ForeColor + &H30303
BotsWindow.CurrentX = BotsTextPosition
BotsWindow.CurrentY = 0
BotsWindow.Print BotsText
Return

'----------------------------------------------------------------------------------

ClearUsersAutoMessages:
Static ST
Static UMessageToType
If AutoMessage = "" Then Return
If Interface.UsersWindow = "" Then Return
If Timer > (LastUsersMessageTime + 2) Then UsersWindow = "": AutoMessage = "": Return
Return





DoAutoReturn:
Static PreviousSelStart
Static TimeLeft
If UsersWindow = "" Then Return
If Interface.AutoReturnTime > 60 Then Return
If UsersWindow.SelStart <> PreviousSelStart Then TimeLeft = Timer + AutoReturnTime: PreviousSelStart = UsersWindow.SelStart
If Timer > TimeLeft Then GoSub SendUserMessage: TimeLeft = Timer + AutoReturnTime
Return



SendUserMessage:
UsersInput = UsersWindow
UserMessageLength = Len(UsersInput)
UserMessage = Trim(UsersWindow)
UsersWindow = ""
Call RecordUserMessage
Call TidyMessage(UserMessage)
NewMessageAvailable = True
Return

'--------------------------------------------------------'

ActOnBotsMessage:
If Len(BotsMessage) = 0 Then Return
If LastBotsMessageTime = BotsMessageTime Then Return
LastBotsMessageTime = BotsMessageTime
BotsWindow.ForeColor = &H0
CurrentMessage = BotsMessage
Interface.BotsWindow.Cls
MessagePart = ""
BobFinishedTyping = True
CurrentMessagePosition = 0
TypePause = 0
TypePosition = 0
WholeMessageTyped = False
WholeMessage = ""
Return

'--------------------------------------------------------'

DoNextMessagePart:
If BobFinishedTyping = False Then Return
GoSub GetNextMessagePart
MessagePart = Trim(MessagePart)
If MessagePart <> "" Then BobFinishedTyping = False: BotsMessagePartTime = (Date + Time): BotsMessagePart = MessagePart: BotsWindow.Cls: CurrentSentence = MessagePart: WholeMessageTyped = False: Return
WholeMessageTyped = True
If CurrentMessage <> "" Then WholeMessage = CurrentMessage: CurrentMessage = ""
Return




TypeMessagePart:
If (Timer - LastMessagePartTime) < TypePause Then Return
If Len(MessagePart) = 0 Then TypePause = 0: TypePosition = 0: BobFinishedTyping = True: Return
If TypePosition >= Len(MessagePart) Then TypePause = (Len(MessagePart) / 40) + 1: MessagePart = "": Return
BobFinishedTyping = False
TypePosition = TypePosition + 1
BotsWindow.CurrentX = 0
BotsWindow.CurrentY = 0
BotsText = Mid(MessagePart, 1, TypePosition + 1)
If BotsWindow.TextWidth(BotsText) > (BotsWindow.Width - 5) Then BotsWindow.Cls: BotsWindow.CurrentX = ((BotsWindow.Width - 5) - BotsWindow.TextWidth(BotsText))
BotsTextPosition = BotsWindow.CurrentX
BotsWindow.Print BotsText
TimeOfLastBotActivity = (Date + Time)
LastMessagePartTime = Timer
Return



GetNextMessagePart:
MessagePart = ""
Loop1:
If CurrentMessagePosition >= Len(CurrentMessage) Then Return
Char = Mid(CurrentMessage, CurrentMessagePosition + 1, 1)
CurrentMessagePosition = CurrentMessagePosition + 1
If Asc(Char) < 32 Then Char = ""
MessagePart = MessagePart + Char
If CurrentMessagePosition > 1 Then If Mid(CurrentMessage, CurrentMessagePosition - 1, 3) Like "#.#" Then GoTo Loop1
If Mid(CurrentMessage, CurrentMessagePosition, 2) Like ".[!.'" + Chr(34) + "]" Then Return
If Mid(CurrentMessage, CurrentMessagePosition, 2) Like ":" + Chr(13) Then Return
'added this line April 2001
If Mid(CurrentMessage, CurrentMessagePosition, 2) Like ": " Then Return
GoTo Loop1


End Sub





Private Sub About_Click()

About1.Show
End Sub



Private Sub Cmd_Click(Index As Integer)

If Cmd(Index).Caption <> "" Then Call SetUsersMessage(Cmd(Index).Caption + ".")

End Sub

Private Sub Command1_Click()

Call DontSayMessage



End Sub

Private Sub Exit_Click(Index As Integer)

SaveMemoryAllowed = True
Call SetUsersMessage("Shutdown.")
End Sub

Private Sub ExitNoSave_Click()

Msg = "Do you really want to exit without saving my memory?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Memory not saved.."   ' Define title.
Ctxt = 1000   ' Define topic
 Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then SaveMemoryAllowed = False
If Response = vbNo Then Exit Sub


Call SetUsersMessage("Shutdown.")


End Sub


Private Sub ExtractFacts_Click()
Call SetUsersMessage("Extract facts.")
End Sub

Private Sub File_Click(Index As Integer)
TimeOfLastUserActivity = (Date + Time)

End Sub

Private Sub Form_Load()

UsersWindow = ""
BotsWindow.ForeColor = &H0

AutoReturnTime = 80




End Sub


Private Sub Form_Resize()
TimeOfLastUserActivity = (Date + Time)

If Interface.WindowState = vbMinimized Then QuietMode = True Else QuietMode = False


Call SetControlsPositions





End Sub


Private Sub Form_Unload(Cancel As Integer)
Static ShuttingDown

If ShuttingDown Then Exit Sub

Interface.Enabled = False

Cancel = -1     ' This prevents the form from closing.

SaveMemoryAllowed = True

UserMessage = "Shutdown."  ' Tell bot to shutdown.

ShuttingDown = True

End Sub


Private Sub Timer1_Timer()

Call UpdateInterfaceWindows

End Sub




Private Sub LoadKnow_Click()

Call SetUsersMessage("Load knowledge file.")



End Sub


Private Sub QuickGuide_Click()
Help1.Show
End Sub



Private Sub QuickMenu_Click()
TimeOfLastUserActivity = (Date + Time)

End Sub

Private Sub SaveMem_Click()

Call SetUsersMessage("Save memory.")

End Sub


Private Sub UsersWindow_KeyDown(KeyCode As Integer, Shift As Integer)

LastPressedKey = KeyCode
KeyPressed = True

TimeOfLastUserActivity = (Date + Time)


If AutoMessage <> "" Then UsersWindow.Text = "": AutoMessage = "": UsersWindow.ForeColor = &H0


End Sub


Private Sub UsersWindow_KeyUp(KeyCode As Integer, Shift As Integer)

TimeOfLastUserActivity = Date + Time

KeyPressed = False
If KeyCode = vbKeyReturn Then
UsersInput = UsersWindow

' remove any return codes within the users message.
Statement = ""
For L = 1 To Len(UsersInput)
If Asc(Mid(UsersInput, L, 1)) > 31 Then Statement = Statement + Mid(UsersInput, L, 1)
Next L

UsersInput = Statement
UserMessageLength = Len(UsersInput)
UserMessage = UsersInput
NewMessageAvailable = True
RecordUserMessage
UsersWindow.Text = ""
End If




If KeyCode = vbKeyUp Then
RecallPreviousMessage
UsersWindow.SelStart = Len(UsersWindow)
End If



If KeyCode = vbKeyDown Then
RecallNextMessage
UsersWindow.SelStart = Len(UsersWindow)
End If

End Sub






