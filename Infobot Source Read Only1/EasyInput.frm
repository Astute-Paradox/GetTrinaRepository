VERSION 5.00
Begin VB.Form EasyInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Pad - Create messages using the mouse."
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "EasyInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text13 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "EasyInput.frx":0442
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   6495
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "EasyInput.frx":047F
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   6495
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "EasyInput.frx":0557
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   2055
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "EasyInput.frx":0620
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   2055
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "EasyInput.frx":065E
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   1695
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "EasyInput.frx":0693
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   1695
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "EasyInput.frx":06D2
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   2535
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "EasyInput.frx":0714
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Oops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6960
      TabIndex        =   6
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Talk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6000
      TabIndex        =   5
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   5775
   End
   Begin VB.TextBox Text4 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "EasyInput.frx":0764
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   2535
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "EasyInput.frx":077F
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   6495
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "EasyInput.frx":07C7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "EasyInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Private LastUsersMessageTime
Private LastBotsMessageTime

Private InMessage
Private OutMessage

Public Sub HidePad()

EasyInput.Hide
OutMessage = "Easy Pad has been switched off."
Exit Sub

End Sub

Private Sub Highlight(Window)


End Sub


Private Function SelectedWord(Position, Text)

SelPosition = Position
If SelPosition = 0 Then SelPosition = 1: StartPosition = 1: GoTo GetEnd
StartPosition = InStrRev(Text, Chr(13) + Chr(10), SelPosition, vbBinaryCompare)
If StartPosition = 0 Then StartPosition = 1 Else StartPosition = StartPosition + 2
GetEnd:
EndPosition = InStr(SelPosition, Text, Chr(13) + Chr(10), vbBinaryCompare)
If EndPosition = 0 Then Exit Function

SelectedWord = Mid(Text, StartPosition, EndPosition - StartPosition)
SelectedWord = LCase(SelectedWord)

End Function


Public Sub ShowPad()

ShowIt:
If EasyInput.Visible Then OutMessage = "Easy Pad is already enabled.": Exit Sub
EasyInput.Show
OutMessage = "Easy Pad is enabled."

If Len(Things) > 32000 Then Text2.Text = Left(Things, 32000): Exit Sub
Text2.Text = Things


End Sub

Public Sub Update()
InMessage = ""
OutMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime
If " " + InMessage Like "*[!a-z]*Easypad[!a-z]*" Then GoTo CheckMore
If " " + InMessage Like "*[!a-z]*Easy pad[!a-z]*" Then GoTo CheckMore
Exit Sub

CheckMore:
If " " + InMessage Like "*[!a-z]off[!a-z]*" Then Call HidePad: GoTo Done
If InMessage Like "*disable*" Then Call HidePad: GoTo Done
If InMessage Like "*Hide*" Then Call HidePad: GoTo Done
If InMessage Like "*close*" Then Call HidePad: GoTo Done
Call ShowPad

Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage)
Exit Sub





End Sub


Private Sub Command1_Click()
Static LastMessage
Message = Text5.Text
If Message = "" Then Message = LastMessage
If Message = "" Then Call SetBotsMessage("Click on some words first."): Exit Sub
Call TidyMessage(Message)
Call SetUsersMessage(Message)
LastMessage = Message
Text5.Text = ""
End Sub

Private Sub Command2_Click()
If Len(Text5.Text) <= 1 Then Text5.Text = "": Exit Sub
p = InStrRev(Text5.Text, " ", Len(Text5.Text) - 1, vbBinaryCompare)
If p = 0 Then Text5.Text = "": Exit Sub
Message = Mid(Text5.Text, 1, p)
Text5.Text = Message

End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text1.SelStart, Text1.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text10.SelStart, Text10.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text11.SelStart, Text11.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text12.SelStart, Text12.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text13.SelStart, Text13.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text2.SelStart, Text2.Text)

Text5.Text = Text5.Text + Word + " "


End Sub


Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text3.SelStart, Text3.Text)

Text5.Text = Text5.Text + LCase(Word) + " "

End Sub


Private Sub Text4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text4.SelStart, Text4.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GoTo SendMessage
Exit Sub

SendMessage:
Message = Text5.Text
Call TidyMessage(Message)
Call SetUsersMessage(Message)
LastMessage = Message
Text5.Text = ""

End Sub


Private Sub Text6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text6.SelStart, Text6.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text7.SelStart, Text7.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text8.SelStart, Text8.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


Private Sub Text9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Word = SelectedWord(Text9.SelStart, Text9.Text)

Text5.Text = Text5.Text + Word + " "

End Sub


