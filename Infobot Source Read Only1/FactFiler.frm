VERSION 5.00
Begin VB.Form FactFiler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a fact file."
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "FactFiler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   120
      Pattern         =   "*"
      TabIndex        =   0
      Top             =   3480
      Width           =   3375
   End
End
Attribute VB_Name = "FactFiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private InMessage
Private OutMessage

Option Compare Text

Private LastBotsMessageTime
Private LastUsersMessageTime
Private Sub ShowFiles()

'Dir1.Path = App.Path
'File1.Path = Dir1.Path
'File1.Pattern = "*.fct"
Dir1.Refresh
File1.Refresh
FactFiler.Visible = True


End Sub



Public Sub Update()

Static PreInMessage
Static LastUsersName
InMessage = ""
OutMessage = ""


If UsersName <> LastUsersName Then PersonalKnowledge = PersonKnowledge + "My name is " + UsersName + "." + Chr(13) + Chr(10) + Chr(13) + Chr(10): LastUsersName = UsersName


If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime Else Exit Sub

If InMessage Like "Load *knowledge*" Then Call ShowFiles: OutMessage = "Select one.": GoTo Done


If PreInMessage Like InMessage Then GoTo Done
PreInMessage = InMessage

If InMessage Like "* *" Then GoTo CheckMessage
If InMessage Like "*=*" Then GoTo CheckMessage
Exit Sub

CheckMessage:
Call AddQuestionMark(InMessage)
If InMessage Like "*[?]." Then Exit Sub
Call ChangeToSubject(InMessage)
Message = " " + InMessage
If Message Like "*[!a-z]I[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]you[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]me[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]your[ s,.!]*" Then GoTo GetPersonal
If Message Like "*[!a-z]my[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]mine[!a-z]*" Then GoTo GetPersonal



GetGeneral:
MyCmd = GetCommand(InMessage): If MyCmd <> "" Then Exit Sub
If InMessage Like "Shut*down*" Then Exit Sub
GeneralKnowledge = GeneralKnowledge + InMessage + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Exit Sub


GetPersonal:
PersonalKnowledge = PersonalKnowledge + InMessage + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Exit Sub

Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage)

End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.txt;*.fct"
File1.Refresh
End Sub


Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub


Private Sub File1_DblClick()
'FileName = File1.FileName
On Error GoTo Done

FullPath = Dir1.Path + "\" + File1.List(File1.ListIndex)

FactFiler.Visible = False
FactFiler.Refresh
DoEvents

If Dir(FullPath) = "" Then GoTo Done



Open FullPath For Input As #1
Size = LOF(1)
Facts = Input(Size, #1)
Close #1

BotReadingFile = File1.List(File1.ListIndex)
BotReadingFullpath = FullPath


BotReadingPosition = 0

ReadingAllowed = True

If Len(Facts) < 100 Then Exit Sub
If Len(Facts) < 1000 Then Call SetBotsMessage("Reading..."): Exit Sub
If Len(Facts) < 5000 Then Call SetBotsMessage("Reading file..."): Exit Sub
Call SetBotsMessage("I will read the file... You may interrupt me at anytime and I will continue to read the file later. If at anytime you don't want me to read from the file then just say 'Stop reading'.")

Done:





End Sub


Private Sub Form_Load()
File1.Pattern = "*.txt;*.fct"

End Sub


