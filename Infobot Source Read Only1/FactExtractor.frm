VERSION 5.00
Begin VB.Form FactExtractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract facts from text files."
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "FactExtractor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "This window shows text files only."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   3015
   End
End
Attribute VB_Name = "FactExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private TextData

Private InMessage
Private OutMessage

Option Compare Text

Private LastBotsMessageTime
Private LastUsersMessageTime

Private Sub ExtractTheFacts()

LastNotedTime = TimeOfLastUserActivity
ExtractedFacts = ""
If TextData = "" Then Exit Sub

Postition = 0
Loop1:
DoEvents: If TimeOfLastUserActivity <> LastNotedTime Then GoTo Abort
If Len(TextData) > 1000 Then GoSub ShowPercentage
GoSub GetNextParagraph
If Paragraph = "" Then GoTo SaveFacts
Paragraph = Replace(Paragraph, Chr(13) + Chr(10), " ")
Call RemoveUnwantedSpaces(Paragraph)
If Len(Paragraph) < 10 Then GoTo Loop1
If Len(Paragraph) > 400 Then GoTo Loop1
If Paragraph Like "*:" Then GoTo Loop1
If Paragraph Like "* * *" Then GoTo CheckWords
GoTo Loop1

CheckWords:
If Paragraph Like "*chapter*" Then GoTo Loop1
If Len(Memory) < 2000 Then GoTo RemoveLeadingRubbish
Temp = Paragraph
ChkWords:
Word = ExtractWord(Temp)
If Word = "" Then GoTo ItsOk
If InStr(1, Memory, Word, vbBinaryCompare) Then GoTo RemoveLeadingRubbish
GoTo ChkWords


RemoveLeadingRubbish:
If Paragraph = "" Then GoTo Loop1
If Left(Paragraph, 1) Like "[a-z0-9$£()]" Then GoTo ItsOk
Paragraph = Right(Paragraph, Len(Paragraph) - 1)
GoTo RemoveLeadingRubbish

ItsOk:
ExtractedFacts = ExtractedFacts + Paragraph + Chr(13) + Chr(10) + Chr(13) + Chr(10)
GoTo Loop1



ShowPercentage:
Percentage = ((100 / Len(TextData)) * Position)
Call Progress.Update(Percentage)
Return



SaveFacts:
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "Extracted Facts.Fct" For Output As #1
Print #1, ExtractedFacts
ExtractedFacts = ""
Close #1
Progress.Hide
Exit Sub


GetNextParagraph:
Paragraph = ""
GPLoop:
If Position >= Len(TextData) Then Return
Position = Position + 1
C = Asc(Mid(TextData, Position, 1))
If C < 32 Then GoTo GPLoop

EndPos = InStr(Position, TextData, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If EndPos = 0 Then Return
Paragraph = Mid(TextData, Position, EndPos - Position)
Position = EndPos + 3
Return



Abort:
Msg = "Facts extraction has been aborted." ' Define message.
Style = vbOKOnly + vbInformation + vbDefaultButton2   ' Define buttons.
Title = "Just a suggestion."   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
Exit Sub



End Sub


Private Sub ShowFiles()

'Dir1.Path = App.Path
'File1.Path = Dir1.Path
'File1.Pattern = "*.fct"
Dir1.Refresh
File1.Refresh
FactExtractor.Visible = True


End Sub

Public Sub Update()

InMessage = ""
OutMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime Else Exit Sub

If InMessage Like "Extract *facts*" Then Call ShowFiles: OutMessage = "Select a file.": GoTo Done




Exit Sub

Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage)

End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.txt"
File1.Refresh

End Sub


Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub


Private Sub File1_DblClick()
On Error GoTo Oops

FullPath = Dir1.Path + "\" + File1.List(File1.ListIndex)

If Dir(FullPath) = "" Then GoTo Done
Open FullPath For Input As #1
Size = LOF(1)
TextData = Input(Size, #1)
Close #1

If Len(TextData) < 50 Then GoTo DontBother

FactExtractor.Hide
DoEvents

Call ExtractTheFacts

Done:
Msg = "The extracted facts have been stored in Answerpad's folder, the file is named 'Extracted Facts.fct'."
Msg = Msg + "I recommend that you go and check this file with a text editor so as to make sure that the facts"
Msg = Msg + " are all in a correct and acceptable state. When you're satisfied with the facts then you can tell"
Msg = Msg + " Answerpad to read them into its memory by using the 'Read Facts' menu command..." + Chr(13) + Chr(10)
Style = vbOKOnly + vbInformation + vbDefaultButton2   ' Define buttons.
Title = "Just a suggestion."   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)

Exit Sub


DontBother:
Msg = "There isn't anything worth bothering about in that file. Try another..." ' Define message.
Style = vbOKOnly + vbInformation + vbDefaultButton2   ' Define buttons.
Title = "Just a suggestion."   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
Exit Sub

Oops:
Exit Sub



End Sub

Private Sub Form_Load()
File1.Pattern = "*.txt"

End Sub


