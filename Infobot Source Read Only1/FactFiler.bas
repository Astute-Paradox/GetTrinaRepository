Attribute VB_Name = "XFactFiler"
Private InMessage
Private OutMessage


Option Compare Text

Private LastBotsMessageTime
Private LastUsersMessageTime

Public Sub Update()

InMessage = ""
OutMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime


If InMessage Like "* *" Then GoTo CheckMessage
If InMessage Like "*=*" Then GoTo CheckMessage
Exit Sub

CheckMessage:
Call AddQuestionMark(InMessage)
If InMessage Like "*[?]." Then Exit Sub
Message = " " + InMessage
If Message Like "*[!a-z]I[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]you[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]me[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]your[ s,.!]*" Then GoTo GetPersonal
If Message Like "*[!a-z]my[!a-z]*" Then GoTo GetPersonal
If Message Like "*[!a-z]mine[!a-z]*" Then GoTo GetPersonal



GetGeneral:
GeneralKnowledge = GeneralKnowledge + InMessage + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Exit Sub



GetPersonal:
PersonalKnowledge = PersonalKnowledge + InMessage + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Exit Sub






End Sub


