Attribute VB_Name = "Mirc"
Public Enabled
Option Compare Text


Private LastUsersMessageTime
Private LastBotsMessageTime

' This module will send the bots messages to the clipboard and read any messages within the clipboard that start with 'Mirc:'
' The purpose of the module is to let the chatbot communicate with Mirc.
' Place this line into the aliases window in Mirc:
'  /checktest /if (infobot: isin $cb) { echo -a $remove($cb,Infobot:) | clipboard }
'
' Also, place these lines into the "Remote" script window in Mirc.
'  on *:TEXT:*:#quake.uk:/clipboard Mirc: $1-
'  on 1:CONNECT:/timer1 0 3 /checktest
'
Public Sub Update()

If Mirc.Enabled Then GoSub GetMircMessage


If UsersMessage = "" Then Exit Sub
If UsersMessage Like "use mirc[!a-z]*" Then GoTo EnableIt
If UsersMessage Like "enable mirc[!a-z]*" Then GoTo EnableIt
If UsersMessage Like "mirc on{!a-z]*" Then GoTo EnableIt
If UsersMessage Like "mirc." Then GoTo EnableIt
If Mirc.Enabled Then If BotsMessageTime <> LastBotsMessageTime Then Clipboard.SetText ("Infobot: " + BotsMessage): LastBotsMessageTime = BotsMessageTime


Exit Sub

EnableIt:
BotsMessage = "Enabled Mirc."
Mirc.Enabled = True
Exit Sub





GetMircMessage:
On Error GoTo Oops
Message = Clipboard.GetText
If Message Like "Infobot:*" Then Return
If Message Like "Mirc: *" Then Call SetUsersMessage(Right(Message, Len(Message) - 6)): Clipboard.Clear
Return
Oops:
Return


End Sub


