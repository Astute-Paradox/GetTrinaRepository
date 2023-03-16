Attribute VB_Name = "Quake"
Public Enabled
Public InMessage



Public Keys As New clsKeyboard

Private ScriptCleared
Private LastScrMessage
Private TalkEnabled
Private SpecMode
Private LastDeathTime
Private OldMessage
Private QWTask

Private LastUsersMessageTime
Private LastBotsMessageTime

Private ServerName
Private PlayersAmount
Private MapName
Private PlayersNames(32)
Private PlayersFrags(32)


Option Compare Text
Private Sub GetLatestQWMessages()
If Quake.Enabled Then GoTo GetLatestMessage
Exit Sub


GetLatestMessage:
On Error GoTo Done

Open "d:\quake\qw\qconsole.log" For Input As #1
Size = LOF(1)
quakelog = Input(Size, #1)
Close #1
If Len(quakelog) = 0 Then Exit Sub



LastLogItem = ""
position = Len(quakelog)
position = InStrRev(quakelog, Chr(13) + Chr(10), position)
If position = 0 Then Exit Sub
Loop1:
If (position - 2) <= 0 Then Exit Sub
If Asc(Mid(quakelog, position, 1)) = 13 Then position = position - 1: GoTo Loop1
If Asc(Mid(quakelog, position, 1)) = 10 Then position = position - 1: GoTo Loop1
ItemEndPosition = position

ItemStartPosition = InStrRev(quakelog, Chr(13) + Chr(10), position - 1) + 2
If ItemStartPosition = 0 Then ItemStartPosition = 1

Message = Mid(quakelog, ItemStartPosition, ItemEndPosition - ItemStartPosition + 1)

GoSub ConvertMessage

If Message = OldMessage Then Exit Sub
OldMessage = Message



If Message Like "?name waffler" Then TalkEnabled = True: Exit Sub
If Message Like "?name ?*" Then TalkEnabled = False: Exit Sub


If LastScrMessage <> "" Then If Message Like "*waffler*:*" + LastScrMessage + "*" Then Exit Sub
If Message Like "]*" Then Exit Sub
If Message Like "*.*.*.*:*" Then Exit Sub
If Message Like "loc file*" Then Exit Sub


GetDeathMessages:
If Timer < LastDeathTime + 10 Then GoTo TestIfPlayerMessage
LastDeathTime = Timer
If Message Like "*burst into*" Then GoTo SetIt
If Message Like "*discharges into*" Then GoTo SetIt
If Message Like "*fell to*" Then GoTo SetIt
If Message Like "*accepts*" Then GoTo SetIt
If Message Like "*tries to*" Then GoTo SetIt
If Message Like "* was *" Then GoTo SetIt
If Message Like "* rides *" Then GoTo SetIt
If Message Like "* chewed *" Then GoTo SetIt
If Message Like "* becomes *" Then GoTo SetIt
If Message Like "* eats *" Then GoTo SetIt
If Message Like "* ate *" Then GoTo SetIt
If Message Like "* entered *" Then GoTo SetIt
GoTo TestIfPlayerMessage
SetIt:
SpecMode = True
UsersName = "QW"
GoTo SetMessage





TestIfPlayerMessage:
If Message Like "*?: ?*" Then GoTo GetPlayerMessage
Exit Sub


GetPlayerMessage:
SpecMode = False
If Message Like "*SPEC]*" Then SpecMode = True
Call ReplaceWords("[SPEC]", "", Message)
position = InStr(1, Message, ":")
If position = 0 Then Exit Sub
Name = Left(Message, position - 1)
Message = Right(Message, Len(Message) - position - 1)
If Message Like "*: ?*" Then Message = "": UsersName = "Unknown": GoTo SaveLog

UsersName = Name

If Message Like "*gg*" Then Call ReplaceWords("gg", "good game", Message)
If Message Like "hf" Then Message = "have fun."
If Message Like "gl" Then Message = "good luck."
If Message Like "heeh" Then Message = "heh."
If Message Like "lool" Then Message = "lol."
If Message Like "loo" Then Message = "lol."
If Message Like "yea" Then Message = "yes."
Call ReplaceWords("ra", "red armor", Message)
Call ReplaceWords("ga", "green armor", Message)
Call ReplaceWords("rl", "rocket launcher", Message)
Call ReplaceWords("u", "you", Message)
Call ReplaceWords("skillz", "skills", Message)

RemoveAsts:
If Left(Message, 1) = "*" Then Message = Mid(Message, 2): GoTo RemoveAsts
Message = Trim(Message)


SetMessage:
Call TidyMessage(Message)
Call SetUsersMessage(Message)

SaveLog:
If Len(quakelog) > 550 Then quakelog = Right(quakelog, 500): Open "d:\quake\qw\qconsole.log" For Output As #1: Print #1, quakelog: Close #1
Done:
Exit Sub




ConvertMessage:
For N = 1 To Len(Message)
C = Asc(Mid(Message, N, 1))
If C > 128 Then Mid(Message, N, 1) = Chr(C - 128)
Next N
Return


End Sub

Private Sub GetServerInfo()
If InMessage Like "S[0-9]*" Then GoTo GetIt
Exit Sub


GetIt:

Dim Servers(15)
Servers(1) = "213.221.174.98:27541"
Servers(2) = "213.221.174.98:27542"
Servers(3) = "213.221.174.98:27543"
Servers(4) = "213.221.174.98:27544"
Servers(5) = "213.221.174.98:27545"
Servers(6) = "213.221.174.193:27522"
Servers(7) = "213.221.174.193:27523"
Servers(8) = "213.221.174.67:27500"
Servers(9) = "213.221.174.67:27501"





ServerNumber = Val(Mid(InMessage, 2, 2))
If ServerNumber = 0 Then Message = "A number from 1 to 9.": GoTo Done
If ServerNumber > 9 Then Message = "I only have ten servers so far.": GoTo Done


IP = Servers(ServerNumber)

Call LoadServerInfo(IP)
If ServerName = "" Then Message = "I can't get the server info.": GoTo Done


If Val(PlayersAmount) = 0 Then Message = "There is nobody on " + ServerName: GoTo Done
If Val(PlayersAmount) = 1 Then Message = PlayersNames(1) + " is alone on " + ServerName: GoTo Done
If Val(PlayersAmount) = 2 Then If PlayersFrags(1) = 0 Then Message = PlayersNames(1) + " and " + PlayersNames(2) + " are farting about on " + ServerName + ".": GoTo Done
If Val(PlayersAmount) = 2 Then If PlayersFrags(1) > 0 Then Message = PlayersNames(1) + " and " + PlayersNames(2) + " are on " + ServerName + ". " + PlayersNames(1) + " is leading by" + Str(Val(PlayersFrags(1)) - Val(PlayersFrags(2))) + " frags on map " + MapName + ".": GoTo Done
If Val(PlayersAmount) > 2 Then If PlayersFrags(1) > 0 Then Message = "There are " + PlayersAmount + " monkeys on " + ServerName + ". " + PlayersNames(1) + " is owning everybody on map " + MapName + ".": GoTo Done
If Val(PlayersAmount) > 2 Then If PlayersFrags(1) = 0 Then Message = "There are players preparing for battle on " + ServerName + ".": GoTo Done



Done:
Call SetBotsMessage(Message)
Exit Sub


End Sub


Private Sub LoadServerInfo(IP)
Dim Item(32)
ServerName = ""

Server = IP

Current = CurDir
ChDir "D:\qstat"
t = Shell("qstat.exe -sort F -raw , -P -qws " + Server + " -of Output.txt", vbMinimizedNoFocus)
ChDir Current

'qstat -raw Inf- -P -qws 212.48.128.177:27500 -of please.txt

WaitTime = Timer + 3
QSWait: DoEvents: If Timer < WaitTime Then GoTo QSWait


Open "d:\qstat\Output.txt" For Input As #1
Size = LOF(1)
ServerInfo = Input(Size, #1)
Close #1
If Len(ServerInfo) = 0 Then GoTo Done

If ServerInfo Like "*,TIMEOUT*" Then GoTo Done

GoSub GetCertainInfo


Done:
Exit Sub





GetCertainInfo:
Line = 1
ServerName = ""
GoSub GetItems
If Item(1) = "" Then Return
ServerName = Item(3)
MapName = Item(4)
PlayersAmount = Item(6)
Line = 2
GCLoop:
GoSub GetItems
If Item(1) = "" Then Return
PlayersNames(Line - 1) = Item(2)
PlayersFrags(Line - 1) = Item(3)
Line = Line + 1
GoTo GCLoop




GetItems:
Item(1) = ""
GoSub GetLineText
If LineText = "" Then Return
ScanItems:
P = 1
Item(1) = ""
INum = 1
SCLoop:
P2 = InStr(P, LineText, ",")
If P2 = 0 Then Return
Item(INum) = Mid(LineText, P, P2 - P)
P = P2 + 1: INum = INum + 1
GoTo SCLoop
Return



GetLineText:
LineText = ""
P = 1
C = 0
GLoop:
P2 = InStr(P, ServerInfo, Chr(13) + Chr(10))
If P2 = 0 Then Return
C = C + 1: If C = Line Then GoTo GotLine
P = P2 + 1
GoTo GLoop
GotLine:
LineText = Mid(ServerInfo, P, (P2 - P) - 1)
Return





GetPlayerInfo:
PlayerInfo = ""
P = InStr(P + 1, ServerInfo, "#")
If P = 0 Then Return
EndPos = InStr(P + 1, ServerInfo, Chr(13) + Chr(10))
If EndPos = 0 Then Return
PlayerInfo = Mid(ServerInfo, P + 4, (EndPos - P) - 1)
Return



GetMapInfoEtc:
Mapinfo = ""
MIS = InStr(1, ServerInfo, Chr(13) + Chr(10))
If MIS = 0 Then Exit Sub
MIS = MIS + 2
MIE = InStr(MIS, ServerInfo, Chr(13) + Chr(10))
If MIE = 0 Then Exit Sub
Mapinfo = Mid(ServerInfo, MIS, (MIE - MIS))
Return



End Sub


Private Sub MakeQWScript()
If Quake.Enabled Then If BotsMessageTime <> LastBotsMessageTime Then LastBotsMessageTime = BotsMessageTime: GoTo MakeScript
Exit Sub

MakeScript:
Message = BotsMessage
On Error GoTo MSError

If TalkEnabled = False Then Exit Sub

If Right(Message, 1) = "." Then Message = Left(Message, Len(Message) - 1)
If Message = "" Then Exit Sub

LastScrMessage = Message
Script = ""
If SpecMode = True Then Script = "say_team " + Chr(34) + Message + Chr(13) + Chr(10): GoTo SaveIt
Script = "say " + Chr(34) + Message + Chr(13) + Chr(10)

SaveIt:
Open "d:\quake\id1\chat.scr" For Output As #1
Print #1, Script
Close #1
ScriptCleared = False

AppActivate QWTask

Keys.PressKeyVK keyF9

MSError:
Exit Sub



ClearScript:
If ScriptCleared Then Return
Script = "wait"
Open "d:\quake\id1\chat.scr" For Output As #1
Print #1, Script
Close #1
ScriptCleared = True
Return





End Sub

Private Sub UseQuake()
If InMessage Like "Use quake." Then GoTo StartIt
If InMessage Like "Enable quake." Then GoTo StartIt
Exit Sub

StartIt:
If Quake.Enabled = True Then Exit Sub


Current = CurDir
ChDir "D:\quake"
t = Shell("glmqwcl.exe -condebug -zone 4096 -bpp 32 -width 1024 -height 768 -no8bit -heapsize 64000 -nojoy -noipx +m_filter 1 -freq 100 +cl_maxfps 30 +yieldcpu 1", vbNormalFocus)
AppActivate t
QWTask = t
ChDir Current

Quake.Enabled = True

Done:
Call SetBotsMessage("Started Quake."): LastBotsMessageTime = BotsMessage
Exit Sub


End Sub

Public Sub Update()


Call GetLatestQWMessages

InMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime

Call GetServerInfo
Call UseQuake
Call MakeQWScript



End Sub

