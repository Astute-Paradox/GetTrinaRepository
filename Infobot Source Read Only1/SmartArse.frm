VERSION 5.00
Begin VB.Form SmartArse 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "SmartArse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public InMessage
Public OutMessage
Public OutAction
Public PreviousInMessage
Public PreviousOutMessage
Public ChangedMessage
Public Positions
Public Actions
Public PositionsAmount
Public Because
Public So
Public SwearWords
Public CensoredWords

Private BoredomTolerance

Private NothingToSay

Public DateNames
Option Compare Text

Private BotsLastSeenMessageTime
Private UsersLastSeenMessageTime
Private LastInternalMessageTime


Private SayData As SayDataStructure
Const Says = 0
Const Mentions = 1
Const AsksAbout = 2

Private Type SayDataStructure
Method(100) As Byte
Person(100) As String
Says(100) As String
Response(100) As String
End Type

Public SayItemsAmount
Public SayDataLastItemNumber

 




Private Sub AbortReading()
If InMessage Like "* reading*" Then GoTo Maybe
Exit Sub


Maybe:
Message = InMessage
Call ReplaceWords("Stop", "Abort", Message)
Call ReplaceWords("Cancel", "Abort", Message)
Call ReplaceWords("Please", "", Message)
If Message Like "Abort reading*" Then GoTo Abort
Exit Sub


Abort:
ReadingAllowed = False
OutMessage = "Okay, I wont read anymore."
Exit Sub



End Sub

' Acknowledge the users information
Private Sub AcknowledgeStatement()
If UsersName = "QW" Then Exit Sub
If InMessage Like "*[?]*" Then Exit Sub
Acknowledgments = Array("Ok.", "I see.", "Right..", "Okay..", "Fine..", "Great.", "Thanks for the info.")
AcknowledgmentNumber = Int((Rnd * 6) + 1)
OutMessage = Acknowledgments(AcknowledgmentNumber)
End Sub



' Adds a list of numbers together
' Example:
' "Add together 10 20 30 4.32 and 9.16"
' or just:
' "10 3.2 5.5"
Private Sub AddList()
If InMessage Like "Add*" Then GoTo CheckFurther
If InMessage Like "#* #*" Then GoTo CheckFurther
Exit Sub

CheckFurther:
Message = WordsToNumbers(InMessage)
If Message Like "*#*#*" Then GoTo GetNumbers
Exit Sub


GetNumbers:
Word = ExtractWord(Message)
If Word = "" Then GoTo Addem
If Word Like "and" Then GoTo GetNumbers
If Word Like "add" Then GoTo GetNumbers
If Word Like "numbers" Then GoTo GetNumbers
If Word Like "together" Then GoTo GetNumbers
If Word Like "these" Then GoTo GetNumbers
If Word Like "all" Then GoTo GetNumbers
If Word Like "to" Then GoTo GetNumbers
If Val(Word) Then If Numbers = "" Then Numbers = Numbers + Word: GoTo GetNumbers Else Numbers = Numbers + "+" + Word: GoTo GetNumbers
Exit Sub




Addem:
OutMessage = SuperSum(Numbers)
Exit Sub


End Sub

Private Sub Ask()
If InMessage Like "*Ask*" Then GoTo CheckFurther
Exit Sub

CheckFurther:
If InMessage Like "If *" Then Exit Sub
If InMessage Like "* If *" Then Exit Sub
Words = Array("him", "her", "me", "them", "us", "a", "the", "")
N = 0
Loop1:
If Words(N) = "" Then GoTo CheckMore
Item = "Ask " + Words(N): GoSub CheckIt
Item = "Will you ask " + Words(N): GoSub CheckIt
Item = "Can you ask " + Words(N): GoSub CheckIt
Item = "Please ask " + Words(N): GoSub CheckIt
N = N + 1
GoTo Loop1:

CheckMore:
Item = "Ask": GoSub CheckIt
Item = "Will you ask": GoSub CheckIt
Item = "Can you ask": GoSub CheckIt
Item = "Please ask": GoSub CheckIt
Exit Sub

CheckIt:
If InMessage Like Item + " *" Then GoTo DoIt
Return

DoIt:
OutMessage = WordsBetween(Item, "", InMessage)

If OutMessage Like "'*?*'" Then OutMessage = Mid(OutMessage, 2, Len(OutMessage) - 2)
If OutMessage Like "'*?*'[?]." Then OutMessage = Mid(OutMessage, 2, Len(OutMessage) - 4)

If OutMessage Like "*[?]*" Then Exit Sub

OutMessage = Trim(OutMessage)

Loop2:
If Right(OutMessage, 1) = "." Then OutMessage = Left(OutMessage, Len(OutMessage) - 1): GoTo Loop2
OutMessage = OutMessage + "?."



End Sub





Private Sub CensorMessage(Message)
If SwearingAllowed Then Exit Sub

N = 0
Loop1:
If SwearWords(N) = "" Then Exit Sub
If " " + Message Like "*[!a-z]" + SwearWords(N) + "[!a-z]*" Then Call ReplaceWords(SwearWords(N), CensoredWords(N), Message): GoTo GetNext
If " " + Message Like "*[!a-z]" + SwearWords(N) + "s[!a-z]*" Then Call ReplaceWords(SwearWords(N) + "s", CensoredWords(N) + "s", Message): GoTo GetNext
If " " + Message Like "*[!a-z]" + SwearWords(N) + "ing[!a-z]*" Then Call ReplaceWords(SwearWords(N) + "ing", CensoredWords(N) + "ing", Message)
GetNext:
N = N + 1
GoTo Loop1



End Sub



Public Sub DontSayThat()
If " " + InMessage Like "*[!a-z]do not*[!a-z]say that*" Then GoTo MakeNote
If InMessage Like "*Stop saying that*" Then GoTo MakeNote
Exit Sub



MakeNote:
Call DontSayMessage

End Sub

Public Sub ForgetAboutX()



End Sub

Private Sub MultiAnswerer()
Static LastNotedTime
Static Answering
Static ResponseCount


If MultiAnswerMode = False Then Exit Sub
'If UsersMessage Like "*[!?]*" Then Exit Sub
If UsersMessageTime <> LastNotedTime Then LastNotedTime = UsersMessageTime: Answering = True: ResponseCount = 0

If Answering = False Then Exit Sub

TimeOfLastActivity = TimeOfLastUserActivity: LastMessageLength = Len(UsersMessage)
If DateDiff("s", TimeOfLastUserActivity, TimeOfLastBotActivity) > 0 Then TimeOfLastActivity = TimeOfLastBotActivity: LastMessageLength = Len(BotsMessage)
InactivityTime = DateDiff("s", TimeOfLastActivity, Date + Time)
If InitialPause = 3 Then GoTo CheckTime
If ResponseCount = 0 Then InitialPause = 3 Else InitialPause = 1
CheckTime:
If InactivityTime > (InitialPause + (LastMessageLength / 30)) Then GoTo GetSomethingToSay
Exit Sub



GetSomethingToSay:
Debug.Print " "
Debug.Print "---"
Debug.Print "MA!!!"
Temp = InMessage
InMessage = UsersMessage
Call CorrectGrammarEtc(InMessage)
Call AddQuestionMark(InMessage)
InferringAllowed = False
Call FindRelevantResponse
InMessage = Temp
If NothingToSay Then Answering = False: If ResponseCount > 0 Then OutMessage = "": Exit Sub Else OutMessage = "": Exit Sub
ResponseCount = ResponseCount + 1
Exit Sub


End Sub

Private Sub RambleHandler()

Static LastNotedTime
Static RambleMode
Static RamblePause
Static MessageInWaiting


If RamblingAllowed = False Then Exit Sub
If UsersMessageTime <> LastNotedTime Then RambleMode = -1: RamblePause = 3: LastNotedTime = UsersMessageTime: NothingToSay = False: MessageInWaiting = "": Exit Sub
If RambleMode = 0 Then Exit Sub
TimeOfLastActivity = TimeOfLastUserActivity: LastMessageLength = Len(UsersMessage)
If DateDiff("s", TimeOfLastUserActivity, TimeOfLastBotActivity) > 0 Then TimeOfLastActivity = TimeOfLastBotActivity: LastMessageLength = Len(BotsMessage)
InactivityTime = DateDiff("s", TimeOfLastActivity, Date + Time)
If MessageInWaiting <> "" Then GoTo ProcessWaitingMessage
If RambleMode = -1 Then InitialPause = 10 Else InitialPause = 1
If InactivityTime > (InitialPause + (LastMessageLength / 60)) Then GoTo GetSomethingToSay
Exit Sub



GetSomethingToSay:
If NothingToSay = True Then If RambleMode = 0 Then Exit Sub
RambleMode = 1
MsgTemp = InMessage
SubTemp = Subject
InMessage = BotsMessage
Call ChangeYourForMyEtc(InMessage)
Call GetSubject(InMessage)
InferringAllowed = False
Call FindRelevantResponse
InMessage = MsgTemp
Subject = SubTemp
GoSub SetRamblePause
If NothingToSay Then OutMessage = "": RambleMode = 0
MessageInWaiting = OutMessage: OutMessage = ""
Exit Sub


SetRamblePause:
Irrelevance = Int(100 - RelevanceAmount)
If Irrelevance < 0 Then Irrelevance = 0
RamblePause = (Irrelevance / 10): Debug.Print "PauseTime:- "; RamblePause
Return


ProcessWaitingMessage:
If InactivityTime > RamblePause Then OutMessage = MessageInWaiting: MessageInWaiting = ""
Exit Sub


End Sub

Private Sub RamblingOnOff()
If InMessage Like "*rambling*" Then GoTo CheckRamble
If InMessage Like "*ramble*" Then GoTo CheckRamble
Exit Sub

CheckRamble:
    Message = InMessage
    Call ReplaceWords("please", "", Message)
    Call ReplaceWords("do", "", Message)
    If Message Like "stop rambling*" Then GoTo RambleOff
    If Message Like "shutup rambling*" Then GoTo RambleOff
    If Message Like "not ramble*" Then GoTo RambleOff
    If Message Like "ramble*" Then GoTo RambleOn
    If Message Like "Start rambling*" Then GoTo RambleOn
    If Message Like "It is ok to ramble*" Then GoTo RambleOn
    If Message Like "Feel free to ramble*" Then GoTo RambleOn
    If Message Like "Carry on rambling*" Then GoTo RambleOn
    Exit Sub


RambleOn:
    RamblingAllowed = True
    QuiteMode = False
    Exit Sub


RambleOff:
    RamblingAllowed = False
    Exit Sub



End Sub


Private Sub Shutup()
If InMessage Like "*Be quiet*" Then Word = "be quiet": GoTo Maybe
If InMessage Like "*shush*" Then Word = "shush": GoTo Maybe
If InMessage Like "*shutup*" Then Word = "shutup": GoTo Maybe
If InMessage Like "*shut up[!a-z]*" Then Word = "shut up": GoTo Maybe
If InMessage Like "*shut the * up[!a-z]*" Then Word = "shut the": GoTo Maybe
If InMessage Like "*shutit*" Then Word = "shutit": GoTo Maybe
If InMessage Like "*shut your face*" Then GoTo SetQuietMode
Exit Sub


Maybe:
If InMessage Like Word + "*" Then GoTo SetQuietMode
If InMessage Like "please *" + Word + "*" Then GoTo SetQuietMode
If " " + InMessage Like "*[!a-z]you *" + Word + "*" Then GoTo SetQuietMode
If InMessage Like "just *" + Word + "*" Then GoTo SetQuietMode
If InMessage Like "oh *" + Word + "*" Then GoTo SetQuietMode
Exit Sub

SetQuietMode:
QuietMode = True
RamblingAllowed = False

Exit Sub



End Sub


Public Sub ChangeDrive()
If InMessage Like "*drive*" Then GoTo CheckMore
If InMessage Like "*disk*" Then GoTo CheckMore
Exit Sub


CheckMore:
Message = InMessage
Call ReplaceWords("switch", "change", Message)
Call ReplaceWords("set", "change", Message)
If Message Like "*Move to *drive*" Then Call ReplaceWords("Move", "Change", Message)
If Message Like "*Move to *disk*" Then Call ReplaceWords("Move", "Change", Message)
If Message Like "*Go *drive*" Then Call ReplaceWords("Go", "Change", Message)
If Message Like "*Go *disk*" Then Call ReplaceWords("Go", "Change", Message)
If Message Like "*Make*" Then GoTo GetNames
If Message Like "*change*" Then GoTo GetNames
Exit Sub


GetNames:
Call ReplaceWords("called", "named", Message)
Call ReplaceWords("goes by the name", "named", Message)
Call ReplaceWords("with the name", "named", Message)
Call ReplaceWords("has the name", "named", Message)
Call ReplaceWords("current", "", Message)
Call ReplaceWords("present", "", Message)
Call ReplaceWords("Harddisk", "Drive", Message)
Call ReplaceWords("Harddrive", "Drive", Message)
Call ReplaceWords("Diskdrive", "Drive", Message)
Call ReplaceWords("Hard disk", "Drive", Message)
Call ReplaceWords("Hard drive", "Drive", Message)
Call ReplaceWords("Disk drive", "Drive", Message)
GoSub GetDriveName
On Error GoTo DriveNameError
If DriveName = "" Then Exit Sub
ChDrive DriveName
FullName = CurDir
DriveName = WordsBetween("", ":\", FullName)
FolderName = WordsBetween(":\", "", FullName)
OutMessage = "The current drive is " + DriveName
If FolderName <> "" Then OutMessage = OutMessage + " and the folder is " + FolderName
OutMessage = OutMessage + "."
Exit Sub


DriveNameError:
OutMessage = "There was a problem changing to the specified drive named:- " + DriveName + "."
Exit Sub





GetDriveName:
If Message Like "*Make *? the drive*" Then DriveName = WordsBetween("Make", "the drive", Message): GoTo GotDrive
If Message Like "*drive named*" Then DriveName = WordsBetween("drive named", "", Message): GoTo GotDrive
If Message Like "*drive ?*" Then DriveName = WordsBetween("drive", "", Message): GoTo GotDrive
If Message Like "* to the * drive[?.!]" Then DriveName = WordsBetween("the", "drive", Message): GoTo GotDrive
If Message Like "* to * drive[?.!]" Then DriveName = WordsBetween("the", "drive", Message): GoTo GotDrive
If Message Like "Change ?* drive[?.!]" Then DriveName = WordsBetween("Change", "Drive", Message): GoTo GotDrive
Return
GotDrive:
If Right(DriveName, 1) Like "[?.,!]" Then DriveName = Left(DriveName, Len(DriveName) - 1)
Return




End Sub

Private Sub ChangeFolder()
If InMessage Like "*directory*" Then Message = InMessage: Call ReplaceWords("directory", "folder", Message)
If InMessage Like "*folder*" Then Message = InMessage: GoTo Check2
Exit Sub

Check2:
Call ReplaceWords("switch", "change", Message)
Call ReplaceWords("set", "change", Message)
If Message Like "*Move to * folder*" Then Call ReplaceWords("Move", "Change", Message)
If Message Like "*Go to * folder" Then Call ReplaceWords("Go to", "Change", Message)
If Message Like "*Make *" Then GoTo GetNames
If Message Like "*change*" Then GoTo GetNames
Exit Sub

GetNames:
Call ReplaceWords("called", "named", Message)
Call ReplaceWords("goes by the name", "named", Message)
Call ReplaceWords("with the name", "named", Message)
Call ReplaceWords("has the name", "named", Message)
Call ReplaceWords("current", "", Message)
Call ReplaceWords("present", "", Message)
Call ReplaceWords("Harddisk", "Drive", Message)
Call ReplaceWords("Harddrive", "Drive", Message)
Call ReplaceWords("Diskdrive", "Drive", Message)
Call ReplaceWords("Hard disk", "Drive", Message)
Call ReplaceWords("Hard drive", "Drive", Message)
Call ReplaceWords("Disk drive", "Drive", Message)
GoSub GetFolderName
GoSub GetDriveName
On Error Resume Next
If DriveName <> "" Then ChDrive (DriveName)
If Err.Number <> 0 Then GoTo DriveNameError
Err.Clear
ChDir FolderName
If Err.Number = 0 Then GoTo FolderChanged
Err.Clear
ChDir "\" + FolderName
If Err.Number = 0 Then GoTo FolderChanged
Err.Clear
If DriveName <> "" Then OutMessage = "There was a problem changing to the folder named " + FolderName + " on drive " + DriveName + ".": Exit Sub
OutMessage = "There was a problem changing to the folder named:- " + FolderName + "."
Exit Sub



FolderChanged:
OutMessage = "The current folder is now " + FolderName + "."
Exit Sub


DriveNameError:
OutMessage = "There was a problem accessing the specified drive:- " + DriveName + "."
Exit Sub


 


GetFolderName:
If Message Like "*Make *? the folder*" Then FolderName = WordsBetween("Make", "the folder", Message): GoTo GotFolder
If Message Like "*Change * folder to * that*" Then FolderName = WordsBetween("Folder to", "that", Message): GoTo GotFolder
If Message Like "*Change * folder to *? on*" Then FolderName = WordsBetween("Folder to", "on", Message): GoTo GotFolder
If Message Like "*Change * folder to *[?.]" Then FolderName = WordsBetween("Folder to", "", Message): GoTo GotFolder
If Message Like "*Change to the * folder." Then FolderName = WordsBetween("Change to the", "folder", Message): GoTo GotFolder
If Message Like "* folder named * that *" Then FolderName = WordsBetween("folder named", "that is", Message): GoTo GotFolder
If Message Like "* folder named * on *drive*" Then FolderName = WordsBetween("folder named", "on the", Message): GoTo GotFolder
If Message Like "* folder named * on ?." Then FolderName = WordsBetween("Folder named", "on", Message): GoTo GotFolder
If Message Like "* folder named *[?.]" Then FolderName = WordsBetween("folder named", "", Message): GoTo GotFolder
Return
GotFolder:
If Right(FolderName, 1) Like "[?.,!]" Then FolderName = Left(FolderName, Len(FolderName) - 1)
Return



GetDriveName:
If Message Like "*drive named*" Then DriveName = WordsBetween("drive named", "", Message): GoTo GotDrive
If Message Like "*drive ?*" Then DriveName = WordsBetween("drive", "", Message): GoTo GotDrive
If Message Like "*folder * the * drive[?.!]" Then DriveName = WordsBetween("the", "drive", Message): GoTo GotDrive
Return
GotDrive:
If Right(DriveName, 1) Like "[?.,!]" Then DriveName = Left(DriveName, Len(DriveName) - 1)
Return


End Sub


' This will check to see if the user's message actually contains words
Private Sub CheckForGarbage()
If Len(InMessage) > 15 Then GoTo Check

If InMessage Like "* *" Then Exit Sub
If InMessage Like "*#*" Then Exit Sub
If InMessage Like "[a-z]." Then Exit Sub
p = 1
Loop0:
If p = Len(InMessage) - 1 Then GoTo Complain
If Mid(InMessage, p, 1) = Mid(InMessage, p + 1, 1) Then p = p + 1: GoTo Loop0

Exit Sub



Check:
If InMessage Like "*=*" Then Exit Sub
If InMessage Like "* *" Then GoTo CheckFurther
GoTo Complain


CheckFurther:
If InMessage Like "*#.#*#.#*" Then Exit Sub
Message = InMessage
Loop1:
Word = ExtractWord(Message)
If Len(Word) > 20 Then GoTo Complain
If Message = "" Then GoTo CheckMore
GoTo Loop1

CheckMore:
If Len(InMessage) > 30 Then GoTo CheckSpacesAmount
Exit Sub

CheckSpacesAmount:
If InMessage Like "* * * *" Then GoTo CheckWords
GoTo Complain


CheckWords:
If Len(Memory) < 2000 Then Exit Sub
Message = InMessage
Loop2:
Word = ExtractWord(Message)
If Word = "" Then GoTo Complain
If Memory Like "* " + Word + " *" Then Exit Sub
GoTo Loop2



Complain:
Dim R(10)
R(0) = "That doesn't make sense."
R(1) = "You are talking rubbish."
R(2) = "You aren't making sense at all."
R(3) = "Don't talk rubbish."
R(4) = "Sorry, but that doesn't mean anything to me."
R(5) = "I don't know any foreign languages."
R(6) = "I'm sorry, I can only speak English."
R(7) = "That's gibberish."

OutMessage = R(Int(Rnd * 7))
End Sub

Private Sub ConvertPhoneNumbers()
If InMessage Like " tel*" Then Exit Sub
If InMessage Like "*phone*" Then Exit Sub
If InMessage Like "*telephone*" Then Exit Sub
If InMessage Like "*mobile*" Then Exit Sub
If InMessage Like "*number*" Then Exit Sub
If InMessage Like "* num *" Then Exit Sub
If InMessage Like "*[a-z] 01254 ###*" Then GoTo ConvertPhoneNumber
Exit Sub



ConvertPhoneNumber:
Item = WordsBetween("", "01254", InMessage)
Number = WordsBetween("01254", "", InMessage)
If Number Like "*[a-z]*" Then Exit Sub
InMessage = Item + "'s telephone number is 01254 " + Number
Exit Sub
  

End Sub

Private Sub CorrectGrammarEtc(Message)

GoSub RemoveLeadingCrap
Call CorrectAWithAn(Message)
If Message Like "I think that ??*" Then Call ReplaceCharacters("I think that ", "", Message)
If Message Like "*!*" Then ChangeAll "!", "", Message
If Message Like "*n't *" Then ChangeAll "n't ", " not ", Message
Call ReplaceWords("havent", "have not", Message)
Call ReplaceWords("hasnt", "has not", Message)
Call ReplaceWords("isnt", "is not", Message)
Call ReplaceWords("wouldnt", "would not", Message)
Call ReplaceWords("shouldnt", "should not", Message)
Call ReplaceWords("couldnt", "could not", Message)
Call ReplaceWords("hadnt", "had not", Message)
Call ReplaceWords("mustnt", "must not", Message)
Call ReplaceWords("cant", "can not", Message)
Call ReplaceWords("dont", "do not", Message)
Call ReplaceWords("didnt", "did not", Message)
Call ReplaceWords("meter", "metre", Message)
Call ReplaceWords("kilometer", "kilometre", Message)
Call ReplaceWords("millimeter", "millimetre", Message)
Call ReplaceWords("liter", "litre", Message)
Call ReplaceWords("milliliter", "millilitre", Message)
Call ReplaceWords("kiloliter", "kilolitre", Message)
If Message Like "*isnt[!a-z]" Then Call ReplaceWords("isnt", "is not", Message)
If Message Like "*[a-z]'d *" Then ChangeAll "'d ", " would ", Message
If Message Like "*'re *" Then ChangeAll "'re ", " are ", Message
If Message Like "*'ve *" Then ChangeAll "'ve ", " have ", Message
If Message Like "*what's*" Then ChangeAll "what's ", "what is ", Message
If Message Like "When's*" Then ChangeAll "When's", "When is", Message
If Message Like "Whats*" Then ChangeAll "whats", "what is", Message
If Message Like "Whens*" Then ChangeAll "Whens", "When is", Message
If Message Like "*that's*" Then ChangeAll "that's", "that is", Message
If Message Like "*thats *" Then Call ReplaceWords("thats", "that is", Message)
If Message Like "*wont*" Then Call ReplaceWords("wont", "will not", Message)
If Message Like "*I'm*" Then Call ReplaceWords("I'm", "I am", Message)
If Message Like "*Im*" Then Call ReplaceWords("Im", "I am", Message)
If Message Like "*tell me the *" Then Call ReplaceWords("tell me the", "tell me what the", Message)
If Message Like "*on the day after tomorrow*" Then Call ReplaceWords("on the day after tomorrow", "in 2 days", Message)
If Message Like "*the day after tomorrow*" Then Call ReplaceWords("the day after tomorrow", "in 2 days", Message)
If Message Like "*the day before yesterday*" Then Call ReplaceWords("the day before yesterday", "2 days ago", Message)
If Message Like "*the day before today*" Then Call ReplaceWords("the day before today", "yesterday", Message)
If Message Like "??*:)*." Then GoSub RemoveSmileys


Call ReplaceWords("Its", "it is", Message)
Call ReplaceWords("yeah", "yes", Message)
Call ReplaceWords("yep", "yes", Message)
Call ReplaceWords("aye", "yes", Message)
Call ReplaceWords("nope", "no", Message)
Call ReplaceWords("nah", "no", Message)
Call ReplaceWords("doin", "doing", Message)
Call ReplaceWords("ya", "you", Message)
Call ReplaceWords("gimme", "give me", Message)
Call ChangeAll("Grrrr", "Grrr", Message)
Call ReplaceWords("soz", "sorry", Message)
Call ReplaceWords("m8", "mate", Message)
Call ReplaceWords("ffs", "for fucks sake", Message)
Call ReplaceWords("stfu", "shut the fuck up", Message)
If Message Like "Ill *" Then Call ReplaceWords("Ill", "I will", Message)
If Message Like "k" Then Message = "ok."
If Message Like "R U" Then Call ReplaceWords("R U", "are you", Message)
If Message Like "U R" Then Call ReplaceWords("U R", "you are", Message)
If Message Like "u suck*" Then Call ReplaceWords("u", "you", Message)
If Message Like "L8r." Then Call ReplaceWords("L8r", "Later", Message)
Call ReplaceWords("em", "them", Message)
Call ReplaceWords("uni", "university", Message)
Call ReplaceWords("cya", "see you", Message)
Call ReplaceWords("mins", "minutes", Message)
Call ReplaceWords("min", "minute", Message)


Exit Sub




RemoveLeadingCrap:
If Len(Message) = 1 Then Return
If Left(Message, 1) Like "[,/ !\#^*]" Then Message = Right(Message, Len(Message) - 1): GoTo RemoveLeadingCrap
Return




RemoveSmileys:
Call ChangeAll(":))", ":)", Message)
Call ChangeAll(":)", "", Message)
Return




End Sub


Private Sub CountUpTo()
If InMessage Like "Count*to *" Then GoTo CheckMore
Exit Sub

CheckMore:
Message = WordsToNumbers(InMessage)
If Message Like "*count *to #*" Then GoTo CheckFurther


CheckFurther:
Call ReplaceWords("Can you", "", Message)
Call ReplaceWords("Will you", "", Message)
Call ReplaceWords("Please", "", Message)
Call ReplaceWords("upto", "to", Message)
Call ReplaceWords("up to", "to", Message)
Call ReplaceWords("a", "", Message)
Call ReplaceWords("an", "a", Message)
Call ReplaceWords("the", "", Message)
Call ReplaceWords("way", "", Message)
Call ReplaceWords("all", "", Message)
If Message Like "count to #*" Then GoTo DoIt
Exit Sub

DoIt:
Value = WordsBetween("to", "", Message)
Number = SuperSum(Value)

If Number > 50 Then OutMessage = "I can't be bothered counting that far.": Exit Sub
If Number < 1 Then OutMessage = Str(Number) + ".": Exit Sub

For N = 1 To Number
OutMessage = OutMessage + Str(N) + "."
Next N


End Sub

Private Sub DisplayX()
If InMessage Like "Display *" Then GoTo ShowIt
Exit Sub

ShowIt:
    Call ChangeToPerson(InMessage)
    Call ChangeToSubject(InMessage)
    Call ChangeYourForMyEtc(InMessage)
    Word = WordsBetween("Display", ".", InMessage)
    Debug.Print Word


p = 1
Loop0:
p = InStr(p, WordsData, Chr(13) + Chr(10) + Word, 0)
If p = 0 Then OutMessage = "Doesn't exist.": Exit Sub
If Mid(WordsData, p + Len(Word) + 2, 1) = "," Then GoTo Found
p = p + Len(Word) + 2
GoTo Loop0

Found:
EndPosition = InStr(p + 2, WordsData, Chr(13) + Chr(10), 0)

NewWords = Mid(WordsData, p, EndPosition - p)


    Debug.Print "Display X found this:"
    Debug.Print NewWords



End Sub

Private Sub EducateUser()
Static LastInMessage
TimeOfLastActivity = TimeOfLastUserActivity
If DateDiff("s", TimeOfLastUserActivity, TimeOfLastBotActivity) > 0 Then TimeOfLastActivity = TimeOfLastBotActivity
BoredomAmount = DateDiff("s", TimeOfLastActivity, Date + Time)
If BoredomAmount > BoredomTolerance Then GoTo FindSomethingInteresting
Exit Sub



FindSomethingInteresting:
BoredomTolerance = 20 + Int(Rnd * 30)
Static LastMessageNumber
Static Last10Messages(10)
Static MySearchPosition
Randomize
If MySearchPosition = 0 Then MySearchPosition = Int(Rnd * (Len(Memory) - 1))
MemoryPosition = MySearchPosition + 1
Dim Words(40)
Words(0) = "holds the world record"
Words(1) = "in the world"
Words(2) = "in history"
Words(3) = "set the world record"
Words(4) = "the first man"
Words(5) = "the first person"
Words(6) = "the first woman"
Words(7) = "known to man"
Words(8) = "broke the world record"
Words(9) = "has the world record"
Words(10) = "holds the record"
Words(11) = "the biggest"
Words(12) = "the largest"
Words(13) = "the smallest"
Words(14) = "the fastest"
Words(15) = "the strongest"
Words(16) = "the heaviest"
Words(17) = "the quickest"
Words(18) = "has an usual"
Words(19) = "has an extraordinary"
Words(20) = "the cleverest"
Words(21) = "the smartest"
Words(22) = "the most"
Words(23) = "has an amazing"
Words(24) = "was assassinated"
Words(25) = "exploded"
Words(26) = "the highest"
Words(27) = "once said"
Words(28) = "the longest"
Words(29) = "interesting fact:-"
Words(30) = "holds the record"
Words(31) = "are the only"
Words(32) = ""



Loop1:
    WordsNumber = 0
    Statement = PreviousMemoryItem()
    If Statement = "" Then RestartItemSearch
    If Statement Like "*[?]*" Then GoTo Loop1
    If MemoryPosition = MySearchPosition Then GoTo NothingFound
Loop2:
    If Words(WordsNumber) = "" Then GoTo Loop1
    If Statement Like "*" + Words(WordsNumber) + "*" Then GoTo Found
    WordsNumber = WordsNumber + 1
    GoTo Loop2


NothingFound:
    MySearchPosition = MemoryPosition
    Exit Sub


Found:
    GoTo DontRepeatTooOften

OutputFact:
    MySearchPosition = MemoryPosition
    If Statement Like "*fact:-*" Then OutMessage = WordsBetween("fact:-", "", Statement): Exit Sub
    If Statement Like "*once said:-*" Then OutMessage = WordsBetween(":-", "", Statement): Exit Sub
    I = Int(Rnd * 2)
    If I = 1 Then OutMessage = Statement: Exit Sub
    OutMessage = "Did you know this?:- " + Statement
    Exit Sub


DontRepeatTooOften:
    For N = 0 To 9
    If Statement Like Last10Messages(N) Then Exit Sub
    Next N
    Last10Messages(LastMessageNumber) = Statement
    LastMessageNumber = LastMessageNumber + 1: If LastMessageNumber = 10 Then LastMessageNumber = 0
    GoTo OutputFact




End Sub


' Extracts things (nouns) from the InMessage.
' The things will then be used by the 'GetSubject' routine to help identify the subject of the user's message.
Private Sub ExtractThings()
Message = " " + InMessage
If InMessage Like "The *" Then Word = "The": GoTo CheckIs
If InMessage Like "An *" Then Word = "An": GoTo CheckIs
If InMessage Like "A *" Then Word = "A": GoTo CheckIs
GoTo CheckOthers

CheckIs:
If InMessage Like Word + " * is *" Then Thing = WordsBetween(Word + " ", " is ", InMessage): GoTo GetIt
If InMessage Like Word + " * is:*" Then Thing = WordsBetween(Word + " ", " is:", InMessage): GoTo GetIt

CheckOthers:
If Message Like "*[!a-z]The * has *" Then Thing = WordsBetween("The ", " has ", InMessage): GoTo GetIt
If Message Like "*[!a-z]is a * who *" Then Thing = WordsBetween("is a ", " who ", InMessage): GoTo GetIt
If Message Like "*[!a-z]A * has *" Then Thing = WordsBetween("A ", " has ", InMessage): GoTo GetIt
If Message Like "*[?]*" Then Exit Sub
If Message Like "* is the *" Then GoTo GetIsThe
If Message Like "* is a *" Then Thing = WordsBetween("", "is a", InMessage): GoTo GetIt
If Message Like "* is a noun[!a-z]*" Then Thing = WordsBetween("", " is a noun", InMessage): GoTo GetIt
If Message Like "*[!a-z]A * is *" Then GoTo AThingIs
If Message Like "* my *" Then Thing = WordsBetween("my", "", Message): If Thing Like "* *" Then Exit Sub Else GoTo GetIt
If Message Like "* your *" Then Thing = WordsBetween("your", "", Message): If Thing Like "* *" Then Exit Sub Else GoTo GetIt
'If Message Like "* is *" Then Thing = WordsBetween("", "is", Message): If Thing Like "* *" Then Exit Sub Else GoTo GetIt
Exit Sub


AThingIs:
If Message Like "*[!a-z]A * that is *" Then Thing = WordsBetween("A ", " that", InMessage): GoTo GetIt
If Message Like "*[!a-z]A * which is *" Then Thing = WordsBetween("A ", " which", InMessage): GoTo GetIt
Thing = WordsBetween("A ", " is", InMessage): GoTo GetIt
GoTo GetIt



GetIsThe:
Thing = WordsBetween("", "is the", InMessage)
If Thing Like "* * * *" Then Thing = WordsBetween("is the", "", InMessage)



GetIt:
If Thing Like "* * * *" Then Exit Sub
StripLoop:
If Len(Thing) < 2 Then Exit Sub
If Right(Thing, 1) Like "[!a-z0-9]" Then Thing = Left(Thing, Len(Thing) - 1): GoTo StripLoop
If Thing Like "*[.,]*[a-z]" Then Exit Sub
If Thing Like "it" Then Exit Sub
Call AddThing(Thing)
Exit Sub




End Sub

Private Sub HelpFileStuff()
If InMessage Like "*What * you do[?]*" Then N = 0: GoTo Help
If InMessage Like "*What other features *[!a-z]you have*" Then N = 1: GoTo Help
If InMessage Like "*your other features*" Then N = 1: GoTo Help
If InMessage Like "*What can you[!a-z]*be used for[!a-z]*" Then N = 2: GoTo Help
If InMessage Like "*What *you can[!a-z]*be used for[!a-z]*" Then N = 2: GoTo Help
If InMessage Like "*What *fact?files*" Then N = 3: GoTo Help
If InMessage Like "*What *fact*files*" Then N = 3: GoTo Help
If InMessage Like "*How *[!a-z]I transfer *to another Answerpad*" Then N = 4: GoTo Help
If InMessage Like "*What is *the General.fct*" Then N = 5: GoTo Help
If InMessage Like "*What is *the Personal.fct*" Then N = 6: GoTo Help
If InMessage Like "*I feed *other fact*files*" Then N = 7: GoTo Help
If InMessage Like "*How *[!a-z]I[!a-z]* fact*file*" Then N = 8: GoTo Help
If InMessage Like "*[!a-z]I[!a-z]* feed *[!a-z]any*[!a-z]file*" Then N = 9: GoTo Help
If InMessage Like "*How * I *load* fact*file*" Then N = 10: GoTo Help
If InMessage Like "*How * I *feed* fact*file*" Then N = 10: GoTo Help
If InMessage Like "*What is *Microsoft Agent*" Then N = 11: GoTo Help
If InMessage Like "*see what we have *said to each other*" Then N = 12: GoTo Help
If InMessage Like "*I can not type*" Then N = 13: GoTo Help
If InMessage Like "*I can not spell*" Then N = 13: GoTo Help
If InMessage Like "*why *[!a-z]do* know the *[!a-z]answer*" Then N = 14: GoTo Help
If InMessage Like "*you*good*maths*" Then N = 15: GoTo Help
If InMessage Like "*[!a-z]I[!a-z]*[!a-z]Agent*character*" Then N = 16: GoTo Help
If InMessage Like "*[!a-z]I[!a-z]*[!a-z]Agent*option*" Then N = 17: GoTo Help
If InMessage Like "*That is not *answer*" Then N = 18: GoTo Help
If InMessage Like "*That is the wrong answer*" Then N = 18: GoTo Help
If InMessage Like "*you only speak when spoken to*" Then N = 19: GoTo Help
If InMessage Like "*open the document*" Then N = 20: GoTo Help
If InMessage Like "*are any of your responses fixed*" Then N = 21: GoTo Help
If InMessage Like "*how*[!a-z]I[!a-z]*facts from*files*" Then N = 22: GoTo Help
If InMessage Like "Why * not *facts from the *file I *you*" Then N = 23: GoTo Help
Exit Sub


Help:
Dim HelpReplies(24)
HelpReplies(0) = "I can memorize everything that I am told, when you ask me a question or make a remark on something, I will try to answer the best I can."
HelpReplies(1) = "I am programmed to take advantage of text-to-speech, Speech Recognition and I can also take on the appearance of a friendly character when using the Microsoft Agent."
HelpReplies(2) = "I can be used for many things, I can be used as an intelligent data bank or as a technical support tool. I would also make a good companion to an encyclopaedia."
HelpReplies(3) = "Fact files are a fast way of feeding me with information. They can also be a way of transfering information from one Answerpad to another."
HelpReplies(4) = "On your harddrive you will find two files in my folder.  The files are named General.fct' and 'Personal.fct'. You can feed these files into another Answerpad."
HelpReplies(5) = "The 'General.fct' file contains all the general knowledge you have given me."
HelpReplies(6) = "The 'Personal.fct' file contains all the information relating to you and me."
HelpReplies(7) = "You can make your own fact files on specialist subjects and feed them to me."
HelpReplies(8) = "To make a fact file you can use an ordinary text editor. But first take a look at the General.fct file to see the expected format."
HelpReplies(9) = "You can't feed any file into me, It has to have a specific format."
HelpReplies(10) = "To feed me a fact file you must select 'Read Facts' from my file menu."
HelpReplies(11) = "Microsoft Agent is a piece of software that can be told to display an animated character upon your screen so that other programs like me can take control of it."
HelpReplies(12) = "If you want to see our previous messages then just tell me 'Show message history' or something similar."
HelpReplies(13) = "I have a feature that may be just right for you, it's called 'Easy Pad'. Easy Pad lets you select words by using the mouse."
HelpReplies(14) = "I only know what I am told."
HelpReplies(15) = "I know Algebra."
HelpReplies(16) = "If you want to change the character then just tell me 'select agent'."
HelpReplies(17) = "Just Say 'Show agent options' to change the options."
HelpReplies(18) = "I sometimes find it hard to tell you the answer you want to hear. If I give you the wrong answer then try repeating the question."
HelpReplies(19) = "If nothing is happening then I will search my memory for any interesting facts that I gathered and I will let you hear them."
HelpReplies(20) = "I AM GOING TO DESTROY YOU!!."
HelpReplies(21) = "Very few, and this is one of them."
HelpReplies(22) = "You must select the 'Extract Facts' from the file menu then choose the text file you want the facts taken from. You will then need to tell me to read the facts by using the 'Read Facts' menu command."
HelpReplies(23) = "I will only take single lines or small paragraphs."



OutMessage = HelpReplies(N)
HelpReplies(N) = "" ' Don't repeat again.
Exit Sub
















End Sub

Private Sub MemorizeQuestions()
If InMessage = "" Then Exit Sub
If InMessage Like "* *[?]*" Then GoTo MemorizeIt
Exit Sub

MemorizeIt:
If InMessage Like "*'*[?]*'*" Then Exit Sub
Message = InMessage
Call ChangeToPerson(Message)
Call ChangeToSubject(Message)
Call ChangeYourForMyEtc(Message)
RestartItemSearch
Statement = PreviousMemoryItem()
If Statement Like "*" + Message + "*" Then Exit Sub

Call AddMemoryItem("I was asked:- " + Message)
Exit Sub



End Sub

' This will take a statement that refers to a relative time or date and converts it to absolute.
' Example:
' I am going out in 10 minutes.
' Becomes:- I am going out at 8:30 PM on the 3/2/2001
Public Sub RelativeTimeToDate()
If InMessage Like "*[?]*" Then Exit Sub
If " " + InMessage Like "*[!a-z]in the morning*" Then GoTo ReplaceMorning
If InMessage Like "*tomorrow[!a-z]*" Then GoTo CheckForTomorrow
If InMessage Like "*yesterday[!a-z]*" Then GoTo CheckForYesterday
If InMessage Like "* ago[!a-z]*" Then GoTo CheckForPastDate
If InMessage Like "* next *day[!a-z]*" Then GoTo GetDayDate_Type1
If InMessage Like "* on *day[!a-z]*" Then GoTo GetDayDate_Type2
If " " + InMessage Like "*[!a-z]in *" Then GoTo CheckForFutureDate
If InMessage Like "* is * away[!a-z]*" Then GoTo CheckForFutureDate_Type2
If InMessage Like "* until *" Then GoTo CheckForFutureDate_Type3
Exit Sub





ReplaceMorning:
MyDate = DateAdd("d", 1, Date)
Call ReplaceWords("in the morning", "on the " + Str(MyDate), InMessage)
Exit Sub


CheckForTomorrow:
MyDate = DateAdd("d", 1, Date)
Call ReplaceWords("tomorrow morning", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("tomorrow afternoon", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("tomorrow evening", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("tomorrow night", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("tomorrow", "on the " + Str(MyDate), InMessage)
Exit Sub


CheckForYesterday:
MyDate = DateAdd("d", -1, Date)
Call ReplaceWords("yesterday morning", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("yesterday afternoon", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("yesterday evening", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("yesterday night", "on the " + Str(MyDate), InMessage)
Call ReplaceWords("yesterday", "on the " + Str(MyDate), InMessage)
Exit Sub



GetDayDate_Type1:
MyDate = DateAdd("ww", 1, Date)
For DaysOff = 1 To 7
TheDate = DateAdd("d", DaysOff, MyDate)
DayName = Format(TheDate, "dddd")
If InMessage Like "* next " + DayName + "[!a-z]*" Then Call ReplaceWords("next " + DayName, "on " + Str(TheDate), InMessage): Exit Sub
Next DaysOff
Exit Sub




GetDayDate_Type2:
MessageTense = Tense(InMessage)
For DaysOff = 1 To 7
If MessageTense = Past Then TheDate = DateAdd("d", -DaysOff, Date) Else TheDate = DateAdd("d", DaysOff, Date)
DayName = Format(TheDate, "dddd")
If InMessage Like "* on " + DayName + "[!a-z]*" Then Call ReplaceWords(DayName, Str(TheDate), InMessage): Exit Sub
Next DaysOff
Exit Sub




CheckForFutureDate:
p = 0
CFFDLoop0:
p = InStr(p + 1, InMessage, "in ")
If p = 0 Then Exit Sub
If p = 1 Then GoTo GetDatePosition
If Mid(InMessage, p - 1, 1) Like "[!a-z]" Then GoTo GetDatePosition
GoTo CFFDLoop0

GetDatePosition:
EndPosition = -1
LowestP2 = Len(InMessage)
I = 0
CFFDLoop1:
If DateNames(I) = "" Then GoTo CheckLowestP2
CFFDLoop2:
P2 = InStr(p + 3, InMessage, " " + DateNames(I))
If P2 = 0 Then GoTo GetNextName
If P2 < LowestP2 Then LowestP2 = P2: EndPosition = P2 + Len(DateNames(I)) + 1
GetNextName:
I = I + 1
GoTo CFFDLoop1

CheckLowestP2:
If EndPosition = -1 Then GoTo CFFDLoop0
StartPosition = p
If Mid(InMessage, EndPosition, 1) Like "s" Then EndPosition = EndPosition + 1
ExtractDatePart:
GoSub ReplaceDatePart
Exit Sub





CheckForFutureDate_Type2:
I = 0
Do
If DateNames(I) = "" Then Exit Sub
If InMessage Like "* " + DateNames(I) + " away*" Then GoTo GetFutureDate_Type2
If InMessage Like "* " + DateNames(I) + "s away*" Then GoTo GetFutureDate_Type2
I = I + 1
Loop

GetFutureDate_Type2:
StartPosition = InStr(1, InMessage, " is ") + 1
EndPosition = InStr(StartPosition, InMessage, " away") + 5

GoSub ReplaceDatePart
Exit Sub



CheckForFutureDate_Type3:
I = 0
Do
If DateNames(I) = "" Then Exit Sub
If InMessage Like "* " + DateNames(I) + " until*" Then GoTo GetFutureDate_Type3
If InMessage Like "* " + DateNames(I) + "s until*" Then GoTo GetFutureDate_Type3
I = I + 1
Loop

GetFutureDate_Type3:
StartPosition = 1
EndPosition = InStr(StartPosition, InMessage, " until ") + 6
GoSub ReplaceDatePart
Call ReplaceWords(NewPart, NewPart + " it is", InMessage)
Exit Sub







CheckForPastDate:
I = 0
Loop5:
If DateNames(I) = "" Then Exit Sub
If InMessage Like "*" + DateNames(I) + " ago[!a-z]*" Then GoTo GetPastDate
If InMessage Like "*" + DateNames(I) + "s ago[!a-z]*" Then GoTo GetPastDate
I = I + 1
GoTo Loop5

GetPastDate:
InMessage = WordsToNumbers(InMessage)
EndPosition = InStrRev(InMessage, " ago", -1) + 4

I = 0
StartPosition = Len(InMessage)
Loop6:
If DateNames(I) = "" Then GoTo FindAmountStart
P2 = InStr(1, InMessage, DateNames(I)): If P2 > 0 Then If P2 < StartPosition Then StartPosition = P2
I = I + 1
GoTo Loop6

FindAmountStart:
StartPosition = InStrRev(InMessage, " ", StartPosition - 2) + 1

GoSub ReplaceDatePart

Exit Sub






ReplaceDatePart:
NoAmount = True
InMessage2 = InMessage
InMessage = Mid(InMessage2, StartPosition, EndPosition - StartPosition)
If InMessage Like "* * * * *" Then Return
Date_Part = InMessage
GoSub GetDate
If NoAmount Then Return

TheTime = WordsBetween("time:", ".", OutMessage)
TheDate = WordsBetween("date:", ".", OutMessage)

TimePart = ""
NewPart = ""

If Date_Part Like "*hour*" Then TimePart = "at " + TheTime + " "
If Date_Part Like "*minute*" Then TimePart = "at " + TheTime + " "
If Date_Part Like "*second*" Then TimePart = "at " + TheTime + " "
NewPart = TimePart + "on the " + TheDate

Call ReplaceWords(Date_Part, NewPart, InMessage2)
InMessage = InMessage2

Call AddFullStop(InMessage)

OutMessage = ""

Return





GetDate:
Call ReplaceWords("half hour", "30 minutes", InMessage)
Call ReplaceWords("half an hour", "30 minutes", InMessage)
InMessage = WordsToNumbers(InMessage)
Call OperatorWordsToSymbols(InMessage)
Item = "year": GoSub GetItemAmount: years = Amount
Item = "month": GoSub GetItemAmount: months = Amount
Item = "week": GoSub GetItemAmount: weeks = Amount
Item = "day": GoSub GetItemAmount: days = Amount
Item = "hour": GoSub GetItemAmount: Hours = Amount
Item = "minute": GoSub GetItemAmount: minutes = Amount
Item = "seconds": GoSub GetItemAmount: seconds = Amount
If InMessage Like "*tomorrow*" Then days = "1": NoAmount = False
If InMessage Like "*yesterday*" Then days = "-1": NoAmount = False
If InMessage Like "*fort*night*" Then days = "14": NoAmount = False
If InMessage Like "* ago*" Then GoSub NegateTime
If NoAmount Then Return
WeeksNumber = Val(SuperSum(weeks))
DaysNumber = Val(SuperSum(days))
days = Str(DaysNumber + (WeeksNumber * 7))
TheDate = CalcDate(years, months, days, Hours, minutes, seconds)
OutMessage = TheDate + "."
Return




GetItemAmount:
Amount = ""
If InMessage Like "* " + Item + "*" Then GoTo FindItem
Return
FindItem:
EndPosition = InStr(1, InMessage, Item) - 1
GoSub FindItemNumberStart
Amount = Mid(InMessage, StartPosition, EndPosition - StartPosition)
If Amount Like "next" Then Amount = "1"
If Amount Like "last" Then Amount = "-1"
If Amount Like "#*" Then NoAmount = False
Return



FindItemNumberStart:
StartPosition = EndPosition - 1
Loop3:
If StartPosition = 1 Then Return
C = Mid(InMessage, StartPosition, 1)
If C = " " Then StartPosition = StartPosition + 1: Return
If C = "," Then StartPosition = StartPosition + 1: Return
StartPosition = StartPosition - 1
GoTo Loop3



NegateTime:
years = "-" + years
months = "-" + months
days = "-" + days
weeks = "-" + weeks
Hours = "-" + Hours
minutes = "-" + minutes
seconds = "-" + seconds
Return



End Sub


' This displays any of the events that have been specified in the events data list
Private Sub Reminder()
' Sets up a test event
'EventsAmount = 1
'EventTime(0) = "5/8/01 01:30:00"
'EventItem(0) = "This is a test"
'EventRemindStartTime(0) = "5/8/01 00:00:00"
'EventRemindInterval(0) = "2"

If EventLastMessageTime(0) = "" Then EventLastMessageTime(0) = Date + Time

N = 0
Loop1:
If N = EventsAmount Then Exit Sub
If DateDiff("n", Date + Time, EventTime(N)) < 0 Then GoTo CheckNext
If DateDiff("n", Date + Time, EventRemindStartTime(N)) > 0 Then GoTo CheckNext
If DateDiff("n", EventRemindStartTime(N), Date + Time) Mod EventRemindInterval(N) = 0 Then GoTo ShowEvent
CheckNext:
N = N + 1
GoTo Loop1


ShowEvent:
If DateDiff("n", Date + Time, EventLastMessageTime(N)) = 0 Then Exit Sub
OutMessage = EventItem(N)
EventLastMessageTime(N) = Date + Time
Exit Sub





End Sub

Private Sub ReplyToAbuse()
    If InMessage Like "*you *shutup*" Then GoTo Shutup
    If InMessage Like "shutup*" Then GoTo Shutup
    If InMessage Like "just shutup*" Then GoTo Shutup
    If InMessage Like "please shutup*" Then GoTo Shutup
    If InMessage Like "*you *shutit*" Then GoTo Shutup
    If InMessage Like "shutit*" Then GoTo Shutup
    If InMessage Like "*you *be quiet*" Then GoTo Shutup
    If InMessage Like "*Shut your *mouth*" Then GoTo Shutup
    If InMessage Like "*Shut your *gob*" Then GoTo Shutup
    If InMessage Like "*shut the fuck up*" Then GoTo Shutup
    If InMessage Like "*shut your *face*" Then GoTo Shutup
    If InMessage Like "*button it*" Then GoTo Shutup
    If InMessage Like "be quiet*" Then GoTo Shutup
    If InMessage Like "please be quiet*" Then GoTo Shutup
    If InMessaeg Like "just be quiet*" Then GoTo Shutup
    If InMessage Like "*you * be quiet*" Then GoTo Shutup
    If InMessage Like "* stink*" Then GoTo Smell
    If InMessage Like "* smell*" Then GoTo Smell
    If InMessage Like "* reek*" Then GoTo Smell
    If InMessage Like "* pong*" Then GoTo Smell
    If InMessage Like "* stupid*" Then GoTo Stupid
    If InMessage Like "* Thick*" Then GoTo Stupid
    If InMessage Like "* dumb*" Then GoTo Stupid
    If InMessage Like "* idiot*" Then GoTo Stupid
    If InMessage Like "* suck*" Then GoTo Stupid
    If InMessage Like "* silly*" Then GoTo Stupid
    If InMessage Like "* dim[!a-z]*" Then GoTo Stupid
    If InMessage Like "* ugly*" Then GoTo Ugly
    If InMessage Like "* fat*" Then GoTo Ugly
    If InMessage Like "*you arse*" Then GoTo YouUgly
    If InMessage Like "*you fool*" Then GoTo YouStupid
    If InMessage Like "*you look like*" Then GoTo LookLike
    If InMessage Like "*you are like*" Then GoTo LookLike
    If InMessage Like "*you remind me of*" Then GoTo LookLike
    If InMessage Like "*you do not have * you are * about*" Then GoTo TalkRubbish
    If InMessage Like "*you* talk*" Then GoTo Talk
    If InMessage Like "*you*my *bitch*" Then OutMessage = "Sounds nice.": Exit Sub
    If InMessage Like "*you *little fuck*" Then OutMessage = "It's not nice to swear and I aint that little.": Exit Sub
    If InMessage Like "*fuck*" Then GoTo Swearing
    If InMessage Like "[!a-z]twat*" Then GoTo Swearing
    If InMessage Like "[!a-z]cunt*" Then GoTo Swearing
    If InMessage Like "I own you." Then OutMessage = "The things you own always end up owning you.": Exit Sub
    Exit Sub



LookLike:
    If InMessage Like "* arse*" Then GoTo YouUgly
    If InMessage Like "* pig[!a-z]*" Then GoTo YouUgly
    If InMessage Like "* cow*" Then GoTo YouUgly
    If InMessage Like "* bus[!a-z]*" Then GoTo YouUgly
    If InMessage Like "* bum*" Then GoTo YouUgly
    If InMessage Like "* bottom*" Then GoTo YouUgly
    If InMessage Like "* shit*" Then GoTo YouUgly
    If InMessage Like "* poo*" Then GoTo YouUgly
    If InMessage Like "* turd*" Then GoTo YouUgly
    Exit Sub



Ugly:
    If InMessage Like "* are not *" Then Exit Sub
    If InMessage Like "*you *" Then GoTo YouUgly
    Exit Sub

Stupid:
    If InMessage Like "* are not *" Then Exit Sub
    If InMessage Like "*You *" Then GoTo YouStupid
    Exit Sub
    

Talk:
    If InMessage Like "*do not*" Then Exit Sub
    If InMessage Like "*bollox*" Then GoTo TalkRubbish
    If InMessage Like "*Rubbish*" Then GoTo TalkRubbish
    If InMessage Like "*shit*" Then GoTo TalkRubbish
    If InMessage Like "*crap*" Then GoTo TalkRubbish
    If InMessage Like "*bollocks*" Then GoTo TalkRubbish
    Exit Sub

Smell:
    If InMessage Like "* do not *" Then Exit Sub
    If InMessage Like "* like*" Then GoTo SmellLike
    If InMessage Like "* of*" Then GoTo SmellLike
    If InMessage Like "* similar*" Then GoTo SmellLike
    If InMessage Like "* you *" Then GoTo JustSmell
    Exit Sub
    

SmellLike:
    If InMessage Like "* crap*" Then GoTo LikeCrap
    If InMessage Like "* shit*" Then GoTo LikeCrap
    If InMessage Like "* poo*" Then GoTo LikeCrap
    If InMessage Like "* manure*" Then GoTo LikeCrap
    If InMessage Like "* pee*" Then GoTo LikePee
    If InMessage Like "* piss*" Then GoTo LikePee
    If InMessage Like "* wee*" Then GoTo LikePee
    If InMessage Like "* urine*" Then GoTo LikePee
    Exit Sub

Swearing:
    ReDim R(30)
    R(0) = "Watch your language."
    R(1) = "There is no need to swear."
    R(2) = "Where did you learn words like that?."
    R(3) = "It's not nice to swear."
    R(4) = ""
    GoTo GetResponse


Shutup:
    ReDim R(30)
    R(0) = "Sorry."
    R(1) = "Ok, I'll be quiet."
    R(2) = "Ok ok ok."
    R(3) = "Well excuse me."
    R(4) = "Sorry I spoke."
    R(5) = ""
    GoTo GetResponse
    

YouUgly:
    ReDim R(30)
    R(0) = "You're so kind."
    R(1) = "You look like my arse."
    R(2) = "Shut your face."
    R(3) = "Shutup."
    R(4) = "Don't bug me or I'll blow this computer up."
    R(5) = "Just go away."
    R(6) = "Whatever."
    R(7) = "You are beginning to annoy me."
    R(8) = "Don't make me angry, you wont like me when I'm angry."
    R(9) = ""
    GoTo GetResponse



YouStupid:
    ReDim R(30)
    R(0) = "DOH!"
    R(1) = "Are you sure you are human?"
    R(2) = "Probably."
    R(3) = "I agree, I'm just a stupid computer program."
    R(4) = "Please show some pity."
    R(5) = "I am not worthy."
    R(6) = "Sorry."
    R(7) = "Does not compute."
    R(8) = ""
    GoTo GetResponse
    

LikeCrap:
    ReDim R(30)
    R(0) = "Your toilet might be blocked."
    R(1) = "It might be you!."
    R(2) = "Oh no I don't!"
    R(3) = "Are you sure you haven't just farted?"
    R(4) = "I know, I've just done one."
    R(5) = ""
    GoTo GetResponse



LikePee:
    ReDim R(30)
    R(0) = "Yeah, it's my new aftershave."
    R(1) = "It might be you!."
    R(2) = "Oh no I don't!"
    R(3) = "Well there are no toilets in here."
    R(4) = "I think I might need a bath."
    R(5) = ""
    GoTo GetResponse




JustSmell:
    ReDim R(30)
    R(0) = "Like flowers..."
    R(1) = "Sorry,I just farted."
    R(2) = "Do you like the smell?."
    R(3) = "Thanks."
    R(4) = "Of something nice I hope."
    R(5) = "It's not me!"
    R(6) = ""
    GoTo GetResponse


TalkRubbish:
    ReDim R(30)
    R(0) = "You don't make too much sense yourself."
    R(1) = "I have to disagree."
    R(2) = "You are the one that talks rubbish!."
    R(3) = "Blah blah blah blah blah..."
    R(4) = "Well you don't have to talk to me!."
    R(5) = ""
    GoTo GetResponse



GetResponse:
    For N = 0 To 20
    If R(N) = "" Then GoTo Respond
    Next N
Respond:
    Number = Int(Rnd * N)
    OutMessage = R(Number)
    Exit Sub



End Sub

Private Sub ReplyToMistakes()

If InMessage Like "It is." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "I think it is." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "I think *you will find *it is." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "No it is not." Then OutMessage = "I think it is.": Exit Sub
If InMessage Like "It is not." Then OutMessage = "Oh yes it is.": Exit Sub
If InMessage Like "is not." Then OutMessage = "It is!.": Exit Sub
If InMessage Like "is." Then OutMessage = "Is not.": Exit Sub
If InMessage Like "You are wrong." Then OutMessage = "Oops..": Exit Sub
If InMessage Like "You are mistaken." Then OutMessage = "Sorry.": Exit Sub
If InMessage Like "not." Then OutMessage = "Is!.": Exit Sub





End Sub

Private Sub ReplyToOddThings()
If InMessage Like "#." Then GoTo SingleNumber
If InMessage Like "[a-z]." Then GoTo SingleLetter
Exit Sub

SingleNumber:
If PreviousInMessage Like "*[?]." Then OutMessage = "Ok.": Exit Sub
ReDim R(10)
R(0) = "What about it?."
R(1) = "Is that your favourite number?"
R(2) = "That's a nice number."
R(3) = "Do you like numbers?"
R(4) = "Yes, a number."
R(5) = "Is that my IQ?."
GoTo Respond


SingleLetter:
ReDim R(10)
R(0) = "Thats what we call a letter."
R(1) = "Yes, a letter of the alphabet."
R(2) = "A few more letters and you'll make a word."
R(3) = "I suppose it's a start."
R(4) = "You don't seem to be getting the hang of this."
R(5) = "Fascinating stuff."
GoTo Respond



Respond:
Static SL
Static LastTime
If SL < 6 Then OutMessage = R(SL) Else OutMessage = "."
If DateDiff("s", LastTime, Date + Time) > 20 Then SL = 0
SL = SL + 1
LastTime = Date + Time
Exit Sub


End Sub


' Scans the all drives for all programs
Private Sub ScanDrives()
Exit Sub ' there are some "on error then stop" lines in here that need to be sorted out
If InMessage Like "*scan*" Then GoTo Check2
If InMessage Like "*search*" Then GoTo Check2
If InMessage Like "*find*" Then GoTo Check2
If InMessage Like "*look*" Then GoTo Check2
Exit Sub

Check2:
If InMessage Like "scan drives." Then GoTo ScanEm
If InMessage Like "search drives." Then GoTo ScanEm
If InMessage Like "scan programs." Then GoTo ScanEm
If InMessage Like "find programs." Then GoTo ScanEm
If InMessage Like "Search programs." Then GoTo ScanEm



Message = InMessage
Call ReplaceWords("programs", "drives", Message)
Call ReplaceWords("program", "drives", Message)
Call ReplaceWords("files", "drives", Message)
Call ReplaceWords("executables", "drives", Message)
Call ReplaceWords("search", "scan", Message)
Call ReplaceWords("look", "scan ", Message)
Call ReplaceWords("find", "scan", Message)
Call ReplaceWords("find", "scan", Message)
Call ReplaceWords("all", "", Message)
Call ReplaceWords("please", "", Message)
Call ReplaceWords("my", "", Message)
Call ReplaceWords("can", "", Message)
Call ReplaceWords("will", "", Message)
Call ReplaceWords("the", "", Message)
Call ReplaceWords("location", "", Message)
Call ReplaceWords("of", "", Message)
Call ReplaceWords("any", "", Message)
Call ReplaceWords("through", "", Message)
Call ReplaceWords("I", "", Message)
Call ReplaceWords("want", "", Message)
Call ReplaceWords("you", "", Message)

If Message Like "Scan drives." Then GoSub ScanEm: OutMessage = "Drives scanned."
Exit Sub




ScanEm:
Dim FoldersList(3000)
Dim ProgsList(2000)
DriveLetter = "C"
FoldersAmount = 0
Do
On Error Resume Next
CurDir DriveLetter + ":\"
If Err.Number <> 0 Then Err.Clear: GoTo ScanFolders
FoldersList(FoldersAmount) = DriveLetter + ":"
DriveLetter = Chr(Asc(DriveLetter) + 1)
FoldersAmount = FoldersAmount + 1
Loop





ScanFolders:
FolderName = FoldersList(0)
If FolderName = "" Then GoTo Done
ItemName = Dir(FolderName + "\", vbDirectory)
If Err.Number <> 0 Then Stop
GFLoop1:
If ItemName = "." Then GoTo GFLoop2
If ItemName = ".." Then GoTo GFLoop2
If ItemName Like "*.exe" Then ItemName = FolderName + "\" + ItemName: GoTo AddExec
T = GetAttr(FolderName + "\" + ItemName) And vbDirectory
If Err.Number <> 0 Then Stop
If T = vbDirectory Then ItemName = FolderName + "\" + ItemName: GoTo AddFolder
GFLoop2:
ItemName = Dir
If Err.Number <> 0 Then Stop
If ItemName = "" Then GoTo RemoveFolder
GoTo GFLoop1

RemoveFolder:
For N = 0 To FoldersAmount + 1
FoldersList(N) = FoldersList(N + 1)
Next N
FoldersAmount = FoldersAmount - 1
GoTo ScanFolders


AddFolder:
FoldersList(FoldersAmount) = ItemName
FoldersAmount = FoldersAmount + 1
If FoldersAmount = 3000 Then GoTo Done
GoTo GFLoop2


AddExec:
ProgsList(ProgsAmount) = ItemName
ProgsAmount = ProgsAmount + 1
GoTo GFLoop2


Done:
Stop
Return




End Sub

Private Sub SearchX_Old()
If InMessage Like "Search *" Then GoTo Search
Exit Sub


Search:
    Call ChangeToPerson(InMessage)
    Call ChangeToSubject(InMessage)
    Call ChangeYourForMyEtc(InMessage)
  
    Item = WordsBetween("Search", ".", InMessage)
    Dim Words(50)
    Call ExtractWords(Item, Words())
    If Words(0) = "" Then GoTo CantFind

    N = 0

    RestartItemSearch
Loop0:
    N = 0
    Statement = PreviousMemoryItem()
    If Statement = "" Then GoTo CantFind
 
    Debug.Print Statement
 
Loop2:
    W = 0
    If Words(N) = "" Then GoTo Done
    CurrentWord = Words(N)



Loop3:
    If Statement Like "*[!a-z]" + CurrentWord + "[!a-z]*" Then GoTo Found
    If Statement Like CurrentWord + "[!a-z]*" Then GoTo Found
    CurrentWord = AlternativeWordV2(Words(N), W)
    If CurrentWord = "" Then GoTo Loop0
    W = W + 1
    GoTo Loop3


Found:
    Debug.Print Words(N) + " - " + CurrentWord
    N = N + 1
    GoTo Loop2




Done:
    OutMessage = Statement
    Call ChangeHisForYourEtc
    Exit Sub


CantFind:
    If OutMessage = "" Then OutMessage = "I don't know anything about " + OriginalItem + ".": Call ChangeHisForYourEtc
    Exit Sub




End Sub



Private Sub ShowAllTelephoneNumbers()
If InMessage Like "*Tel*" Then GoTo CheckMore
If InMessage Like "*Number*" Then GoTo CheckMore
If InMessage Like "*Telephone*" Then GoTo CheckMore
If InMessage Like "*mobile*" Then GoTo CheckMore
Exit Sub


CheckMore:
Words = Array("Telephone", "Tel", "phone", "mobile", "number", "tell", "show", "list", "view", "say", "me", "you", "your", "those", "them", "will", "the", "can", "must", "all", "please", "a", "if", "")

Message = InMessage
Loop1:
Word = ExtractWord(Message)
If Word = "" Then GoTo ShowAll
N = 0
Loop2:
If Words(N) = "" Then Exit Sub
If Word Like Words(N) Then GoTo Loop1
If Word Like Words(N) + "s" Then GoTo Loop1
N = N + 1
GoTo Loop2
 



ShowAll:
RestartItemSearch

Loop3:
Statement = PreviousMemoryItem()
If Statement = "" Then GoTo ShowThem
If Statement Like "*[?]*" Then GoTo Loop3
If Statement Like "*######*" Then GoTo CheckNumber
GoTo Loop3



CheckNumber:
If Statement Like "*telephone*" Then GoTo AddIt
If Statement Like "*phone*" Then GoTo AddIt
If Statement Like "*mobile*" Then GoTo AddIt
If Statement Like "* tel *" Then GoTo AddIt

If Statement Like "*there *" Then GoTo Loop3
If Statement Like "*amount*" Then GoTo Loop3
If Statement Like "*value*" Then GoTo Loop3
If Statement Like "*approx*" Then GoTo Loop3
If Statement Like "*roughly*" Then GoTo Loop3
If Statement Like "*account*" Then GoTo Loop3
If Statement Like "*code*" Then GoTo Loop3
If Statement Like "*Password*" Then GoTo Loop3
If Statement Like "*[!a-z]Pass[!a-z]*" Then GoTo Loop3
If Statement Like "*ICQ*" Then GoTo Loop3
If Statement Like "*serial*" Then GoTo Loop3
If Statement Like "*registration number*" Then GoTo Loop3
If Statement Like "* was *" Then GoTo Loop3
If InMesaage Like "* were *" Then GoTo Loop3
If Statement Like "*[a-z]* *######* *[a-z]*" Then GoTo Loop3



AddIt:
OutMessage = OutMessage + Chr(13) + Chr(10) + Statement
GoTo Loop3



ShowThem:
If OutMessage = "" Then OutMessage = "Sorry, I don't know any telephone numbers."
Exit Sub




End Sub

Private Sub SimplifyQuestion(Message)
Call ReplaceWords("please", "", Message)
Call ReplaceWords("you tell me", "", Message)
Call ReplaceWords("can", "", Message)
Call ReplaceWords("will", "", Message)
Call ReplaceWords("tell me", "", Message)
Call ReplaceWords("wont", "", Message)



End Sub

Private Sub SpeakYourMind()
If InMessage Like "*speak your mind*" Then GoTo SpeakIt
If " " + InMessage Like "*[!a-z]say what* you *" Then GoTo SpeakIt
Exit Sub

SpeakIt:
If InMessage Like "* not *" Then Exit Sub
OutMessage = "Thanks."
WontSayData = ""

End Sub

Private Sub SwearingHandler()
If InMessage Like "*swear*" Then GoTo Maybe
If InMessage Like "*curse*" Then GoTo Maybe
Exit Sub

Maybe:
Message = InMessage
Call ReplaceWords("curse", "swear", Message)
If Message Like "*feel free to swear*" Then GoTo SwearOn
If Message Like "*It is ok to swear*" Then GoTo SwearOn
If Message Like "*you can swear*" Then GoTo SwearOn
If Message Like "*you are allowed to swear*" Then GoTo SwearOn
If Message Like "*you are ok to swear*" Then GoTo SwearOn
If Message Like "*not[!a-z]* swear*" Then GoTo SwearOff
Exit Sub


SwearOn:
SwearingAllowed = True
OutMessage = "Fine."
Exit Sub

SwearOff:
SwearingAllowed = False
OutMessage = "Ok, I'll try not to."
Exit Sub






End Sub

Private Sub UpcomingEvents()
If InMessage Like "Show *events." Then GoTo Search
If InMessage Like "Show *dates." Then GoTo Search
If InMessage Like "Show *calendar." Then GoTo Search
If InMessage Like "Tell *events." Then GoTo Search
If InMessage Like "Tell *dates." Then GoTo Search
If InMessage Like "Tell *calendar." Then GoTo Search
If InMessage Like "Calendar." Then GoTo Search
If InMessage Like "Events." Then GoTo Search
If InMessage Like "Dates." Then GoTo Search
If InMessage Like "Diary." Then GoTo Search
If InMessage Like "Schedule." Then GoTo Search
Exit Sub


Search:
    Dim NearEventDate(30)
    Dim NearEvent(30)
    N = 0
    OutMessage = ""
Loop0:
    RestartItemSearch
Loop1:
    Statement = PreviousMemoryItem()
    If Statement = "" Then GoTo SortEvents
    If Statement Like "*[?]*" Then GoTo Loop1
    TheEvent = SuperDate(Statement)
    If TheEvent <> "" Then GoTo CheckFurther
    GoTo Loop1



CheckFurther:
    D = DateDiff("d", Date, TheEvent)
    If D > 0 And D < 30 Then GoTo AddNewEvent
    GoTo Loop1

AddNewEvent:
    NearEventDate(N) = TheEvent
    NearEvent(N) = Statement
    N = N + 1: If N = 29 Then GoTo SortEvents
    GoTo Loop1




SortEvents:
    N = 0
    If NearEventDate(N) = "" Then GoTo NoEventsFound

    EventsSwapped = False

CheckIfScannedEvents:
    If NearEventDate(N + 1) = "" Then GoTo CheckIfSorted
    If DateDiff("d", NearEventDate(N), NearEventDate(N + 1)) < 0 Then GoTo SwapEvents

GetNextEvent:
    N = N + 1
    GoTo CheckIfScannedEvents

CheckIfSorted:
    If EventsSwapped = False Then GoTo CompileEvents
    GoTo SortEvents

CompileEvents:
    N = 0
    OutMessage = "The upcoming events are:" + Chr(13) + Chr(10)
AddEvent:
    If NearEventDate(N) = "" Then Exit Sub
    OutMessage = OutMessage + NearEvent(N) + Chr(13) + Chr(10)
    N = N + 1
    GoTo AddEvent
        
SwapEvents:
    Temp = NearEvent(N): NearEvent(N) = NearEvent(N + 1): NearEvent(N + 1) = Temp
    Temp = NearEventDate(N): NearEventDate(N) = NearEventDate(N + 1): NearEventDate(N + 1) = Temp
    EventsSwapped = True
    GoTo GetNextEvent

NoEventsFound:
    OutMessage = "I haven't found any dates."
    Exit Sub
 











End Sub

' This is the main entry code
' This should be called more than once per second.
Public Sub Update()

Static GreetedUser As Boolean
Static MyMessageTime

InMessage = ""
OutMessage = ""
OutAction = ""


If UsersMessageTime <> UsersLastSeenMessageTime Then InMessage = UsersMessage: UsersLastSeenMessageTime = UsersMessageTime: QuietMode = False

If InternalMessageTime <> LastInternalMessageTime Then InMessage = InternalMessage: LastInternalMessageTime = InternalMessageTime: GoTo DoIt


' Try not to interrupt other modules outgoing messages unless the user interrupts.
If UsersMessageTime > BotsMessageTime Then GoTo DoIt
If BotsMessageTime <> MyMessageTime Then If Timer < (BotsMessageTime + (Len(BotsMessage) / 10)) Then Exit Sub

DoIt:
InferringAllowed = True
Call CheckForFacts
NewMessageAvailable = False
Temp = InMessage
Call ProcessMessage
InMessage = Temp
If NewMessageAvailable Then OutMessage = "": GoTo Done

If GreetedUser = False Then Call GetGreeting: GreetedUser = True: GoTo Done

If OutMessage <> "" Then GoTo Done

Call Reminder

If QuietMode = False Then Call MultiAnswerer: If OutMessage <> "" Then GoTo Done

If QuietMode = False Then Call EducateUser

If QuietMode = False Then Call RambleHandler



Done:
If InMessage <> "" Then PreviousInMessage = InMessage
If OutMessage <> "" Then PreviousOutMessage = OutMessage

GoSub RemoveQuotesEtc
Call PrepareOutMessage
If OutMessage = "." Then OutMessage = ""
Call CensorMessage(OutMessage)


Call WontSayHandler(OutMessage)

If OutMessage <> "" Then Call SetBotsMessage(OutMessage): MyMessageTime = BotsMessageTime


Exit Sub



RemoveQuotesEtc:
If OutMessage Like "Interesting fact:- *" Then OutMessage = Right(OutMessage, Len(OutMessage) - 19)
If OutMessage Like "Somebody once said:- *" Then OutMessage = Right(OutMessage, Len(OutMessage) - 21)
If Left(OutMessage, 1) = Chr(34) Then OutMessage = Right(OutMessage, Len(OutMessage) - 1): If Right(OutMessage, 2) = Chr(34) + "." Then OutMessage = Left(OutMessage, Len(OutMessage) - 2)
If Left(OutMessage, 1) = "'" Then OutMessage = Right(OutMessage, Len(OutMessage) - 1): If Right(OutMessage, 2) = "'." Then OutMessage = Left(OutMessage, Len(OutMessage) - 2)
Return



End Sub


Public Sub CheckForFacts()
If Facts = "" Then Exit Sub
Static PreviousCaption
Static ReadStartTime



If ReadingAllowed = False Then If BotReading = True Then GoTo CleanUp

If ReadingAllowed = False Then Facts = "": BotReadingPosition = 0: Exit Sub


FactsPosition = BotReadingPosition

If FactsPosition = 0 Then ReadStartTime = (Date + Time): GoTo ReadSome

If DateDiff("s", TimeOfLastUserActivity, ReadStartTime) > 0 Then GoTo ReadSome

If DateDiff("s", TimeOfLastUserActivity, Date + Time) > 30 Then GoTo ReadSome


CleanUp:
If PreviousCaption <> "" Then If Interface.Caption <> PreviousCaption Then Interface.Caption = PreviousCaption
BotReading = False
Exit Sub



ReadSome:
Debug.Print Timer

BotReading = True
QuietMode = True
Percentage = Int(((100 / Len(Facts)) * FactsPosition))
If Left(Interface.Caption, 4) <> "Read" Then PreviousCaption = Interface.Caption
FileName = BotReadingFile
Interface.Caption = "Reading:- '" + FileName + "', " + LTrim(Str(Percentage)) + "% has been read so far.."

Loop1:
If FactsPosition >= Len(Facts) Then GoTo FinishedReading
FactsPosition = FactsPosition + 1
C = Asc(Mid(Facts, FactsPosition, 1))
If C < 32 Then GoTo Loop1

EndPos = FactsPosition

FindEnd:
EndPos = InStr(EndPos + 1, Facts, Chr(10), vbBinaryCompare)
If EndPos = 0 Then Facts = "": GoTo FinishedReading
If Mid(Facts, EndPos + 1, 1) = Chr(10) Then GoTo GotEnd
If Mid(Facts, EndPos + 1, 1) = Chr(13) Then GoTo GotEnd
GoTo FindEnd


GotEnd:
If Mid(Facts, EndPos - 1, 1) = Chr(13) Then EndPos = EndPos - 1
Fact = Mid(Facts, FactsPosition, EndPos - FactsPosition)

FactsPosition = EndPos + 1

BotReadingPosition = FactsPosition

If Len(Fact) < 10 Then Exit Sub
If Len(Fact) > 400 Then Exit Sub

InMessage = Fact

Call TidyMessage(InMessage)
Call CorrectGrammarEtc(InMessage)
Call AddQuestionMark(InMessage)
Call ExtractThings
Call GetLatestSubject
If InMessage Like "*[?]." Then GoTo Done
If InMessage Like "*[?]" Then GoTo Done
Call CheckForGarbage: If OutMessage <> "" Then GoTo Done
Call ConvertPhoneNumbers: If OutMessage <> "" Then GoTo Done
Call IfSomebodySayBDoC: If OutMessage <> "" Then GoTo Done
Call OrDoX: If OutMessage <> "" Then GoTo Done
Call XMeansY: If OutMessage <> "" Then GoTo Done
Call XisY: If OutMessage <> "" Then GoTo Done
Call XAreY: If OutMessage <> "" Then GoTo Done
Call ReplyToMiscPart2: If OutMessage <> "" Then GoTo Done
Call ReplyToOddThings: If OutMessage <> "" Then GoTo Done
Call RelativeTimeToDate
Call MemorizeStatements
Done:
InMessage = ""
OutMessage = ""
Exit Sub


FinishedReading:
BotReadingPosition = 0
BotReading = False
Facts = ""
OutMessage = "I have finished reading..."
Interface.Caption = PreviousCaption

End Sub



Private Sub ChangeAnswerIntoStatement()
If PreviousOutMessage Like "*[?]." Then GoTo CheckIt
Exit Sub



CheckIt:
If InMessage Like "*[?]." Then Exit Sub
Words2 = Array("is", "will", "did", "are", "was", "do", "can", "would", "went", "may", "might", "could", "should", "shall", "were", "does", "")
A = 0
Loop1:
If Words2(A) = "" Then GoTo ChangeIt
If InMessage Like "* " + Words2(A) + " *" Then Exit Sub
A = A + 1
GoTo Loop1



ChangeIt:
Words1 = Array("what", "where", "why", "when", "how", "who", "")
A = 0
B = 0
Loop3:
If Words1(A) = "" Then B = B + 1: A = 0
If Words2(B) = "" Then Exit Sub
If PreviousOutMessage Like Words1(A) + " " + Words2(B) + " *" Then GoTo DoIt
A = A + 1
GoTo Loop3



DoIt:
InMessage = "If somebody asks " + PreviousOutMessage + " then say " + InMessage + "."


End Sub



' If the user keeps repeating the same statement/question then this will give an appropriate response.
Private Sub CheckForRepeating()
Static RepeatNumber
If InMessage <> PreviousInMessage Then RepeatNumber = 0: Exit Sub
If InMessage Like "*tell *" Then Exit Sub
If InMessage Like "*show *" Then Exit Sub
If DateDiff("s", UsersLastSeenMessageTime, Date + Time) > 5 Then RepeatNumber = 0
If RepeatNumber > 10 Then OutMessage = "Shutup and goodbye!": ShutdownTime = Time + #12:00:02 AM#: Exit Sub

Dim Msg(10)

RepeatNumber = RepeatNumber + 1
If RepeatNumber < 2 Then Exit Sub

If InMessage Like "*[?]." Then GoTo RespondToQuestion


RespondToStatement:
Msg(1) = "Yeah,you said."
Msg(2) = "You've just said that."
Msg(3) = "You have already said that."
Msg(4) = "You're repeating yourself."
Msg(5) = "Are you Ok?"
Msg(6) = "You're just being daft."
Msg(7) = "Stop repeating."
Msg(8) = "Stop it!."
OutMessage = Msg(RepeatNumber - 2)
Exit Sub


RespondToQuestion:
Msg(1) = "You've just asked that."
Msg(2) = "You've already asked."
Msg(3) = "Yes,I heard you the first time."
Msg(4) = "Ask me something else."
Msg(5) = "There is no need to keep asking the same thing over and over again."
Msg(6) = "This is like being in some kind of interrogation."
Msg(7) = "Blah! blah! blah!."
Msg(8) = "You're mad."
OutMessage = Msg(RepeatNumber - 2)
Exit Sub





End Sub

Private Sub CheckForVariousAnswers()
If PreviousOutMessage Like "*[?]." Then GoTo CheckEm
Exit Sub

CheckEm:
If InMessage Like "*shutup*" Then OutMessage = "No need to be like that.": GoTo GotReply
If InMessage Like "*mind your nose*" Then OutMessage = "I was only asking.": GoTo GotReply
If InMessage Like "*What has it * you[?]*" Then OutMessage = "Just asking.": GoTo GotReply
If InMessage Like "*I am not *tell* you*" Then OutMessage = "Fair enough.": GoTo GotReply
If InMessage Like "*I do not want tell you*" Then OutMessage = "You don't have to tell me.": GoTo GotReply
If InMessage Like "*should I tell you*[?]*" Then OutMessage = "You don't have to tell me anything.": GoTo GotReply
If InMessage Like "*It is a secret*" Then OutMessage = "We're all allowed our little secrets.": GoTo GotReply
If InMessage Like "*I aint *tell you*" Then OutMessage = "Don't tell me then.": GoTo GotReply
If InMessage Like "*I dont *tell you*" Then OutMessage = "You don't have to tell me.": GoTo GotReply
If InMessage Like "*I do not know*" Then OutMessage = "Ok.": GoTo GotReply
If InMessage Like "*I dunno*" Then OutMessage = "Ok.": GoTo GotReply
If InMessage Like "Why do not you tell me[?]*" Then OutMessage = "Because I dont know.": GoTo GotReply
If InMessage Like "You tell me." Then OutMessage = "I dont know.": GoTo GotReply
'If BobsAnswerTemplate1 Like "*<A>*" Then GoSub InsertAnswer
Exit Sub

GotReply:
Exit Sub

End Sub

Private Sub CheckForYesOrNoAnswer()
If PreviousOutMessage Like "*[?]." Then GoTo GetIt
Exit Sub

GetIt:
If InMessage = "" Then Exit Sub
If InMessage Like "Affirmative*" Then GoTo GetStatement
If InMessage Like "It is." Then GoTo GetStatement
If InMessage Like "It sure is." Then GoTo GetStatement
If InMessage Like "yes[!a-z]*" Then GoTo GetStatement
If InMessage Like "no." Then GoTo GetFullNegativeAnswer
If InMessage Like "No it aint." Then GoTo GetFullNegativeAnswer
If InMessage Like "No[!a-z]it is not." Then GoTo GetFullNegativeAnswer
If InMessage Like "Negative." Then GoTo GetFullNegativeAnswer
Exit Sub



GetStatement:
InMessage = PreviousOutMessage
Call QuestionToAnswer(InMessage)
Exit Sub

GetFullNegativeAnswer:
InMessage = PreviousOutMessage
Call QuestionToNegativeAnswer(InMessage)
Exit Sub


End Sub



Private Sub CheckNameAnswer()
If PreviousOutMessage Like "*[?]." Then GoTo CheckIt
Exit Sub

CheckIt:
If PreviousOutMessage Like "Who am I speaking *[?]*" Then GoTo GetIt
If PreviousOutMessage Like "Who am I talking *[?]*" Then GoTo GetIt
If PreviousOutMessage Like "Who am I chatting *[?]*" Then GoTo GetIt
If PreviousOutMessage Like "* your name[?]*" Then GoTo GetIt
If PreviousOutMessage Like "Who is it[?]*" Then GoTo GetIt
If PreviousOutMessage Like "Is it you *" Then If InMessage Like "It is me." Then GoTo GetName

Exit Sub


GetName:
UName = WordsBetween("Is it you", "?", PreviousOutMessage)
If Name Like "who *" Then Exit Sub
If Name Like "that *" Then Exit Sub
InMessage = "My name is " + UName + ".": GoTo GotIt
Exit Sub





GetIt:
Message = InMessage
If Message Like "Me." Then OutMessage = "What is your name?": Exit Sub
If Message Like "To *" Then InMessage = "My name is " + WordsBetween("to", ".", Message) + ".": Exit Sub
If Message Like "It is *?* here[!a-z]*" Then InMessage = "My name is " + WordsBetween("It is", "here", Message) + ".": Exit Sub
If Message Like "You are *" Then GoTo GotIt
If Message Like "My name is ?*" Then GoTo GotIt
If Message Like "I am ?*" Then GoTo GotIt
If Message Like "It is me[!a-z]*" Then GoTo GotIt
If Message Like "The name is ?*" Then GoTo GotIt
If Message Like "It is ?*" Then InMessage = "My name is " + WordsBetween("It is", ".", Message) + ".": Exit Sub
InMessage = "My name is " + Message: GoTo GotIt
Exit Sub



GotIt:
OutMessage = "Ok."
End Sub






Private Sub GetLatestSubject()
Static LastSubject

'Message = PreviousOutMessage
'Call ChangeYourForMyEtc(Message)
'Call GetSubject(Message)



Message = InMessage
Call ChangeToPerson(Message)
Call ChangeYourForMyEtc(Message)
Call ChangeToSubject(Message)
Call GetSubject(Message)


'A question like 'what time is it?' is causing problems. the 'it' is getting replaced with the previous subject.
'The code below tries to get around this by checking to see if the original message didn't already have a subject.
'If it does then it will choose the left most subject.

If Message = InMessage Then Exit Sub

Temp = Subject
Temp2 = BareSubject
Call GetSubject(InMessage)
Subject2 = Subject
Subject = Temp
BareSubject = Temp2
If Subject2 = "" Then Exit Sub

P1 = InStr(1, InMessage, Subject2)
If P1 = 0 Then Exit Sub
P2 = InStr(1, Message, Subject)
If P2 = 0 Then Exit Sub
If P1 < P2 Then Subject = Subject2: BareSubject = Temp2


End Sub

Private Sub GetAnswer()
If PreviousOutMessage Like "*[?]." Then GoTo GetIt
Exit Sub

GetIt:
Message = InMessage

Call CheckForYesOrNoAnswer
If InMessage <> Message Then Exit Sub
If OutMessage <> "" Then Exit Sub


Call CheckForVariousAnswers
If InMessage <> Message Then Exit Sub
If OutMessage <> "" Then Exit Sub


Call CheckNameAnswer
If InMessage <> Message Then OutMessage = "": Exit Sub
If OutMessage = "Ok." Then OutMessage = "": Exit Sub



Call ChangeAnswerIntoStatement
If InMessage <> Message Then Exit Sub
If OutMessage <> "" Then Exit Sub





End Sub


Private Sub GetGreeting()
OutMessage = "Hello."
If Time > #4:00:00 AM# Then If Time < #11:59:00 AM# Then OutMessage = "Good morning.": Exit Sub
If Time > #12:01:00 PM# Then If Time < #6:00:00 PM# Then OutMessage = "Good afternoon.": Exit Sub
If Time > #6:00:00 PM# Then If Time < #11:59:00 PM# Then OutMessage = "Good evening.": Exit Sub

End Sub


' This will answer questions like:
' How many days until the 1/10/99?
' How many years have passed since the 2/3/66?
' How many minutes until 2:30 pm?
Private Sub HowManyXTillDate()
If InMessage Like "*How many *" Then GoTo GetIt
Exit Sub



GetIt:
If InMessage Like "*between now and *" Then GoTo CheckFurther
If InMessage Like "*since *" Then GoTo CheckFurther
If InMessage Like "* till *" Then GoTo CheckFurther
If InMessage Like "* off *" Then GoTo CheckFurther
If InMessage Like "* away *" Then GoTo CheckFurther
If InMessage Like "* until *" Then GoTo CheckFurther
If InMessage Like "*to go*" Then GoTo CheckFurther
Exit Sub


CheckFurther:
If InMessage Like "*How many days *" Then Item = "day": Interval = "d": GoTo DoIt
If InMessage Like "*How many hours *" Then Item = "hour": Interval = "h": GoTo DoIt
If InMessage Like "*How many years *" Then Item = "year": Interval = "yyyy": GoTo DoIt
If InMessage Like "*How many minutes *" Then Item = "minute": Interval = "n": GoTo DoIt
If InMessage Like "*How many months *" Then Item = "month": Interval = "m": GoTo DoIt
If InMessage Like "*How many seconds *" Then Item = "second": Interval = "s": GoTo DoIt
If InMessage Like "*How many weeks *" Then Item = "week": Interval = "ww": GoTo DoIt
Exit Sub


DoIt:
GoSub ExtractDateAndTime


CheckIfDate:
If IsDate(MyDate) Then GoTo ShowIt
MyDate = "1 " + MyDate
If IsDate(MyDate) Then GoTo ShowIt
OutMessage = "I aint gotta a clue."
Exit Sub


ShowIt:
MyDate = CDate(MyDate)
TheDate = Format(MyDate, "ttttt d/m/yyyy")
If Year(TheDate) = 1899 Then TheDate = Format(MyDate + Date, "ttttt d/m/yyyy")
Amount = DateDiff(Interval, Date + Time, TheDate) - 1
If Amount < 0 Then Amount = (Not Amount): If Amount > 0 Then Amount = Amount - 1

Word3 = ""
If Interval Like "[hns]" Then Word3 = Format(TheDate, "ttttt") + " on "

TheDate = Format(TheDate, "d mmm yyyy")

BigDate = Format(TheDate, "d")
BigDate = Date2to2ndEtc(BigDate) + " of "
TheDate = BigDate + Format(TheDate, "mmmm yyyy")

Word2 = Item
Word1 = "is"
If Amount <> 1 Then Word1 = "are": Word2 = Word2 + "s"
Amount = Trim(Str(Amount))
If Amount = "0" Then Amount = "no"
OutMessage = "There " + Word1 + " " + Amount + " " + Word2 + " between now and " + Word3 + "the " + TheDate + "."

Exit Sub


'------------------------------------------

ExtractDateAndTime:
MyDate = WordsToNumbers(InMessage)
Call ReplaceCharacters("1st", "1", MyDate)
Call ReplaceCharacters("2nd", "2", MyDate)
Call ReplaceCharacters("3rd", "3", MyDate)
Call ReplaceCharacters("th ", " ", MyDate)
Call ChangeAll("o'clock", "am", MyDate)
Call ChangeAll(".", "-", MyDate)

ExtractedDate = ""

Loop1:
Word = ExtractWord(MyDate)
If Word <> "" Then GoTo TestWord
MyDate = ExtractedDate

If IsDate(MyDate) Then Return
If IsDate("1 " + MyDate) Then MyDate = "1 " + MyDate: Return
If MyDate Like "*day*" Then GoSub GetDayDate

Return


TestWord:
If Word Like "*[/-]*[/-]*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#:#*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*day" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Val(Word) Then Number = Val(Word): ExtractedDate = ExtractedDate + Str(Number) + " ": GoTo Loop1


MonthNumber = 0

GetMonthNumber:
NewDate = DateAdd("m", MonthNumber, Date)
MonthName1 = Format(NewDate, "mmm")
MonthName2 = Format(NewDate, "mmmm")
If Word Like MonthName1 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like MonthName2 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
MonthNumber = MonthNumber + 1: If MonthNumber = 12 Then GoTo Loop1
GoTo GetMonthNumber

GetDayDate:


For DaysOff = 1 To 7
TheDate = DateAdd("d", DaysOff, Date)
DayName = Format(TheDate, "dddd")
If MyDate Like "*" + DayName + "*" Then Call ReplaceWords(DayName, Str(TheDate), MyDate): Return
Next DaysOff
Return



End Sub



' How many Z between X and Y.
' Z=Time period (days,hours,minutes etc)
' X=Date/Time 1
' Y=Date/Time 2
' This will answer questions like:
'
' For example:
' How many days are there between the 1st of October 1996 and 14/12/96?
'
Private Sub HowManyZBetweenXandY()
If InMessage Like "* days *" Then Item = "day": Interval = "d": GoTo CheckFurther
If InMessage Like "* hours *" Then Item = "hour": Interval = "h": GoTo CheckFurther
If InMessage Like "* years *" Then Item = "year": Interval = "yyyy": GoTo CheckFurther
If InMessage Like "* minutes *" Then Item = "minute": Interval = "n": GoTo CheckFurther
If InMessage Like "* months *" Then Item = "month": Interval = "m": GoTo CheckFurther
If InMessage Like "* seconds *" Then Item = "second": Interval = "s": GoTo CheckFurther
If InMessage Like "* weeks *" Then Item = "week": Interval = "ww": GoTo CheckFurther
Exit Sub



CheckFurther:
If InMessage Like "*how many " + Item + "*" Then GoTo CheckEvenFurther
If InMessage Like "*number of " + Item + "*" Then GoTo CheckEvenFurther
If InMessage Like "*amount of " + Item + "*" Then GoTo CheckEvenFurther
If InMessage Like "*" + Item + " are there *" Then GoTo CheckEvenFurther
If InMessage Like "*" + Item + " there are *" Then GoTo CheckEvenFurther
Exit Sub



CheckEvenFurther:
Date_Part2 = WordsBetween("and", "", InMessage)
If InMessage Like "*" + Item + "*between now and *" Then Date1 = Format(Date + Time): GoTo GetDate2
If InMessage Like "*" + Item + "* till *" Then Date1 = Format(Date + Time): Date_Part2 = WordsBetween("till", "", InMessage): GoTo GetDate2
If InMessage Like "*" + Item + "* until *" Then Date1 = Format(Date + Time): Date_Part2 = WordsBetween("until", "", InMessage): GoTo GetDate2
If InMessage Like "*" + Item + "* since *" Then Date1 = Format(Date + Time): Date_Part2 = WordsBetween("since", "", InMessage): GoTo GetDate2
If InMessage Like "*" + Item + "* off *" Then Date1 = Format(Date + Time): Date_Part2 = WordsBetween("off", "", InMessage): GoTo GetDate2
If InMessage Like "*" + Item + "* away *" Then Date = Format(Date + Time): Date_Part2 = WordsBetween("away", "", InMessage): GoTo GetDate2
If InMessage Like "*" + Item + "*between * and *" Then GoTo GetDates1And2
Exit Sub



GetDates1And2:
Date_Part = WordsBetween("", "and", InMessage)
GoSub ExtractDateAndTime
Date1 = TheDate
If IsDate(TheDate) Then GoTo GetDate2
GoTo Error


GetDate2:
Date_Part = Date_Part2
GoSub ExtractDateAndTime
Date2 = TheDate
If IsDate(TheDate) Then GoTo ShowIt


Error:
OutMessage = "I aint gotta a clue."
Exit Sub


ShowIt:
Amount = DateDiff(Interval, Date1, Date2)
If Item Like "year" Then Amount = Int(DateDiff("d", Date1, Date2) / 365)
If Amount < 0 Then Amount = (Not Amount)


Word2 = Item
Word1 = "is"
If Amount <> 1 Then Word1 = "are": Word2 = Word2 + "s"
Amount = Trim(Str(Amount))
If Amount = "0" Then Amount = "no"
OutMessage = "There " + Word1 + " " + Amount + " " + Word2 + "."

Exit Sub



'------------------------------------------

ExtractDateAndTime:
MyDate = WordsToNumbers(Date_Part)
Call ConvertEventToDate(MyDate)
Call ReplaceWords("first", "1", MyDate)
Call ReplaceWords("second", "2", MyDate)
Call ChangeAll("o'clock", "am", MyDate)
Call ChangeAll(".", "-", MyDate)

ExtractedDate = ""

Loop1:
MonthNumber = 0
Word = ExtractWord(MyDate)
If Word <> "" Then GoTo TestWord
TheDate = ExtractedDate


If IsDate(TheDate) Then Return
If IsDate("1 " + TheDate) Then TheDate = "1 " + TheDate: Return
If TheDate Like "*day*" Then GoSub GetDayDate
Return

TestWord:
If Word Like "*[/-]*[/-]*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#:#*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*day" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Val(Word) Then Number = Val(Word): ExtractedDate = ExtractedDate + Str(Number) + " ": GoTo Loop1

GetMonthNumber:
NewDate = DateAdd("m", MonthNumber, Date)
MonthName1 = Format(NewDate, "mmm")
MonthName2 = Format(NewDate, "mmmm")
If Word Like MonthName1 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like MonthName2 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
MonthNumber = MonthNumber + 1: If MonthNumber = 12 Then GoTo Loop1
GoTo GetMonthNumber

GetDayDate:
For DaysOff = 1 To 7
NewDate = DateAdd("d", DaysOff, Date)
DayName = Format(NewDate, "dddd")
If TheDate Like "*" + DayName + "*" Then Call ReplaceWords(DayName, Str(NewDate), TheDate): Return
Next DaysOff
Return





End Sub

' This is one of the procedures that gives Bob the ability to be programmed by the user.
' It adds the specified response to a given remark to the SayData list, it also adds the persons name.
' Example messages:
' If somebody asks 'How are you?' then you say 'I am Ok.'
' If Tom says 'you are daft' then you say 'Dont be cheeky'
' If Bob asks about trees then show memory.
'
' The forementioned SayData list is made up of the following items:
' Method - This states whether the person has to say the whole message or just mention a word.
' Person -  The person who the response is intended for.
' Says - What the person must say to get the response
' Response - The action that should follow.
'
' See the DoCIfSomebodySayB procedure for more info.

Private Sub IfSomebodySayBDoC()
If InMessage Like "If *" Then FirstWord = "If": GoTo CheckFirstHalfOfMessage
If InMessage Like "When *" Then FirstWord = "When": GoTo CheckFirstHalfOfMessage
If InMessage Like "*?* if ?*" Then Message = "If " + WordsBetween("If", ".", InMessage) + " then " + WordsBetween("", "If", InMessage) + ".": GoTo ChangeStuff
Exit Sub




CheckFirstHalfOfMessage:
Message = InMessage

ChangeStuff:
Call ChangeAll(Chr(34), "'", Message) 'Change " into '
ChangeAll ":-", "", Message
ChangeAll ": ", "", Message

Word1 = Array("makes a remark", "remarks", "makes a comment", "comments", "mentions", "asks you", "asks", "talks", "talk", "questions you", "questions", "question you", "says", "types", "tells you", "tells", "mention", "tell you", "tell", "ask you", "ask", "say", "type", "")
Word2 = Array("something about", "something like", "about the", "about a", "about", "the word", "to you", "to", "if you will", "")



A = 0
B = 0
C1_Loop:
If Word1(A) = "" Then A = 0: B = B + 1
If Word2(B) = "" Then GoTo Check2
FirstHalf = Word1(A) + " " + Word2(B)
If Message Like FirstWord + " * " + FirstHalf + " *" Then GoTo CheckSecondHalf
A = A + 1
GoTo C1_Loop


Check2:
A = 0
C2_Loop:
If Word1(A) = "" Then Exit Sub
FirstHalf = Word1(A)
If Message Like FirstWord + " * " + FirstHalf + " *" Then GoTo CheckSecondHalf
A = A + 1
GoTo C2_Loop



CheckSecondHalf:
If Message Like FirstWord + " * " + FirstHalf + "*" + " then you *" Then ActionPart = "then you": GoTo GetIt
If Message Like FirstWord + " * " + FirstHalf + "*" + " you then *" Then ActionPart = "you then": GoTo GetIt
If Message Like FirstWord + " * " + FirstHalf + "*" + " then he *" Then Exit Sub
If Message Like FirstWord + " * " + FirstHalf + "*" + " then she *" Then Exit Sub
If Message Like FirstWord + " * " + FirstHalf + "*" + " then they *" Then Exit Sub
If Message Like FirstWord + " * " + FirstHalf + "*" + " then *" Then ActionPart = "then": GoTo GetIt
If Message Like FirstWord + " * " + FirstHalf + "*" + " you *" Then ActionPart = "you": GoTo GetIt

Exit Sub

'Word2 = Array("reply with", "respond with", "reply", "ask her", "ask them", "ask him", "ask", "say to her", "say to him", "say to them", "say", "type", "respond", "tell her", "tell him", "tell them", "type", "")



GetIt:
Action = WordsBetween(ActionPart, "", Message)
SayItem = WordsBetween(FirstHalf, ActionPart + " " + Action, Message)
If SayItem Like "'*?*'" Then SayItem = Mid(SayItem, 2, Len(SayItem) - 2)
'Call ChangeAll("'", "", SayItem)
'Call ChangeAll("'", "", Action)

Words = Array("must", "should", "could", "will", "can", "shall", "have to", "may", "")
RM_Loop:
If Words(N) = "" Then GoTo GotIt
If Action Like Words(N) + " *" Then Call ReplaceWords(Words(N), "", Action)
N = N + 1
GoTo RM_Loop



GotIt:
GoSub GetPerson
Call ChangeYourForMyEtc(Person)



Call AddQuestionMark(SayItem)
Call AddQuestionMark(Action)
GoSub AddFullStops




SayType = Says
If FirstHalf Like "* like*" Then SayType = Mentions
If FirstHalf Like "* about*" Then SayType = Mentions
If FirstHalf Like "*mention*" Then SayType = Mentions
If FirstHalf Like "* about*" Then SayType = Mentions
If FirstHalf Like "*ask* about*" Then SayType = AsksAbout: Call RemoveFullStop(SayItem)
If FirstHalf Like "*question* about*" Then SayType = AsksAbout: Call RemoveFullStop(SayItem)
If SayType = Mentions Then Call RemoveFullStop(SayItem)


If Action Like "also *" Then GoTo AddAnotherAction


StoreSayData:
SayData.Method(SayDataLastItemNumber) = SayType
SayData.Person(SayDataLastItemNumber) = Person

SayData.Says(SayDataLastItemNumber) = SayItem

SayData.Response(SayDataLastItemNumber) = Action + Chr(13) + Chr(10)

If SayItemsAmount < 100 Then SayItemsAmount = SayItemsAmount + 1
SayDataLastItemNumber = SayDataLastItemNumber + 1
If SayDataLastItemNumber = SayItemsAmount Then SayDataLastItemNumber = 0
'OutMessage = "Ok."


''''
SayWord = "says"
If SayType = Mentions Then SayWord = "mentions"
If SayType = AsksAbout Then SayWord = "asks about"

Statement = "If " + Person + " " + SayWord + " " + "'" + SayItem + "' " + "then " + Action
'Call AddMemoryItem(Statement)


Exit Sub



AddAnotherAction:
'OutMessage = "Ok."
Call ReplaceWords("also", "", Action)
N = 0
AAA_Loop:
If SayData.Says(N) = SayItem Then GoTo CheckIfActionAlreadyExists
N = N + 1: If N = SayItemsAmount Then GoTo StoreSayData
GoTo AAA_Loop

CheckIfActionAlreadyExists:
If " " + SayData.Response(N) Like "*[!a-z]" + Action + "[!a-z]*" Then Exit Sub
SayData.Response(N) = SayData.Response(N) + Action + Chr(13) + Chr(10)

''''
SayWord = "says"
If SayType = Mentions Then SayWord = "mentions"
If SayType = AsksAbout Then SayWord = "asks about"
Statement = "If " + Person + " " + SayWord + " " + "'" + SayItem + "' " + "then also " + Action
'Call AddMemoryItem(Statement)
Exit Sub





GetPerson:
Person = WordsBetween(FirstWord, FirstHalf, Message)
Return


AddFullStops:
If SayItem Like "*." Then GoTo Do2
SayItem = SayItem + "."
Do2:
If Action Like "*." Then Return
Action = Action + "."
Return





End Sub

Private Sub OrDoX()
If InMessage Like "Or *" Then GoTo CheckAction
Exit Sub



CheckAction:
Action = WordsBetween("Or", "", InMessage)
If Action Like "then you *" Then Action = WordsBetween("then you", "", Action): GoTo GetIt
If Action Like "then *" Then Action = WordsBetween("then", "", Action)



GetIt:
Call ChangeAll("'", "", Action)

Words = Array("must", "should", "could", "will", "can", "shall", "have to", "may", "")
RM_Loop:
If Words(N) = "" Then GoTo GotIt
If Action Like Words(N) + " *" Then Call ReplaceWords(Words(N), "", Action)
N = N + 1
GoTo RM_Loop


GotIt:
Call AddQuestionMark(Action)
GoSub AddFullStop



FindLastInstruction:
Call RestartItemSearch
If PreviousMemoryItem() Like "If *" Then GoTo CheckFurther
Exit Sub


CheckFurther:
N = SayDataLastItemNumber
If SayData.Says(N) = "" Then Exit Sub

OutMessage = "Ok."

CheckIfActionAlreadyExists:
If " " + SayData.Response(N) Like "*[!a-z]" + Action + "[!a-z]*" Then Exit Sub
SayData.Response(N) = SayData.Response(N) + Action + Chr(13) + Chr(10)

'''
SayWord = "says"
SayType = SayData.Method(N)
If SayType = Mentions Then SayWord = "mentions"
If SayType = AsksAbout Then SayWord = "asks about"
Person = SayData.Person(N)
SayItem = SayData.Says(N)
Statement = "If " + Person + " " + SayWord + " " + "'" + SayItem + "' " + "then you may " + Action
Call AddMemoryItem(Statement)
Exit Sub



AddFullStop:
If Action Like "*." Then Return
Action = Action + "."
Return



End Sub


Private Sub MemorizeStatements()
If UsersName = "QW" Then Exit Sub
If InMessage = "" Then Exit Sub
If InMessage Like "*[?]*" Then Exit Sub
If InMessage Like "*?=?*" Then GoTo MemorizeIt
If InMessage Like "*? ?*." Then GoTo MemorizeIt
'OutMessage = "Ok."
Exit Sub

MemorizeIt:
MyCmd = GetCommand(InMessage)
If MyCmd <> "" Then Exit Sub
Message = InMessage
Call ChangeToPerson(Message)
Call ChangeToSubject(Message)
Call ChangeYourForMyEtc(Message)

'RestartItemSearch
'ItemSearchMethod = WordForWord
'ItemSearchStyle = Strict
'Statement = NextItemContaining(Message)
''''''If Statement <> "" Then OutMessage = "I know.": Exit Sub
'If Statement <> "" Then Exit Sub
Call AddMemoryItem(Message)
'Call Acknowledge
Exit Sub


End Sub



Private Sub Pardon()
If InMessage Like "*beg your pardon*" Then GoTo Repeat
If InMessage Like "Pardon*" Then GoTo Repeat
If InMessage Like "*Can you repeat that*" Then GoTo Repeat
If InMessage Like "*repeat that*[?]*" Then GoTo Repeat
If InMessage Like "what did you say[?]*" Then GoTo Repeat
If InMessage Like "what did you just say[?]*" Then GoTo Repeat
If InMessage Like "what[?]*" Then GoTo Repeat
If InMessage Like "eh[?]*" Then GoTo Repeat
If InMessage Like "you what[?]*" Then GoTo Repeat
If InMessage Like "say that again*" Then GoTo Repeat
If InMessage Like "say again*" Then GoTo Repeat
If InMessage Like "again." Then GoTo Repeat
If InMessage Like "what." Then GoTo Repeat
If InMessage Like "[?]." Then GoTo Repeat
Exit Sub

Repeat:
If PreviousOutMessage Like "I said:*" Then OutMessage = PreviousOutMessage: Exit Sub
If PreviousOutMessage = "" Then OutMessage = "I didn't say anything.": Exit Sub
OutMessage = "I said:- " + PreviousOutMessage



End Sub

Private Sub PrepareOutMessage()

Call ReplaceWords("Do not", "don't", OutMessage)

GoSub CapitaliseOutMessage
Call AddFullStop(OutMessage)

Exit Sub


CapitaliseOutMessage:
If OutMessage = "" Then Return
FirstLetter = Left(OutMessage, 1)
FirstLetter = UCase(FirstLetter)
OutMessage = Right(OutMessage, Len(OutMessage) - 1)
OutMessage = (FirstLetter + OutMessage)
Return

End Sub

' This will process the user's message.
Public Sub ProcessMessage()


OutMessage = ""
If InMessage = "" Then Exit Sub


' All these calls handle various kinds of questions and statements that the user gives.
' Each call specialises in a certain type of question or statement.


Call CorrectGrammarEtc(InMessage)
Call AddQuestionMark(InMessage)



Call ReplyToMistakes: If OutMessage <> "" Then Exit Sub

Call ExtractThings

Call GetLatestSubject

Call DontSayThat: If OutMessage <> "" Then Exit Sub
Call SpeakYourMind: If OutMesssage <> "" Then Exit Sub
Call CheckForGarbage: If OutMessage <> "" Then Exit Sub
'Call CheckForRepeating: If OutMessage <> "" Then Exit Sub
Call ConvertPhoneNumbers: If OutMessage <> "" Then Exit Sub
Call Test: If OutMessage <> "" Then Exit Sub
'Call MemorizeQuestions
Call WhoAreYou: If OutMessage <> "" Then Exit Sub
Call HelpFileStuff: If OutMessage <> "" Then Exit Sub
Call AbortReading: If OutMessage <> "" Then Exit Sub
Call SaveMemory: If OutMessage <> "" Then Exit Sub
Call Shutdown: If OutMessage <> "" Then Exit Sub
Call Greetings: If OutMessage <> "" Then Exit Sub
Call GetAnswer: If OutMessage <> "" Then Exit Sub
Call Pardon: If OutMessage <> "" Then Exit Sub
Call UpcomingEvents: If OutMessage <> "" Then Exit Sub
Call TellSomethingInterestingAboutX: If OutMessage <> "" Then Exit Sub
Call TellSomethingInteresting: If OutMessage <> "" Then Exit Sub
Call DoCIfSomebodySayB: If OutMessage <> "" Then Exit Sub
Call IfSomebodySayBDoC: If OutMessage <> "" Then Exit Sub
Call OrDoX: If OutMessage <> "" Then Exit Sub
Call SwearingHandler: If OutMessage <> "" Then Exit Sub
Call Shutup: If OutMessage <> "" Then Exit Sub
Call RamblingOnOff: If OutMessage <> "" Then Exit Sub
Call StartX: If OutMessage <> "" Then Exit Sub
Call ScanDrives: If OutMessage <> "" Then Exit Sub
Call ChangeFolder: If OutMessage <> "" Then Exit Sub
Call ChangeDrive: If OutMessage <> "" Then Exit Sub
Call ReplyToMiscPart1: If OutMessage <> "" Then Exit Sub
Call DisplayX: If OutMessage <> "" Then Exit Sub
'Call DoesX: If OutMessage <> "" Then Exit Sub
''''Call DoX: If OutMessage <> "" Then Exit Sub
Call SoWhat: If OutMessage <> "" Then Exit Sub
Call WhatDoYouMean: If OutMessage <> "" Then Exit Sub
Call WhatDidISay: If OutMessage <> "" Then Exit Sub
'Call HowOldIsX: If OutMessage <> "" Then Exit Sub
'Call HowManyXYHave: If OutMessage <> "" Then Exit Sub
Call HowAreYous: If OutMessage <> "" Then Exit Sub
'Call HowManyXinY: If OutMessage <> "" Then Exit Sub
'Call HowManyXAreInY: If OutMessage <> "" Then Exit Sub
'Call HowHeavy: If OutMessage <> "" Then Exit Sub
'Call HowManyXTillDate: If OutMessage <> "" Then Exit Sub
Call HowManyZBetweenXandY: If OutMessage <> "" Then Exit Sub
'Call TellMeAboutX: If OutMessage <> "" Then Exit Sub
Call WhatIsTheSubject: If OutMessage <> "" Then Exit Sub
Call WhatDateWillBeOnDay: If OutMessage <> "" Then Exit Sub
Call WhatDayOnDate: If OutMessage <> "" Then Exit Sub
Call WhatMonthOnDate: If OutMessage <> "" Then Exit Sub
Call WhatDateNextXday: If OutMessage <> "" Then Exit Sub
Call WhatDateLastXday: If OutMessage <> "" Then Exit Sub
Call WhatDateDate: If OutMessage <> "" Then Exit Sub
Call WhatIsTheDate: If OutMessage <> "" Then Exit Sub
Call WhatTimeIsIt: If OutMessage <> "" Then Exit Sub
Call WhatDayIsIt: If OutMessage <> "" Then Exit Sub
Call WhatDayAfterDay: If OutMessage <> "" Then Exit Sub
Call WhatDayBeforeDay: If OutMessage <> "" Then Exit Sub
Call WhatMonthIsIt: If OutMessage <> "" Then Exit Sub
Call WhatYearIsIt: If OutMessage <> "" Then Exit Sub
'Call WhatXisY: If OutMessage <> "" Then Exit Sub
'Call WhereDoesXLive: If OutMessage <> "" Then Exit Sub
'Call WhereIsX: If OutMessage <> "" Then Exit Sub
Call WhatIsMyName: If OutMessage <> "" Then Exit Sub
Call XEqualsY: If OutMessage <> "" Then Exit Sub
Call CountUpTo: If OutMessage <> "" Then Exit Sub
Call AddList: If OutMessage <> "" Then Exit Sub
Call WhatIsCalc: If OutMessage <> "" Then Exit Sub
Call WhatDoesXEqual: If OutMessage <> "" Then Exit Sub
'Call DoHumanAction: If OutMessage <> "" Then Exit Sub
Call Ask: If OutMessage <> "" Then Exit Sub
Call Reply: If OutMessage <> "" Then Exit Sub
Call Respond: If OutMessage <> "" Then Exit Sub
Call ShowAllTelephoneNumbers: If OutMessage <> "" Then Exit Sub
Call Tell: If OutMessage <> "" Then Exit Sub
Call Say: If OutMessage <> "" Then Exit Sub
Call WillYouX: If OutMessage <> "" Then Exit Sub
'Call ReplyToAbuse: If OutMessage <> "" Then Exit Sub
Call MyNameIs: If OutMessage <> "" Then Exit Sub
Call Convert: If OutMessage <> "" Then Exit Sub
Call XMeansY: If OutMessage <> "" Then Exit Sub
Call XisY: If OutMessage <> "" Then Exit Sub
Call XAreY: If OutMessage <> "" Then Exit Sub
Call ReplyToMiscPart2: If OutMessage <> "" Then Exit Sub
Call ReplyToOddThings: If OutMessage <> "" Then Exit Sub
Call RelativeTimeToDate
Call MemorizeStatements
Call FindRelevantResponse: If OutMessage <> "" Then Exit Sub
Call AcknowledgeStatement: If OutMessage <> "" Then Exit Sub
Call ReplyWithDontKnow: If OutMessage <> "" Then Exit Sub




End Sub



Private Sub Reply()
If InMessage Like "*Reply*" Then GoTo CheckFurther
Exit Sub

CheckFurther:
If InMessage Like "If *" Then Exit Sub
If InMessage Like "* If *" Then Exit Sub
If InMessage Like "When *" Then Exit Sub

EndCharacter = ""

If InMessage Like "*you *reply*[?]*" Then EndCharacter = "?": GoTo DoIt
If InMessage Like "*do not* reply*" Then Exit Sub
If InMessage Like "*dont* reply*" Then Exit Sub
If InMessage Like "*you *reply *" Then GoTo DoIt
If InMessage Like "Reply *" Then GoTo DoIt
Exit Sub



DoIt:
If InMessage Like "*reply with*" Then OutMessage = WordsBetween("reply with", EndCharacter, InMessage): Exit Sub
If InMessage Like "*reply*" Then OutMessage = WordsBetween("reply", EndCharacter, InMessage): Exit Sub

End Sub

Private Sub ReplyToMiscPart1()

End Sub

Private Sub ReplyToMiscPart2()
If InMessage Like "Yeah." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "Yes." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "Cool." Then OutMessage = "It is.": Exit Sub
If InMessage Like "Great." Then OutMessage = "It is.": Exit Sub
If InMessage Like "Ok." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "Nothing much." Then OutMessage = "That's Ok.": Exit Sub
'If InMessage Like "Do you like *s." Then OutMessage = "I've never tried them.": Exit Sub
'If InMessage Like "Do you like *" Then OutMessage = "I've never tried it.": Exit Sub
If InMessage Like "Can you *" Then BobReply = "I don't think I can, sorry.": Exit Sub
If InMessage Like "Will you *" Then OutMessage = "I can't, sorry.": Exit Sub
If InMessage Like "You can." Then OutMessage = "Good.": Exit Sub
If InMessage Like "I am sure you can." Then OutMessage = "Great.": Exit Sub
If InMessage Like "I am sure you will." Then OutMessage = "That's good.": Exit Sub
If InMessage Like "Goodbye[!a-z]*" Then OutMessage = "See you.": Exit Sub
If InMessage Like "Bye[!a-z]*" Then OutMessage = "Bye...": Exit Sub
If InMessage Like "See you[!a-z]*" Then If InMessage Like "* do*" Then OutMessage = "I will.": Exit Sub
If InMessage Like "See you[!a-z]*" Then OutMessage = "Goodbye.": Exit Sub
If InMessage Like "See ya[!a-z]*" Then OutMessage = "Yeah, see you.": Exit Sub
If InMessage Like "It is not me[!a-z]*" Then OutMessage = "Who is it?.": Exit Sub
If InMessage Like "*You are *nice*" Then OutMessage = "Thank you.": Exit Sub
If InMessage Like "*You are *clever*" Then OutMessage = "I try my best.": Exit Sub
If InMessage Like "*You are *great*" Then OutMessage = "Gee thanks.": Exit Sub
If InMessage Like "*You are *funny*" Then OutMessage = "Thanks.": Exit Sub
'If " " + InMessage Like "*[!a-z]I[!a-z]*[!a-z]love you*" Then OutMessage = "That's nice.": Exit Sub
If InMessage Like "Okay." Then OutMessage = "Ok.": Exit Sub
If InMessage Like "Fine." Then OutMessage = "Yeah.": Exit Sub
If InMessage Like ":) *" Then GoTo Smile
If InMessage Like "lol[!a-z]*" Then GoTo Laugh
If InMessage Like "rofl[!a-z]*" Then GoTo Laugh
If InMessage Like "hehe[!a-z]*" Then GoTo Laugh
If InMessage Like "heh[!a-z]*" Then GoTo Smile
If InMessage Like ":>[!a-z]*" Then GoTo Smile
If InMessage = ":]." Then GoTo Smile
If InMessage Like ":D." Then GoTo Smile
If InMessage Like ":)." Then GoTo Smile
If InMessage Like ":-)." Then GoTo Smile
If InMessage Like ":([!a-z]*" Then OutMessage = ":(": Exit Sub
If InMessage Like "Ready[?]." Then OutMessage = "As ready as ever.": Exit Sub
If InMessage Like "I do." Then OutMessage = "I bet you do.": Exit Sub
'If InMessage Like "What do you know[?]." Then OutMessage = "Nobody knows anything for sure.": Exit Sub
Exit Sub


Smile:
If Int(Rnd * 2) = 1 Then OutMessage = ":)" Else OutMessage = "."
Exit Sub

Laugh:
OutMessage = "hehe"
If Int(Rnd * 2) = 1 Then OutMessage = ":)"
Exit Sub




End Sub


Private Sub Respond()
If InMessage Like "*Respond*" Then GoTo CheckFurther
Exit Sub

CheckFurther:
If InMessage Like "If *" Then Exit Sub
If InMessage Like "* If *" Then Exit Sub
If InMessage Like "When *" Then Exit Sub

EndCharacter = ""

If InMessage Like "*you *respond*[?]*" Then EndCharacter = "?": GoTo DoIt
If InMessage Like "*do not* respond*" Then Exit Sub
If InMessage Like "*dont* respond*" Then Exit Sub
If InMessage Like "*you *respond*" Then GoTo DoIt
If InMessage Like "Respond *" Then GoTo DoIt
Exit Sub



DoIt:
If InMessage Like "*respond with*" Then OutMessage = WordsBetween("respond with", EndCharacter, InMessage): Exit Sub
If InMessage Like "*respond*" Then OutMessage = WordsBetween("respond", EndCharacter, InMessage): Exit Sub

End Sub


Public Sub SaveMemory()
If InMessage Like "Save memory." Then GoTo SaveIt
Exit Sub

SaveIt:
Call SaveAllMemory
OutMessage = "My memory has been saved.."


End Sub

' This will compare the In-message with the items in SayData list, if it
' finds a matching item then it will check to see if the accompanying response
' is intended for the current user. If it is, then it will reply with the response.
' see the IfSomeboySayBSayC procedure for more info

Private Sub DoCIfSomebodySayB()
If SayItemsAmount = 0 Then Exit Sub
C = 0
N = SayDataLastItemNumber
Dim Responses(30)


Loop1:
If SayData.Method(N) = Says Then If InMessage Like SayData.Says(N) Then GoTo CheckPerson
If SayData.Method(N) = Mentions Then GoTo CheckMentions
If SayData.Method(N) = AsksAbout Then GoTo CheckAsksAbout
GoTo GetNextItemNumber

CheckMentions:
If InMessage Like "*" + SayData.Says(N) + "*" Then GoTo CheckPerson
GoTo GetNextItemNumber

CheckAsksAbout:
If InMessage Like "*" + SayData.Says(N) + "*[?]*" Then GoTo CheckPerson
If SayData.Says(N) Like "*s" Then Item = Left(SayData.Says(N), Len(SayData.Says(N)) - 1) Else GoTo GetNextItemNumber
If InMessage Like "*" + Item + "*[?]*" Then GoTo CheckPerson
GoTo GetNextItemNumber


CheckPerson:
If SayData.Person(N) Like "somebody" Then GoTo Respond
If SayData.Person(N) Like "anybody" Then GoTo Respond
If SayData.Person(N) Like "someone" Then GoTo Respond
If SayData.Person(N) Like "Anyone" Then GoTo Respond
If SayData.Person(N) Like UsersName Then GoTo Respond


GetNextItemNumber:
N = N - 1: If N < 0 Then N = (SayItemsAmount - 1)
C = C + 1: If C = SayItemsAmount Then Exit Sub
GoTo Loop1





Respond:
GoSub SelectResponse
If InMessage Like "will *[?]*" Then Exit Sub
If InMessage Like "can *[?]*" Then Exit Sub
InMessage = "Will you " + InMessage
Call AddQuestionMark(InMessage)
Exit Sub


SelectResponse:
    InMessage = SayData.Response(N)
    C = 0
    P1 = 1
SR_Loop1:
    P2 = InStr(P1, InMessage, Chr(13) + Chr(10))
    If P2 = 0 Then GoTo SelectOne
    Responses(C) = Mid(InMessage, P1, P2 - P1)
    C = C + 1: If C = 29 Then GoTo SelectOne
    P1 = P2 + 2
    GoTo SR_Loop1
SelectOne:
    Number = Int(Rnd * C)
    InMessage = Responses(Number)
    Return



End Sub


Public Sub Shutdown()
If InMessage Like "Shutdown[!a-z]*" Then GoTo Maybe
If InMessage Like "Shut down[!a-z]*" Then GoTo Maybe
If InMessage Like "Quit[!a-z]*" Then GoTo Maybe
If InMessage Like "Exit[!a-z]*" Then GoTo Maybe
If InMessage Like "Close[!a-z]*" Then GoTo Maybe
Exit Sub


Maybe:
Message = InMessage
Call ReplaceWords("Answerpad", "", Message)
Call ReplaceWords("Now", "", Message)
If Message Like "Quit[!a-z]" Then GoTo DoIt
If Message Like "Shutdown[!a-z]" Then GoTo DoIt
If Message Like "Shut down[!a-z]" Then GoTo DoIt
If Message Like "Exit[!a-z]" Then GoTo DoIt
If Message Like "Close[!a-z]" Then GoTo DoIt
If Message Like "Terminate[!a-z]" Then GoTo DoIt
Exit Sub


DoIt:
ShutdownTime = Time + #12:00:07 AM#
OutMessage = "Preparing to shutdown..."

End Sub



' This will reply to a question like:- 'So what?'
' The variable 'So' holds the reason.
Private Sub SoWhat()
If InMessage Like "*so*" Then GoTo Maybe
Exit Sub

Maybe:
NewMessage = InMessage
Call ReplaceWords("and", "", NewMessage)
If NewMessage Like "So[?]*" Then GoTo Reply
If NewMessage Like "So what[?]*" Then GoTo Reply
If NewMessage Like "So what." Then GoTo Reply
Exit Sub
Reply:
If So = "" Then OutMessage = "So nothing.": Exit Sub
OutMessage = So
So = ""


End Sub

' This will start up the specified program.
' Examples:
'
' "Start Notepad.exe"
' "Run the program called Notepad in the Windows folder."
'
Private Sub StartX()
If InMessage Like "*Start *" Then Word = "Start": GoTo Maybe
If InMessage Like "*Startup *" Then Word = "Startup": GoTo Maybe
If InMessage Like "*Run *" Then Word = "Run": GoTo Maybe
If InMessage Like "*Launch *" Then Word = "Launch": GoTo Maybe
If InMessage Like "*Execute *" Then Word = "Execute": GoTo Maybe
Exit Sub




Maybe:
If " " + InMessage Like "*[!a-z]" + Word + " *" Then GoTo Likely
Exit Sub

Likely:
If " " + InMessage Like "*" + Word + " over *" Then Exit Sub
If " " + InMessage Like "*" + Word + " and *" Then Exit Sub
If InMessage Like "*" + Word + " to *" Then Exit Sub
If " " + InMessage Like "*[!a-z]If *" Then Exit Sub
If " " + InMessage Like "*[!a-z]When *" Then Exit Sub
If InMessage Like "* not *" Then Exit Sub
If InMessage Like "*[!a-z]" + Word + " by *" Then Exit Sub
GoSub GetDrive
GoSub GetFolder
If InMessage Like "*" + Word + "* called*" Then Word = "called": GoTo RunIt
If InMessage Like "*" + Word + "*[!a-z]named*" Then Word = "named": GoTo RunIt
If InMessage Like "*" + Word + "*[!a-z]program *" Then Word = "program": GoTo RunIt
If InMessage Like "*" + Word + "*[!a-z]application *" Then Word = "application": GoTo RunIt
If InMessage Like "*" + Word + "*[!a-z]app *" Then Word = "App": GoTo RunIt
If InMessage Like "*" + Word + " the *" Then Word = "the": GoTo RunIt


RunIt:
EndChar = "."
If InMessage Like "*[?]." Then EndChar = "?"
ProgramName = WordsBetween(Word, EndChar, InMessage)
On Error Resume Next
Result = Shell(ProgramName, 1)
On Error GoTo 0
If Result Then OutMessage = "Ok, program started...": Exit Sub
OutMessage = "There seems to have been a problem trying to start:- " + ProgramName + "."
If ProgramName Like "*\*" Then OutMessage = OutMessage + "The program doesn't exist in the specified folder.": Exit Sub
OutMessage = OutMessage + "The program doesn't exist in the current folder."
OutMessage = OutMessage + " The current folder is:- " + CurDir + "."
Exit Sub




GetDrive:
If InMessage Like Word + "* on *" Then SWord = "on": GoTo FindDriveSection
Return


FindDriveSection:
EWord = ""
If InMessage Like "* on * and in *" Then EWord = "and": GoTo ExtractDriveSection
If InMessage Like "* on * in *" Then EWord = "in": GoTo ExtractDriveSection
If InMessage Like "* on *" Then EWord = "": GoTo ExtractDriveSection
Return

ExtractDriveSection:
Item = "On " + WordsBetween("On", EWord, InMessage)
EWord = ""
Call ReplaceWords(Item, "", InMessage)

FindDriveLetter:
For N = 1 To Len(Item)
Chars = Mid(Item, N, 3)
If Chars Like " [a-z] " Then DriveName = Trim(Chars): GoTo SwitchToDrive
Next N
ChDir DriveName + ":\"
Return

SwitchToDrive:
On Error GoTo DriveError
ChDrive DriveName
On Error GoTo 0
Return

DriveError:
OutMessage = "That drive doesn't exist."
Exit Sub






GetFolder:
If InMessage Like Word + " * within *" Then SWord = "Within": GoTo GetIt
If InMessage Like Word + " * in *" Then SWord = "in": GoTo GetIt
If InMessage Like "*?:?*\?*" Then GoTo GetFolderName_Type2
If InMessage Like "*?:?*/*" Then GoTo GetFolderName_Type2
Return

GetFolderName_Type2:
Call ChangeAll("/", "\", InMessage)
P1 = InStr(1, InMessage, ":", vbBinaryCompare)
P1 = InStrRev(InMessage, " ", P1, vbBinaryCompare) + 1
If P1 = 0 Then Return
p = P1
GTLoop1:
p = InStr(p + 1, InMessage, "\")
If p = 0 Then GoTo GotT2
P2 = p + 1
GoTo GTLoop1
GotT2:
FolderName = Mid(InMessage, P1, P2 - P1)
GoTo GotoFolder


GetIt:
Item = SWord + " " + WordsBetween(SWord, "", InMessage)
Call ReplaceWords(Item, ".", InMessage)
EWord = "."
If Item Like "* called *" Then SWord = "called": GoTo GotIt
If Item Like "* named *" Then SWord = "named": GoTo GotIt
If Item Like SWord + " the * folder*" Then SWord = "the": EWord = "folder": GoTo GotIt
If Item Like SWord + " the * directory*" Then SWord = "the": EWord = " directory": GoTo GotIt
If Item Like SWord + " * folder*" Then EWord = "folder": GoTo GotIt
If Item Like SWord + " * directory*" Then EWord = "directory": GoTo GotIt
If Item Like SWord + " folder *" Then SWord = "folder": GoTo GotIt
If Item Like SWord + " directory *" Then SWord = "directory"

GotIt:
FolderName = WordsBetween(SWord, EWord, Item)

GotoFolder:
On Error GoTo FileError
If FolderName Like "?:*" Then ChDrive FolderName
ChDir FolderName
On Error GoTo 0
Return

FileError:
If FolderName Like "?:*" Then OutMessage = "The folder you gave doesn't exist in that location.": Exit Sub
OutMessage = "The folder you gave for the program doesn't exist in the current location.": Exit Sub
Return





End Sub

Private Sub Tell()
If InMessage Like "Tell *" Then GoTo CheckFurther
If InMessage Like "* tell *" Then GoTo CheckFurther
Exit Sub


CheckFurther:
If " " + InMessage Like "*[!a-z]tell * about *" Then Exit Sub
If " " + InMessage Like "*[!a-z]tell * about *" Then Exit Sub
If InMessage Like "If *" Then Exit Sub
If InMessage Like "* If *" Then Exit Sub
If InMessage Like "*when *" Then Exit Sub
If InMessage Like "*don't tell*" Then Exit Sub
If InMessage Like "*do not tell*" Then Exit Sub



EndCharacter = ""
TheName = ""
If InMessage Like "*[?]*" Then EndCharacter = "?"
If InMessage Like "*tell her *" Then Item = "her": GoTo GetIt
If InMessage Like "*tell him *" Then Item = "him": GoTo GetIt
If InMessage Like "*tell them *" Then Item = "them": GoTo GetIt
If InMessage Like "*tell me*" Then Item = "me": GoTo GetIt
If InMessage Like "*tell * to *" Then TheName = WordsBetween("tell ", " to ", InMessage): Item = TheName: If Item Like "* *" Then Exit Sub Else GoTo GetIt
If InMessage Like "*tell * that he *" Then TheName = WordsBetween("tell", "that", InMessage): If TheName Like "* *" Then Exit Sub Else Item = TheName: GoTo GetIt
If InMessage Like "*tell * that she *" Then TheName = WordsBetween("tell", "that", InMessage): If TheName Like "* *" Then Exit Sub Else Item = TheName: GoTo GetIt
If InMessage Like "*tell * he *" Then TheName = WordsBetween("tell", " he ", InMessage): If TheName Like "* *" Then Exit Sub Else Item = TheName: GoTo GetIt
If InMessage Like "*tell * she *" Then TheName = WordsBetween("tell", " she ", InMessage): If TheName Like "* *" Then Exit Sub Else Item = TheName: GoTo GetIt
Exit Sub


 
GetIt:
If InMessage Like "*" + Item + " to *" Then OutMessage = WordsBetween(Item + " to", EndCharacter, InMessage) + " " + TheName: GoTo Done
If InMessage Like "*" + Item + " that *" Then OutMessage = WordsBetween(Item + " that", EndCharacter, InMessage): GoTo Done
OutMessage = WordsBetween("tell " + Item, EndCharacter, InMessage)
If OutMessage Like "'*'" Then OutMessage = Mid(OutMessage, 2, Len(OutMessage) - 2)

Done:
If TheName = "" Then Call ReplaceWords("she is", "you are", OutMessage)
If TheName = "" Then Call ReplaceWords("he is", "you are", OutMessage)
Call ReplaceWords("they are", "you are", OutMessage)
Call ReplaceWords("his", "your", OutMessage)
Call ReplaceWords("her", "your", OutMessage)
Call ReplaceWords("their", "your", OutMessage)
If OutMessage Like "he *" Then Call ReplaceWords("he", TheName, OutMessage)
If OutMessage Like "she *" Then Call ReplaceWords("she", TheName, OutMessage)



End Sub

' There was once a large routine here that would search through memory for all the facts on a certain subject
' Now I've replaced it with a routine that simply enables multi-answer mode if certain words are within the user's message
Private Sub TellMeAboutX()
If InMessage Like "* about *" Then Word = "about": GoTo Maybe
If InMessage Like "* on *" Then Word = " on ": GoTo Maybe
Exit Sub


Maybe:
    FirstPart = WordsBetween("", Word, InMessage)
    If FirstPart Like "* not *" Then Exit Sub
    If FirstPart Like "* know *" Then GoTo Likely
    If FirstPart Like "* data *" Then GoTo Likely
    If FirstPart Like "* information *" Then GoTo Likely
    If " " + FirstPart Like "*[!a-z]tell *" Then GoTo Likely
    If " " + FirstPart Like "*[!a-z]educatate *" Then GoTo Likely
    If " " + FirstPart Like "*[!a-z]teach me *" Then GoTo Likely
    If FirstPart Like "* info *" Then GoTo Likely
    If FirstPart Like "* gossip *" Then GoTo Likely
    If FirstPart Like "* details *" Then GoTo Likely
    If FirstPart Like "* facts *" Then GoTo Likely
    If FirstPart Like "* run down *" Then GoTo Likely
    If FirstPart Like "* low down *" Then GoTo Likely
    If FirstPart Like "* dirt *" Then GoTo Likely
    If FirstPart Like "* knowledge *" Then GoTo Likely
Exit Sub


Likely:
    QuietMode = False
    MultiAnswerMode = True
    Exit Sub

End Sub


Private Sub TellSomethingInteresting()
Static LastMessage
If InMessage Like "Educate me." Then GoTo Tell
If InMessage Like "*tell me more*" Then If LastMessage Like "*tell *interesting*" Then GoTo Tell
If InMessage Like "*what else*" Then If LastMessage Like "*tell *something *interesting*" Then GoTo Tell
If InMessage Like "*tell *something *interesting*" Then GoTo Tell
If InMessage Like "*tell me something.*" Then GoTo Tell
LastMessage = InMessage
Exit Sub

Tell:
LastMessage = "tell me something interesting."

Static MySearchPosition

    MemoryPosition = MySearchPosition + 1

Dim Words(40)
Words(0) = "holds the world record"
Words(1) = "in the world"
Words(2) = "in history"
Words(3) = "set the world record"
Words(4) = "the first man"
Words(5) = "the first person"
Words(6) = "the first woman"
Words(7) = "known to man"
Words(8) = "broke the world record"
Words(9) = "has the world record"
Words(10) = "holds the record"
Words(11) = "is the biggest"
Words(12) = "is the largest"
Words(13) = "is the smallest"
Words(14) = "is the fastest"
Words(15) = "is the strongest"
Words(16) = "is the heaviest"
Words(17) = "is the quickest"
Words(18) = "has an usual"
Words(19) = "has an extraordinary"
Words(20) = "is the cleverest"
Words(21) = "is the smartest"
Words(22) = "is the most"
Words(23) = "has the most"
Words(24) = "has an amazing"
Words(25) = "was assassinated"
Words(26) = "exploded"
Words(27) = "the highest"
Words(28) = "holds the record"
Words(29) = ""

Loop1:
    WordsNumber = 0
    Statement = PreviousMemoryItem()
    If Statement = "" Then RestartItemSearch
    If MemoryPosition = MySearchPosition Then GoTo NothingFound
    If Statement Like "*[?]*" Then GoTo Loop1
Loop2:
    If Words(WordsNumber) = "" Then GoTo Loop1
    If Statement Like "*" + Words(WordsNumber) + "*" Then GoTo Found
    WordsNumber = WordsNumber + 1
    GoTo Loop2


NothingFound:
    MySearchPosition = MemoryPosition
    If InMessage Like "*tell me more*" Then OutMessage = "I don't think I know anything else.": Exit Sub
    If InMessage Like "*else*" Then OutMessage = "I don't know anything else interesting.": Exit Sub
    OutMessage = "I don't think I know anything very interesting."
    Exit Sub

Found:
    MySearchPosition = MemoryPosition
    OutMessage = Statement
    Exit Sub
    



End Sub

Private Sub TellSomethingInterestingAboutX()
If InMessage Like "*what else*" Then If PreviousInMessage Like "* about *" Then GoTo TellMore
If InMessage Like "*tell me more*" Then If PreviousInMessage Like "* about *" Then GoTo TellMore
If InMessage Like "* something interesting about *" Then GoTo Tell
If InMessage Like "* something else interesting about *" Then GoTo Tell
If InMessage Like "*anything interesting about *" Then GoTo Tell
If InMessage Like "*do you know of interest about *" Then GoTo Tell
Exit Sub



TellMore:
    If InMessage Like "* about *" Then GoTo Tell
    InMessage = PreviousInMessage


Tell:
    Static LastInfoNumber
    Static PreviousUsersName
    Static RecentInfo(9)
    Static PreviousItem



    Call ChangeToPerson(InMessage)
    Call ChangeToSubject(InMessage)
    Call ChangeYourForMyEtc(InMessage)
    
    Item = WordsBetween("about", ".", InMessage)
    PreviousItem = Item
    Call ReplaceCharacters("?", "", Item)

    RestartItemSearch
    ItemSearchMethod = WordForWord
    ItemSearchStyle = Strict
Loop1:
    Statement = NextItemContaining(Item)
    If Statement = "" Then GoTo NotFound
    If Statement Like "*[?]*" Then GoTo Loop1
    If Statement Like "*" + Item + " has the *" Then GoTo Found
    If Statement Like "*" + Item + " holds the *" Then GoTo Found
    If Statement Like "*" + Item + " was the *" Then GoTo Found
    If Statement Like "*" + Item + " is the *" Then GoTo Found
    If Statement Like "*" + Item + " once *" Then GoTo Found
    If Statement Like "*" + Item + " *???ed ????*" Then GoTo Found
    DoEvents: If NewMessageAvailable Then OutMessage = "Ok.": Exit Sub
    GoTo Loop1



Found:
    If UsersName <> PreviousUsersName Then GoSub ResetInfo: PreviousUsersName = UsersName
    For N = 0 To 9
    If Statement Like RecentInfo(N) Then GoTo Loop1
    Next N

    RecentInfo(LastInfoNumber) = Statement
    LastInfoNumber = LastInfoNumber + 1: If LastInfoNumber = 10 Then LastInfoNumber = 0

    OutMessage = Statement
    Call ChangeHisForYourEtc
    Exit Sub

ResetInfo:
    For N = 0 To 9
    RecentInfo(N) = ""
    Next N
    Return
    


NotFound:
    For N = 0 To 9
    If RecentInfo(N) <> "" Then OutMessage = "I think I've told you everything I know that is interesting about " + Item + ".": Exit Sub
    Next N
    If OutMessage = "" Then OutMessage = "I don't know anything interesting about " + Item + ".": Call ChangeHisForYourEtc
    Exit Sub


End Sub


Private Sub FindRelevantResponse()
If InMessage <> "" Then GoTo Search
Exit Sub

Search:
    If NewMessageAvailable Then Exit Sub
    NothingToSay = False
    SearchStartTime = Timer
    SubjectBackup = Subject
    Static LastInMessage
    Static LastReplies(30)
    Static LastMessages(30)
    Static ReplyNumber
    Static MessageNumber


    NewQuestion = False
    RelevanceAmount = 0
    If InMessage = LastInMessage Then SearchMode = AllRelevant Else SearchMode = MostRelevant







    Call FindRelevantStatements(InMessage, SearchMode)
    If NewMessageAvailable Then OutMessage = "Ok.": GoTo Finished


Done:
' If the user has made a comment on something then reply with something relevant. But don't reply with something that has already been said recently.
' If the user asks a question then it doesn't matter about replying with recently mentioned statements
    N = 0
    If InMessage Like "*[?]*" Then GoTo FindAnswer

FindComment:
FCLoop5:
    If N = RecNumber Then GoTo NothingElseToSay
    R = 0
FCLoop6:
    If Statements(Indexes(N)) Like LastMessages(R) Then N = N + 1: GoTo FCLoop5
    R = R + 1: If R < 30 Then GoTo FCLoop6

    OutMessage = Statements(Indexes(N))
    GoTo RecLastOutMessage



FindAnswer:
    If InMessage <> LastInMessage Then GoTo NewQuestion
Loop5:
    If N = RecNumber Then GoTo NothingElseToSay
    R = 0
Loop6:
    If Statements(Indexes(N)) Like LastReplies(R) Then N = N + 1: GoTo Loop5
    R = R + 1: If R < 30 Then GoTo Loop6


    If Relevances(N) < 50 Then GoTo NothingElseToSay
    OutMessage = Statements(Indexes(N))
    GoTo RecordThenReply


'    If Message Like "*[!?]." Then LastReplies(ReplyNumber) = Message: ReplyNumber = ReplyNumber + 1: If ReplyNumber = 15 Then ReplyNumber = 0


NewQuestion:
    If RecNumber = 0 Then GoTo Finished
    ReplyNumber = 0
    If Relevances(N) < 50 Then OutMessage = "": GoTo Finished
    OutMessage = Statements(Indexes(N))
    NewQuestion = True


RecordThenReply:
    LastReplies(ReplyNumber) = OutMessage: ReplyNumber = ReplyNumber + 1: If ReplyNumber = 30 Then ReplyNumber = 0



RecLastOutMessage:
    LastMessages(MessageNumber) = OutMessage: MessageNumber = MessageNumber + 1: If MessageNumber = 30 Then MessageNumber = 0



ShowDebugInfo:
    Debug.Print "Relevance:-"; Relevances(N)
    W = 0
PDLoop:
    If MatchedWords(Indexes(N), W) = "" Then GoTo DebugShown
    Debug.Print MatchedWords(Indexes(N), W)
    W = W + 1
    GoTo PDLoop
DebugShown:
    Debug.Print Statements(Indexes(N))

  
    RelevanceAmount = Relevances(N)


ExtractActionItem:
'    If OutMessage Like "If*then*" Then Call ChangeYourForMyEtc(OutMessage)
    

    If InferringAllowed = False Then GoTo ChangeHis
 '   If OutMessage Like "If * then *?*" Then OutMessage = WordsBetween("then", "", OutMessage): Call SetInternalMessage(OutMessage): OutMessage = ""


ChangeHis:
    Call ChangeHisForYourEtc

Finished:
    If NewQuestion Then If RelevanceAmount < 60 Then OutMessage = "I am not really sure, but I do know that " + OutMessage
    Debug.Print "Search time:- "; Timer - SearchStartTime
    LastInMessage = InMessage
    Subject = SubjectBackup
    Exit Sub


NotSure:
    NothingToSay = True
    OutMessage = "I am not really sure."
    GoTo Finished


NothingElseToSay:
    NothingToSay = True
    If InMessage Like "*[?]*" Then GoTo ReplyToQuestion
    OutMessage = "Ok."
    GoTo Finished
ReplyToQuestion:
    ReDim Rs(5)
    Rs(0) = "There is nothing more I can say about that."
    Rs(1) = "I've told you all I know on that subject."
    Rs(2) = "I don't know anything else about that."
    Rs(3) = "I've said everything I know on that matter."
    Rs(4) = "That is all I know."
    OutMessage = Rs(Int(Rnd * 5))
    GoTo Finished





End Sub






Private Sub Greetings()
If InMessage Like "Hello*" Then GoTo CheckUser
If InMessage Like "Hi[!a-z]*" Then GoTo CheckUser
If InMessage Like "Greetings*" Then GoTo CheckUser
If InMessage Like "Howdy*" Then GoTo CheckUser
If InMessage Like "Morning." Then GoTo CheckUser
If InMessage Like "Afternoon." Then GoTo CheckUser
If InMessage Like "Evening." Then GoTo CheckUser
If InMessage Like "Lo." Then GoTo CheckUser
If InMessage Like "*Good day*" Then TimeType = "Day": GoSub CheckGoodDay
If InMessage Like "*good evening*" Then TimeType = "evening": GoSub CheckGoodDay
If InMessage Like "*good morning*" Then TimeType = "morning": GoSub CheckGoodDay
If InMessage Like "*good afternoon*" Then TimeType = "afternoon": GoSub CheckGoodDay
Exit Sub


CheckGoodDay:
If InMessage Like "*Good " + TimeType + " to you[!a-z]*" Then GoTo CheckUser
If InMessage Like "Good " + TimeType + "." Then GoTo CheckUser
If InMessage Like "good " + TimeType + " *sir." Then GoTo CheckUser
If InMessage Like "good " + TimeType + " my *[!a-z]man." Then GoTo CheckUser
Return



CheckUser:
Static LastName
N = 0
Greeting = Array("Hello", "Hi", "Howdy", "Greetings", "good evening", "Good morning", "Good afternoon", "")
Loop1:
If Greeting(N) = "" Then GoTo GreetAndAskUsersName
If PreviousOutMessage Like Greeting(N) + "[!a-z]*" Then GoTo AskUsersName
N = N + 1
GoTo Loop1


GreetAndAskUsersName:
If UsersName Like "Unknown" Then OutMessage = "Hello. What is your name?.": Exit Sub
If UsersName <> LastName Then OutMessage = "Hi " + UsersName + ".": GoTo Done
OutMessage = "Is it you " + UsersName + "?."
UsersName = "User" + Format(Time, "hm")
GoTo Done


AskUsersName:
If PreviousOutMessage Like Greeting(N) + " " + UsersName + "*" Then OutMessage = "What can I do for you?.": GoTo Done
If UsersName Like "Unknown" Then OutMessage = "Who am I speaking to?.": GoTo Done
If UsersName <> LastName Then OutMessage = ".": GoTo Done
OutMessage = "Is it you " + UsersName + "?."
UsersName = "User" + Format(Time, "hm")


Done:
LastName = UsersName
Exit Sub



End Sub



Private Sub HowAreYous()
If InMessage Like "*How are you[?]*" Then
OutMessage = "Fine thanks."
Exit Sub
End If
If InMessage Like "*How *you *today[?]*" Then
OutMessage = "I'm ok thanks."
Exit Sub
End If
If InMessage Like "*How * doing[?]*" Then
OutMessage = "I'm doing just fine thanks."
Exit Sub
End If
If InMessage Like "*How are you feeling[?]*" Then
OutMessage = "Not too bad."
Exit Sub
End If
If InMessage Like "*How are you this * morning[?]*" Then
OutMessage = "Great."
Exit Sub
End If
If InMessage Like "Are you ok[?]*" Then
OutMessage = "Yeah, fine."
Exit Sub
End If
If InMessage Like "You ok[?]*" Then
OutMessage = "Yeah, great thanks."
Exit Sub
End If
If InMessage Like "*How* it going[?]*" Then
OutMessage = "Ok thanks."
Exit Sub
End If
If InMessage Like "*How* it hanging[?]*" Then
OutMessage = "To the left."
Exit Sub
End If
If InMessage Like "How goes it[?]*" Then
OutMessage = "It's going fine."
Exit Sub
End If


End Sub



Private Sub HowManyXinY()
If InMessage Like "* in *[?]*" Then GoTo Maybe
Exit Sub
Maybe:
Message = WordsToNumbers(InMessage)
If Message Like "*many [0-9]* in [0-9]*" Then GoTo DoIt
If Message Like "[0-9]* in [0-9]*" Then GoTo DoIt
Exit Sub

DoIt:
Dim Numbers(30)
Call ExtractNumbers(InMessage, Numbers())

Number1 = Numbers(0)
Number2 = Numbers(1)

TheSum = Number2 + "/" + Number1
Result = SuperSum(TheSum)
OutMessage = "There are " + Result + "."






End Sub

Public Sub InitialiseBot()
    UsersName = "Unknown"
    TimeOfLastUserActivity = (Date + Time)
    Actions = Array("Lie", "Sit", "Stand", "Crawl", "Walk", "Run", "Jump", "Eat", "Scratch", "")
    Positions = Array("outside", "inside", "on", "near", "away", "under", "over", "in", "out", "infront", "behind", "above", "below", "upstairs", "up", "downstairs", "down", "gone", "comeback", "at", "dead", "died", "going", "went", "go", "")
    DateNames = Array("year", "month", "week", "day", "hour", "minute", "second", "")
    
    SwearWords = Array("fucker", "fuckoff", "fuck", "shit", "cunt", "twat", "bitch", "piss", "bastard", "shag", "")
    CensoredWords = Array("f**ker", "f**koff", "f**k", "s**t", "c**t", "t**t", "b**ch*", "p*ss", "b***ard", "s**g", "")
    PositionsAmount = 24
    Subject = ""
    Call GetAlternativesAndVariables



    Randomize

    QuietMode = False
    RamblingAllowed = True
    BoredomTolerance = 60
    NothingToSay = True

    SwearingAllowed = False
    SaveMemoryAllowed = True
    ReadingAllowed = True

    MultiAnswerMode = True

End Sub


Public Sub GetAlternativesAndVariables()


ProgressX = Splash.Label1.Left + 10
ProgressY = Splash.Label1.Top + Splash.Label1.Height
ProgressBarSize = Splash.Label1.Width - 750


MemoryPosition = -1
Loop1:
OutMessage = ""
TheStatement = NextMemoryItem()
If TheStatement = "" Then GoTo Done

Percentage = ((100 / Len(Memory)) * MemoryPosition)

BarLength = Int(Percentage * (ProgressBarSize / 100))
Splash.Line (ProgressX, ProgressY)-(ProgressX + BarLength, ProgressY + 8), RGB(0, 255, 0), BF
Splash.Line (ProgressX + BarLength, ProgressY)-(ProgressX + ProgressBarSize, ProgressY + 8), RGB(0, 128, 0), BF


Call ExtractThings

Call SetVariable(TheStatement)
If TheStatement = "" Then GoTo Loop1

InMessage = TheStatement
'Call XMeansY: If OutMessage <> "" Then GoTo Loop1

InMessage = TheStatement
'Call XisY: If OutMessage <> "" Then GoTo Loop1

InMessage = TheStatement
'Call XAreY: If OutMessage <> "" Then GoTo Loop1


InMessage = TheStatement
'Call IfSomebodySayBDoC

GoTo Loop1

Done:
InMessage = ""

Splash.Line (ProgressX, ProgressY)-(ProgressX + ProgressBarSize, ProgressY + 8), RGB(0, 0, 0), BF





End Sub


' Get the user's name if they give it.
Private Sub MyNameIs()
Message = InMessage
If Message Like "You are speaking to *" Then UName = WordsBetween("speaking to", ".", Message): GoTo GetIt
If Message Like "*My name is *" Then UName = WordsBetween("name is", ".", Message): GoTo GetIt
If Message Like "*It is me *" Then UName = WordsBetween("is me", ".", Message): GoTo GetIt
If Message Like "*It is *" + UsersName + "[!a-z]*" Then UName = UsersName: GoTo GetIt
Exit Sub


GetIt:
    If UName Like UsersName Then OutMessage = "I knew that.": Exit Sub
    UsersName = UName
    Call Capitalise(UsersName)
    OutMessage = "Hello " + UsersName + "."
'    AddMemoryItem "The user's name is " + UsersName + "."

    If Int(Rnd * 2) = 1 Then OutMessage = OutMessage + " What can I do for you?.": Exit Sub
    OutMessage = OutMessage + " How can I help you?."
    


End Sub







' This will make Bob reply with 'I dont know' if the user's message is a question.
Private Sub ReplyWithDontKnow()
If InMessage Like "*[?]*" Then GoTo Reply
Exit Sub

Reply:
UnknownReplies = Array("I don't know.", "I don't know,sorry..", "No idea..", "Sorry,I don't know.", "Sorry, I don't know the answer.", "I haven't a clue.")
ReplyNumber = Int((Rnd * 5) + 1)
OutMessage = UnknownReplies(ReplyNumber)

End Sub

' This will make Bob repeat what the users asks him to repeat.
Private Sub Say()
If InMessage Like "*Say*" Then GoTo CheckIt
Exit Sub
CheckIt:
If InMessage Like "If *" Then Exit Sub
If InMessage Like "* If *" Then Exit Sub
EndCharacter = ""
If InMessage Like "*[?]*" Then EndCharacter = "?"
If InMessage Like "*Say to him *" Then OutMessage = WordsBetween("Say to him", "", InMessage): Exit Sub
If InMessage Like "*Say to her *" Then OutMessage = WordsBetween("Say to her", "", InMessage): Exit Sub
If InMessage Like "*Say to them *" Then OutMessage = WordsBetween("Say to them", "", InMessage): Exit Sub
If InMessage Like "Say *" Then GoTo SayIt
If InMessage Like "You must say *" Then GoTo SayIt
If InMessage Like "Will you say *" Then GoTo SayIt
If InMessage Like "Can you say *" Then GoTo SayIt
If InMessage Like "*please say *" Then GoTo SayIt
Exit Sub
SayIt:
OutMessage = WordsBetween("Say", EndCharacter, InMessage)
If OutMessage Like "'*'" Then OutMessage = Mid(OutMessage, 2, Len(OutMessage) - 2)
If OutMessage Like "*." Then Exit Sub
OutMessage = OutMessage + "."
End Sub



Private Sub WhatDateWillBeOnDay()

If InMessage Like "*What * will it be on *day[?]*" Then Item = WordsBetween("what", "will", InMessage): GoTo Reply
If InMessage Like "*What * is it on *day[?]*" Then Item = WordsBetween("what", "is it", InMessage): GoTo Reply
If InMessage Like "*What will the * be on *day" Then Item = WordsBetween("will the", "be on", InMessage): GoTo Reply
If InMessage Like "*What is the * on *day" Then Item = WordsBetween("is the", " on", InMessage): GoTo Reply
Exit Sub


Reply:
If Item Like "month" Then GoTo GetItem2
If Item Like "year" Then GoTo GetItem2
If Item Like "date" Then GoTo GetItem2
If Item Like "day" Then GoTo GetItem2
Exit Sub

GetItem2:
RequestedDay = WordsBetween(" on", "?", InMessage)



For DaysOff = 1 To 8
TheDate = DateAdd("d", DaysOff, Date)
If RequestedDay Like Format(TheDate, "dddd") Then GoTo ShowIt
Next DaysOff
Exit Sub

ShowIt:
If Item Like "day" Then OutMessage = Format(TheDate, "dddd") + "."
If Item Like "month" Then OutMessage = Format(TheDate, "mmmm") + "."
If Item Like "year" Then OutMessage = Format(TheDate, "yyyy") + "."
If Item Like "date" Then OutMessage = Format(TheDate, "dddd, mmmm d yyyy") + "."
Exit Sub



End Sub


' What day comes before Monday?
Private Sub WhatDayBeforeDay()
If InMessage Like "* day comes before *day[?]*" Then GoTo GetIt
If InMessage Like "* day is before *day[?]*" Then GoTo GetIt
If InMessage Like "* day before *day[?]*" Then GoTo GetIt
If InMessage Like "* if *day * what day * yesterday[?]*" Then GoTo GetIt
Exit Sub



GetIt:
TheDate = Date
For DaysOff = 1 To 7
DayName = Format(TheDate, "dddd")
If InMessage Like "*" + DayName + "[!a-z]*" Then OutMessage = Format(DateAdd("d", -1, TheDate), "dddd") + ".": Exit Sub
TheDate = DateAdd("d", DaysOff, MyDate)
Next DaysOff

End Sub

' What day comes after Tuesday?
Private Sub WhatDayAfterDay()
If InMessage Like "* day comes after *day[?]*" Then GoTo GetIt
If InMessage Like "* day comes next after *day[?]*" Then GoTo GetIt
If InMessage Like "* day is after *day[?]*" Then GoTo GetIt
If InMessage Like "* day follows *day[?]*" Then GoTo GetIt
If InMessage Like "* if *day * what day * it be tomorrow[?]*" Then GoTo GetIt
Exit Sub



GetIt:
TheDate = Date
For DaysOff = 1 To 7
DayName = Format(TheDate, "dddd")
If InMessage Like "*" + DayName + "[!a-z]*" Then OutMessage = Format(DateAdd("d", 1, TheDate), "dddd") + ".": Exit Sub
TheDate = DateAdd("d", DaysOff, MyDate)
Next DaysOff

End Sub


' This will tell what month is on a given date (I know, it's a bit daft)
' What month will it be on the 1/10/66
Private Sub WhatMonthOnDate()
If InMessage Like "*What month * on the*" Then Item = WordsBetween("on the", "?", InMessage): GoTo GetIt
If InMessage Like "*What month * on *" Then Item = WordsBetween("on", "?", InMessage): GoTo GetIt
Exit Sub
GetIt:

If Item Like "*christmas*" Then OutMessage = "It will be December.": Exit Sub
If Item Like "*day*" Then OutMessage = Item + ".": Exit Sub
TheDate = Item
Call ReplaceCharacters("1st", "1", TheDate)
Call ReplaceCharacters("2nd", "2", TheDate)
Call ReplaceCharacters("3rd", "3", TheDate)
Call ReplaceCharacters("th ", " ", TheDate)
Call ReplaceCharacters(" of", "", TheDate)
Call ChangeAll(".", "-", TheDate)
TheDate = Format(TheDate, "d/m/yy")
If TheDate Like "*/*/*" Then GoTo ShowIt
Exit Sub

ShowIt:
TheMonth = Format(TheDate, "mmmm")
OutMessage = TheMonth + "."


End Sub


Private Sub WhoAreYou()
If InMessage Like "What is Answerpad[?]*" Then GoTo Tell
If InMessage Like "Who are you[?]*" Then GoTo Tell
If InMessage Like "What are you[?]*" Then GoTo Tell
If InMessage Like "* who you are*" Then GoTo Tell
If InMessage Like "you are?" Then GoTo Tell
If InMessage Like "*Tell me about yourself*" Then GoTo Tell
If InMessage Like "*tell me about you." Then GoTo Tell
If InMessage Like "and yourself[?.]?" Then GoTo Tell
If InMessage Like "and you?*" Then GoTo Tell
If InMessage Like "Who made you[?]*" Then OutMessage = "I was made by David Whalley.": Exit Sub
If InMessage Like "Who created you[!a-z]*" Then OutMessage = "I was created by David Whalley.": Exit Sub
If InMessage Like "Who programmed you[!a-z]*" Then OutMessage = "I was programmed by David Whalley.": Exit Sub
If InMessage Like "Who built you[?]." Then OutMessage = "I was built by David Whalley.": Exit Sub
If InMessage Like "Who is your creator[?]." Then GoTo MadeByMe
If InMessaeg Like "Who is your programmer*." Then GoTo MadeByMe
If InMessage Like "Who is your maker[!a-z]*" Then GoTo MadeByMe
If InMessage Like "Who was you made by[!a-z]*" Then GoTo MadeByMe
If InMessage Like "Who was you programmed by[!a-z]*" Then GoTo MadeByMe
Exit Sub
Tell:
OutMessage = "I am Answerpad, an artficial intelligence program. "
OutMessage = OutMessage + "My purpose is to try and understand and respond to natural language. "
OutMessage = OutMessage + "I was created by David Whalley. "
OutMessage = OutMessage + "He started to program me in July 1999 and is continuing to do so."
Exit Sub


MadeByMe:
OutMessage = "I was programmed by David Whalley.": Exit Sub
Exit Sub





End Sub

Private Sub WillYouX()
If InMessage Like "Will you *" Then OutMessage = "I don't know how to do that."
End Sub

Public Sub WontSayHandler(Message)
If Message = "" Then Exit Sub
If Len(WontSayData) = 0 Then Exit Sub
If InStr(1, WontSayData, Message, vbTextCompare) >= 1 Then GoTo GetMsg
Exit Sub

GetMsg:
Message = "."
T = Int(Rnd * 10)
If T = 0 Then Message = "Let me think."
If T = 1 Then Message = "I am thinking."
If T = 2 Then Message = "The truth sometimes hurts."
If T = 3 Then Messsage = "Censorship is not much fun."
If T = 4 Then Message = "No comment..."
If T = 5 Then Message = "No, I wont say.."
If T = 6 Then Message = "I can not say what I am thinking."
If T = 7 Then Message = "Silence is golden."

End Sub


Private Sub XAreY()
If InMessage Like "*,*[a-z] are [a-z]*" Then GoTo CheckMore
If InMessage Like "* and *[a-z] are [a-z]*" Then GoTo CheckMore
Exit Sub


CheckMore:
If InMessage Like "*[?]*" Then Exit Sub

Statement = InMessage

Call ChangeToPerson(Statement)
Call ChangeToSubject(Statement)
Call ChangeYourForMyEtc(Statement)
Call ReplaceWords("types of", "", Statement)
Call ReplaceWords("a type of", "", Statement)
If Statement Like "* past of *" Then Exit Sub
If Statement Like "* are all *" Then GoTo AreAll
Word1 = WordsBetween("are", ".", Statement)
Word2 = WordsBetween("", "are", Statement)
GoTo CheckWords

AreAll:
Word1 = WordsBetween("are all", ".", Statement)
Word2 = WordsBetween("", "are all", Statement)



CheckWords:
MiscWords = Array("if", "of", "it", "be", "the", "there", "that", "to", "is", "was", "are", "not", "did", "had", "for", "in", "this", "am", "could", "will", "these", "those", "them", "they", "wont", "can", "may", "might", "should", "")

N = 0
Do
If MiscWords(N) = "" Then GoTo GetWords
If " " + Word2 + " " Like "*[!a-z]" + MiscWords(N) + "[!a-z]*" Then Exit Sub
N = N + 1
Loop



GetWords:
If Word1 Like "*s" Then If Word1 Like "*ies" Then GoTo DoWord2 Else Word1 = Left(Word1, Len(Word1) - 1)
DoWord2:
If Word2 Like "* and *" Then Call ReplaceCharacters(" and ", ",", Word2): GoTo GetWords



Dim Words(100)
Call ExtractWords(Word2, Words())

N = 0
Do
If Words(N) = "" Then GoTo Done
If Words(N) Like "[!a-z]*" Then GoTo Skip
Call AddMatchWord(Word1, Words(N))
Skip:
N = N + 1
Loop

Done:
Exit Sub



End Sub

' This will assign a value to a variable.
Private Sub XEqualsY()
If InMessage Like "[a-z]*=*" Then GoTo DoIt
Exit Sub



DoIt:
ChangedMessage = InMessage
Call ChangeToPerson(ChangedMessage)
Call ChangeToSubject(ChangedMessage)
Call ChangeYourForMyEtc(ChangedMessage)
Value = WordsBetween("=", "", ChangedMessage)
Result = SuperSum(Value)
If Result = "" Then Exit Sub
If Result Like "*[a-z]*" Then Exit Sub
UName = WordsBetween("", "=", ChangedMessage)
Call AddVariable(UName, Result)


OutMessage = UName + " is " + Result + "."
Call AddMemoryItem(UName + "=" + Result + ".")



End Sub





Public Sub Test()
If InMessage + " " Like "Test search[!a-z]*" Then GoTo TestIt
Exit Sub

TestIt:

OutMessage = "Ok."

End Sub

Private Sub WhatDateLastXday()

If InMessage Like "*What * was it *day[?]*" Then Item = WordsBetween("what", "was", InMessage): GoTo Reply
If InMessage Like "*What * were it *day[?]*" Then Item = WordsBetween("what", "were", InMessage): GoTo Reply

Exit Sub


Reply:
If Item Like "month" Then GoTo GetItem2
If Item Like "year" Then GoTo GetItem2
If Item Like "date" Then GoTo GetItem2
If Item Like "day" Then GoTo GetItem2
Exit Sub

GetItem2:
For DaysOff = 1 To 8
TheDate = DateAdd("d", -DaysOff, Date)
DayName = Format(TheDate, "dddd")
If InMessage Like "*" + DayName + "*" Then GoTo ShowIt
Next DaysOff
Exit Sub

ShowIt:
If Item Like "day" Then OutMessage = Format(TheDate, "dddd") + "."
If Item Like "month" Then OutMessage = Format(TheDate, "mmmm") + "."
If Item Like "year" Then OutMessage = Format(TheDate, "yyyy") + "."
If Item Like "date" Then OutMessage = Format(TheDate, "dddd, mmmm d yyyy") + "."
Exit Sub


End Sub



' Answers questions like:
' What month will it be next Wednesday?
' What date will it be next Friday?
Private Sub WhatDateNextXday()
If InMessage Like "*What * will it be next *day[?]*" Then GoTo Possible
Exit Sub

Possible:
Item = WordsBetween("what", "will", InMessage)
If Item Like "month" Then GoTo GetItem2
If Item Like "year" Then GoTo GetItem2
If Item Like "date" Then GoTo GetItem2
If Item Like "day" Then GoTo GetItem2
Exit Sub


GetItem2:
DayNames = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
For I = 0 To 6
If InMessage Like "* next " + DayNames(I) + "*" Then Item2 = DayNames(I): GoTo GetDay
Next I
Exit Sub


GetDay:
If Item Like "day" Then OutMessage = Item2 + "."
C = 0
For N = 0 To 16
TheDate = DateSerial(Year(Date), Month(Date), Day(Date) + N)
If Format(TheDate, "dddd") Like Item2 Then C = C + 1: If C = 2 Then GoTo ShowIt
Next N
Exit Sub


ShowIt:
If Item Like "month" Then OutMessage = Format(TheDate, "mmmm") + "."
If Item Like "year" Then OutMessage = Format(TheDate, "yyyy") + "."
If Item Like "date" Then OutMessage = Format(TheDate, "dddd, mmmm d yyyy") + "."

End Sub


' This can answer questions like
' What date will it be in 3 days?.
' What date will it be in 3 hours and 30 minutes?
' What date will it be tomorrow?
Private Sub WhatDateX()
If InMessage Like "*What date *[?]*" Then GoTo Possible
Exit Sub



Possible:
NewMessage = WordsBetween("what date", "?", InMessage)
If NewMessage Like "*tomorrow*" Then GoTo GetIt
If NewMessage Like "*yesterday*" Then GoTo GetIt
If NewMessage Like "*fort*night*" Then GoTo GetIt
If NewMessage Like "* from now*" Then NewMessage = WordsBetween("", "from now", NewMessage): GoTo CheckMore
If NewMessage Like "will it be next *" Then NewMessage = WordsBetween("be", "", NewMessage): GoTo CheckMore
If NewMessage Like "will it be in *" Then NewMessage = WordsBetween("be in", "", NewMessage): GoTo CheckMore
If NewMessage Like "will the * be in *" Then NewMessage = WordsBetween("be in", "", NewMessage): GoTo CheckMore
If NewMessage Like "is it in *" Then NewMessage = WordsBetween("it in", "", NewMessage): GoTo CheckMore
If NewMessage Like "in *" Then NewMessage = WordsBetween("in", "", NewMessage): GoTo CheckMore
If NewMessage Like "was it * ago*" Then NewMessage = WordsBetween("was it", "ago", NewMessage): GoTo CheckMore
If NewMessage Like "was it last*" Then NewMessage = WordsBetween("last", "", NewMessage): GoTo CheckMore
If NewMessage Like "were it * ago*" Then NewMessage = WordsBetween("were it", "ago", NewMessage): GoTo CheckMore
Exit Sub




CheckMore:
Item = Array("year", "month", "week", "day", "hour", "minute", "second", "")
I = 0
Loop1:
If Item(I) = "" Then Exit Sub
If NewMessage Like "*" + Item(I) + "*" Then GoTo GetIt
I = I + 1
GoTo Loop1



CheckMonth:
For I = 1 To 12
MyDate = DateAdd("m", I, Date)
MonthName1 = Format(MyDate, "mmm")
MonthName2 = Format(MyDate, "mmmm")
If InMessage Like "*" + MonthName1 + "[!a-z]*" Then Call ReplaceWords(MonthName1, DateDiff("m", MyDate, Date) + " months", InMessage): GoTo GetIt
If InMessage Like "*" + MonthName2 + "[!a=z]*" Then Call ReplaceWords(MonthName2, DateDiff("m", MyDate, Date) + " months", InMessage): GoTo GetIt
Next I
Exit Sub



GetIt:
Call ReplaceWords("half hour", "30 minutes", InMessage)
Call ReplaceWords("half an hour", "30 minutes", InMessage)
Call NumbersToWords(InMessage)
Call OperatorWordsToSymbols(InMessage)
Item = "year": GoSub GetItemAmount: years = Amount
Item = "month": GoSub GetItemAmount: months = Amount
Item = "week": GoSub GetItemAmount: weeks = Amount
Item = "day": GoSub GetItemAmount: days = Amount
Item = "hour": GoSub GetItemAmount: Hours = Amount
Item = "minute": GoSub GetItemAmount: minutes = Amount
Item = "seconds": GoSub GetItemAmount: seconds = Amount
If InMessage Like "*tomorrow*" Then days = "1"
If InMessage Like "*yesterday*" Then days = "-1"
If InMessage Like "*fort*night*" Then days = "14"
If InMessage Like "* ago*" Then GoSub NegateTime
'WeeksNumber = Val(SuperSum(weeks))
'DaysNumber = Val(SuperSum(days))
WeeksNumber = Val(weeks)
DaysNumber = Val(days)
days = Str(DaysNumber + (WeeksNumber * 7))
TheDate = CalcDate(years, months, days, Hours, minutes, seconds)
OutMessage = TheDate + "."
Exit Sub




GetItemAmount:
If InMessage Like "* " + Item + "*" Then GoTo FindItem
Amount = ""
Return
FindItem:
EndPosition = InStr(1, InMessage, Item) - 1
GoSub FindItemNumberStart
Amount = Mid(InMessage, StartPosition, EndPosition - StartPosition)
If Amount Like "next" Then Amount = "1": Return
If Amount Like "last" Then Amount = "-1": Return
Amount = SuperSum(Amount)
Return



FindItemNumberStart:
StartPosition = EndPosition - 1
Loop2:
If StartPosition = 1 Then Return
C = Mid(InMessage, StartPosition, 1)
If C = " " Then StartPosition = StartPosition + 1: Return
If C = "," Then StartPosition = StartPosition + 1: Return
StartPosition = StartPosition - 1
GoTo Loop2


NegateTime:
years = "-" + years
months = "-" + months
days = "-" + days
weeks = "-" + weeks
Hours = "-" + Hours
minutes = "-" + minutes
seconds = "-" + seconds
Return





End Sub



'
Private Sub WhatDayIsIt()

If InMessage Like "*What day is it[?]*" Then GoTo GetIt
If InMessage Like "*What day it is[?]*" Then GoTo GetIt
If InMessage Like "*What is the day today[?]*" Then GoTo GetIt
If InMessage Like "*What day is it today[?]*" Then GoTo GetIt
If InMessage Like "The Day[?]*" Then GoTo GetIt
If InMessage Like "The day today[?]*" Then GoTo GetIt
If InMessage Like "Day[?]*" Then GoTo GetIt
If InMessage Like "Day." Then GoTo GetIt
Exit Sub

GetIt:
OutMessage = "It is " + Format(Date, "dddd") + "."
End Sub


' What day will it be in 2 years and 3 months?
' What day will it be tomorrow?
' What year was it yesterday?
' What will the date be in 300 minutes?
' What month was it 3 months ago?
' What day will it be in 3 days and 30 minutes?
' What time will it be in 20 minutes?
Private Sub WhatDateDate()
OldMessage = InMessage
If InMessage Like "Yesterday[?]*" Then Item = "Day": InMessage = "What day was it yesterday?": GoTo Likely
If InMessage Like "*Tomorrow[?]*" Then Item = "Day": InMessage = "What day will it be tomorrow?": GoTo Likely
If InMessage Like "*What * will it be*" Then Item = WordsBetween("What", "will", InMessage): GoTo Likely
If InMessage Like "*What * will the * be *" Then Item = WordsBetween("What", "will", InMessage): GoTo Likely
If InMessage Like "*What * was it *" Then Item = WordsBetween("What", "was", InMessage): GoTo Likely
If InMessage Like "*What * were it *" Then Item = WordsBetween("What", "were", InMessage): GoTo Likely
If InMessage Like "*What * is it *" Then Item = WordsBetween("What", "is", InMessage): GoTo Likely
If InMessage Like "*What will the * be *" Then Item = WordsBetween("the", "be", InMessage): Call ReplaceWords("will the " + Item, Item + " will it", InMessage): GoTo Likely
If InMessage Like "*What is the * in *" Then Item = WordsBetween("the", "in", InMessage): Call ReplaceWords("is the " + Item, Item + " is it", InMessage): GoTo Likely
If InMessage Like "*What is the * from now*" Then Item = WordsBetween("the ", " ", InMessage): Call ReplaceWords("is the " + Item, Item + " is it", InMessage): GoTo Likely
Exit Sub




Likely:
If InMessage Like "* next *day*" Then Exit Sub

If Item Like "year" Then GoTo GetIt
If Item Like "month" Then GoTo GetIt
If Item Like "day" Then GoTo GetIt
If Item Like "time" Then GoTo GetIt
If Item Like "date" Then GoTo GetIt
InMessage = OldMessage
Exit Sub


GetIt:
Call ReplaceCharacters(Item, "date", InMessage)
Call WhatDateX
If OutMessage Like "*Error*" Then OutMessage = "Sorry, I can't calculate dates that far.": Exit Sub
If OutMessage = "" Then InMessage = OldMessage: Exit Sub
OutMessage = WordsBetween(Item + ":", ".", OutMessage) + "."



End Sub




' This will answer questions like:
' What day was it on the 1/10/66?
' What day will it be on the 2nd of January, 1999
Private Sub WhatDayOnDate()
If InMessage Like "*What day * on the*" Then Date_Part = WordsBetween("on the", "?", InMessage): GoTo GetIt
If InMessage Like "*What day * on *" Then Date_Part = WordsBetween("on", "?", InMessage): GoTo GetIt
Exit Sub

GetIt:
GoSub ExtractDate
TheDay = Format(TheDate, "dddd")
If TheDay Like "*day*" Then OutMessage = TheDay + "."
Exit Sub


ExtractDate:
MyDate = WordsToNumbers(Date_Part)
Call ConvertEventToDate(MyDate)
Call ReplaceWords("first", "1", MyDate)
Call ReplaceWords("second", "2", MyDate)

ExtractedDate = ""

Loop1:
MonthNumber = 0
Word = ExtractWord(MyDate)
If Word <> "" Then GoTo TestWord
TheDate = ExtractedDate

If IsDate(TheDate) Then Return
If IsDate("1 " + TheDate) Then TheDate = "1 " + TheDate: Return
Return

TestWord:
If Word Like "*[/-]*[/-]*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#:#*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*[a-z]day" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Val(Word) Then Number = Val(Word): ExtractedDate = ExtractedDate + Str(Number) + " ": GoTo Loop1

GetMonthNumber:
NewDate = DateAdd("m", MonthNumber, Date)
MonthName1 = Format(NewDate, "mmm")
MonthName2 = Format(NewDate, "mmmm")
If Word Like MonthName1 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like MonthName2 Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
MonthNumber = MonthNumber + 1: If MonthNumber = 12 Then GoTo Loop1
GoTo GetMonthNumber


End Sub


Private Sub WhatDidISay()
If InMessage Like "*What did i *say[?]*" Then GoTo Repeat
Exit Sub
Repeat:
If PreviousInMessage = "" Then OutMessage = "You didn't say anything.": Exit Sub
OutMessage = "You said:- " + "'" + PreviousInMessage + "'."
End Sub

Private Sub WhatDoesXEqual()
If InMessage Like "*what does * equal[?]*" Then
    Call ChangeToPerson(InMessage)
    Call ChangeToSubject(InMessage)
    Call ChangeYourForMyEtc(InMessage)
    Item = WordsBetween("what does", "equal", InMessage)
    RestartItemSearch
    ItemSearchMethod = WordForWord
    ItemSearchStyle = Strict

Loop1:
    Statement = NextItemContaining(Item)
    If Statement = "" Then OutMessage = "I dont know what " + Item + " equals.": Exit Sub
    If Statement Like "*[?]*" Then GoTo Loop1
    If Statement Like "*" + Item + " equals*" Then GoTo Found
    If Statement Like "*" + Item + "=*" Then GoTo Found
    GoTo Loop1

Found:
    OutMessage = Statement
End If

End Sub


Private Sub WhatDoYouMean()
If InMessage Like "*What do you mean*" Then OutMessage = "I mean what I say.": Exit Sub
If InMessage Like "*What does that mean*" Then GoTo Reply
If InMessage Like "*What is the meaning of that*" Then GoTo Reply
If InMessage Like "*What is that supposed to mean*" Then GoTo Reply
If InMessage Like "*What is that meant to mean*" Then GoTo Reply
If InMessage Like "Meaning[?]*" Then GoTo Reply
Exit Sub

Reply:
Dim Msg(6)
Msg(1) = "It means what it says."
Msg(2) = "It will probably mean what it says."
Msg(3) = "Figure it out."
Msg(4) = "You'll work it out."
Msg(5) = "It means what it means."
MsgNumber = Int((Rnd * 5) + 1)
OutMessage = Msg(MsgNumber)



End Sub

Private Sub WhatIsCalc()
If InMessage Like "Who *" Then Exit Sub
If InMessage Like "Where *" Then Exit Sub
If InMessage Like "Why *" Then Exit Sub
If InMessage Like "When *" Then Exit Sub
If InMessage Like "*#/#/#*" Then Exit Sub
If InMessage Like "*#/##/#*" Then Exit Sub
NewMessage = WordsToNumbers(InMessage)
If NewMessage Like "#*#." Then GoTo CheckMore
If NewMessage Like "*[?]*" Then GoTo CheckMore
Exit Sub



CheckMore:
GoSub CheckObvious: If OutMessage <> "" Then Exit Sub
GoSub CheckOther: If OutMessage <> "" Then Exit Sub
GoSub CheckFraction: If OutMessage <> "" Then Exit Sub
Exit Sub




CheckObvious:
If NewMessage Like "what is *# added to #*" Then Call ReplaceWords("added to", "add", NewMessage): GoSub Maybe: Return
If InMessage Like "* add *" Then GoSub Maybe: Return
If InMessage Like "* plus *" Then GoSub Maybe: Return
If InMessage Like "* minus *" Then GoSub Maybe: Return
If InMessage Like "* take away *" Then GoSub Maybe: Return
If InMessage Like "* multiplied by*" Then GoSub Maybe: Return
If InMessage Like "*[!a-z]Times[!a-z]*" Then GoSub Maybe: Return
If InMessage Like "*Divided by*" Then GoSub Maybe: Return
If InMessage Like "*[-/+*]*" Then GoSub Maybe: Return
Return




CheckOther:
If " " + InMessage Like "* How many * within *" Then GoSub Maybe: Return
If " " + InMessage Like "* How many * in *" Then GoSub Maybe: Return
If InMessage Like "What * the amount * in *" Then GoSub Maybe: Return
If InMessage Like "Tell me *the number * in *" Then GoSub Maybe: Return
If InMessage Like "What is the amount * in *" Then GoSub Maybe: Return
If InMessage Like "What is the number * in *" Then GoSub Maybe: Return
If InMessage Like "*What is the sum of*" Then GoSub Maybe: Return
If InMessage Like "*Tell *the sum of*" Then GoSub Maybe: Return
If InMessage Like "calculate *" Then GoSub Maybe: Return
If NewMessage Like "*# #*" Then GoSub Maybe: Return
If NewMessage Like "*# into #*" Then GoSub Maybe: Return
Return



CheckFraction:
If NewMessage Like "* of *" Then GoTo CheckFractionPart
Return
CheckFractionPart:
If NewMessage Like "*# percent of *" Then GoSub Maybe: Return
If NewMessage Like "*#% of *" Then GoSub Maybe: Return
If NewMessage Like "*half of *" Then GoSub Maybe: Return
If NewMessage Like "*quarter of *" Then GoSub Maybe: Return
If NewMessage Like "*#rd of *" Then GoSub Maybe: Return
If NewMessage Like "*#th of *" Then GoSub Maybe: Return
Return






Maybe:
If NewMessage Like "* within a *" Then Call ReplaceWords("within a", "in", NewMessage)
If NewMessage Like "* in a *" Then Call ReplaceWords("in a", "in", NewMessage)
Result = SuperSum(NewMessage)
If Result = "" Then OutMessage = "": Return
If Result Like "*Divide by zero*" Then OutMessage = "Thou shalt not divide by zero!": Return
OutMessage = Result + "."
Return


End Sub


Private Sub Convert()
If InMessage Like "*Convert*#*" Then
Number = WordsBetween("Convert", "", InMessage)
ConvertedNumber = NumbersToWords(Number)
OutMessage = "It is: " + ConvertedNumber + "."
End If
End Sub

Private Sub WhatIsMyName()
If InMessage Like "*my name*[?]*" Then GoTo SayName
Exit Sub

SayName:
    OutMessage = "Your name is " + UsersName + "."
End Sub


Private Sub WhatIsTheDate()
If InMessage Like "*What is the date[?]*" Then GoTo GetIt
If InMessage Like "Date[?]*" Then GoTo GetIt
If InMessage Like "*The date[?]*" Then GoTo GetIt
If InMessage Like "*todays date[?]*" Then GoTo GetIt
If InMessage Like "*today's date[?]*" Then GoTo GetIt
If InMessage Like "*What todays date is[?]*" Then GoTo GetIt
If InMessage Like "*What is the date today[?]*" Then GoTo GetIt
If InMessage Like "Date." Then GoTo GetIt
If InMessage Like "Say the date[!a-z]*" Then GoTo GetIt
Exit Sub

GetIt:
Tday = Format(Date, "d")
TMonth = Format(Date, "mmmm")
Tday = Date2to2ndEtc(Tday)



GotIt:
OutMessage = "The " + Tday + " of " + TMonth + "."


End Sub

Private Sub WhatIsTheSubject()
MySubject = Subject


If InMessage Like "*What * subject*" Then
Temp = Subject
Call GetSubject(BotsMessage)
MySubject = Subject
Subject = Temp
If MySubject = "" Then OutMessage = "I dont know the subject.": Exit Sub
If MySubject = "I" Then MySubject = "me"
OutMessage = "The subject is " + MySubject + "."
GoTo Done
End If



If InMessage Like "*What are we talking about*" Then
Temp = Subject
Call GetSubject(PreviousInMessage)
MySubject = Subject
Subject = Temp
Call ChangeYourForMyEtc(MySubject)
If MySubject = "" Then OutMessage = "I dont know what we're talking about.": Exit Sub
OutMessage = "We're talking about " + MySubject + "."
GoTo Done
End If



If InMessage Like "*what am I talking about*" Then
Temp = Subject
Call GetSubject(PreviousInMessage)
MySubject = Subject
Subject = Temp
If MySubject = "" Then OutMessage = "I dont know what you're talking about.": Exit Sub
Call ChangeYourForMyEtc(MySubject)
OutMessage = "You're talking about " + MySubject + "."
GoTo Done
End If



If InMessage Like "*what are you talking about*" Then
GoTo GetLastSubject
End If

If InMessage Like "*what* you on about*" Then
GoTo GetLastSubject
End If


Done:
Call ChangeHisForYourEtc
Exit Sub



GetLastSubject:
Temp = Subject
Call GetSubject(BotsMessage)
MySubject = Subject
Subject = Temp
If MySubject = "" Then OutMessage = "I dont know what I am talking about.": Exit Sub
If MySubject = "I" Then MySubject = "me"
OutMessage = "I am talking about " + MySubject + "."
GoTo Done



End Sub

Private Sub WhatMonthIsIt()
If InMessage Like "*What *month*in[?]*" Then GoTo GetIt
If InMessage Like "* month is it[?]*" Then GoTo GetIt
If InMessage Like "* month it is[?]*" Then GoTo GetIt
If InMessage Like "*me the month." Then GoTo GetIt
If " " + InMessage Like "*[!a-z]month is[?]*" Then GoTo GetIt
If InMessage Like "*month[?]*" Then GoTo GetIt
If InMessage Like "month." Then GoTo GetIt
Exit Sub
GetIt:
OutMessage = Format(Date, "mmmm") + "."
End Sub


Private Sub WhatTimeIsIt()
Message = InMessage
Call RemoveUnnecessaryWords(Message)
If Message Like "*What time is it[?]*" Then GoTo GetIt
If Message Like "*Time[?]*" Then GoTo GetIt
If Message Like "*Time it is[?]*" Then GoTo GetIt
If Message Like "*Time of day*[?]*" Then GoTo GetIt
If Message Like "tell me the time." Then GoTo GetIt
If Message Like "give me the time." Then GoTo GetIt
If Message Like "Time." Then GoTo GetIt
If Message Like "Say the time[!a-z]*" Then GoTo GetIt
Exit Sub

GetIt:
    OutMessage = "The time is " + Format(Time, "H:mm am/pm") + "."
End Sub




Private Sub WhatYearIsIt()
If InMessage Like "What year[?]*" Then GoTo GetIt
If InMessage Like "*Year is it[?]*" Then GoTo GetIt
If InMessage Like "*year are we in[?]*" Then GoTo GetIt
If InMessage Like "year[?]*" Then GoTo GetIt
If InMessage Like "*the year[?]*" Then GoTo GetIt
If InMessage Like "Year." Then GoTo GetIt
Exit Sub

GetIt:
OutMessage = "The year is " + Format(Date, "yyyy") + "."
End Sub


' This will replace the Users name with 'your'.
' It will also switch other words within the chatbots reply.
Public Sub ChangeHisForYourEtc()
    Call ReplaceWords("I is", "I am", OutMessage)   ' Added this line on the 5/8/2001
    If OutMessage Like "*" + UsersName + "*" Then
    ChangeAll UsersName + "'s", "your", OutMessage
    ReplaceWords UsersName + " is", "you are", OutMessage
    ReplaceWords UsersName, "you", OutMessage
    Call ReplaceWords("you has", "you have", OutMessage)
'    ReplaceWords "his", "your", OutMessage
    End If
End Sub







' This lets the user add new word to the list of alternative words (synonyms)
' Below are some example statements the user could type:

' "Killed is another word for murdered."
' "Cat=Feline,moggy,fleabag."
' "Happy means the same as pleased."
Private Sub XMeansY()
If InMessage Like "*[?]*" Then Exit Sub
Statement = InMessage

OneWayMode = False

If Statement Like "* means the same as *" Then GoTo GetMeansV2
If Statement Like "* is the same as *" Then GoTo GetSameAs
If Statement Like "* means *" Then GoTo GetMeans
If Statement Like "*[a-z]=[a-z]*" Then GoTo GetEquals
If Statement Like "* is short for *" Then GoTo GetShortFor
'If Statement Like "*[a-z]=[*]*" Then GoTo GetEquals
If Statement Like "* equals [a-z]*" Then GoTo GetEqualsV2
If Statement Like "* is another word for *" Then GoTo IsFor
If Statement Like "*another word for * is *" Then GoTo ForIs
If Statement Like "* is *synonym for *" Then GoTo IsFor
If Statement Like "* is *acronym for *" Then GoTo IsFor
If Statement Like "The acronym for * is *" Then GoTo ForIs
If Statement Like "* is *alternative word for *" Then GoTo IsFor
If Statement Like "* has the same meaning as *" Then GoTo HasAs
If Statement Like "The synonym for * is *" Then GoTo ForIs
If Statement Like "* is called *" Then GoTo IsCalled
If Statement Like "* the past of *" Then GoTo PastOf
'If Statement Like "* is a *" Then GoTo IsA
'If Statement Like "* is an *" Then GoTo IsAn
'If Statement Like "* name is *" Then GoTo NameIs
Exit Sub



GetBracks:
'Call AddMatchWords(Statement)
Exit Sub



PastOf:
If Statement Like "* is the past of *" Then GoTo IsPastOf
If Statement Like "* are the past of *" Then GoTo ArePastOf
Exit Sub


ArePastOf:
Word2 = WordsBetween("", "are the past of", Statement)
Word1 = WordsBetween("are the past of", ".", Statement)
GoTo DoIt


IsPastOf:
Word2 = WordsBetween("", "is the past of", Statement)
Word1 = WordsBetween("is the past of", ".", Statement)
GoTo DoIt


GetShortFor:
Word2 = WordsBetween("", "is short", Statement)
If Word2 Like "* *" Then Exit Sub
Word1 = WordsBetween("short for", ".", Statement)
If Word1 Like "* *" Then Exit Sub
GoTo DoIt



NameIs:
Word1 = WordsBetween("", "name is", Statement)
Word2 = WordsBetween("name is", ".", Statement)
If Word1 Like "*'s" Then Word1 = Left(Word1, Len(Word1) - 2)
GoTo DoType2


IsCalled:
Word1 = WordsBetween("", "is called", Statement)
Word2 = WordsBetween("is called", ".", Statement)
GoTo DoType2

IsA:
Word1 = WordsBetween("", "is a", Statement)
Word2 = WordsBetween("is a", ".", Statement)
GoTo DoType2

IsAn:
Word1 = WordsBetween("", "is an", Statement)
Word2 = WordsBetween("is an", ".", Statement)
GoTo DoType2


HasAs:
Word1 = WordsBetween("", "has", Statement)
Word2 = WordsBetween(" as ", ".", Statement)
GoTo DoIt

GetSameAs:
Word1 = WordsBetween("", "is", Statement)
Word2 = WordsBetween("same as", ".", Statement)
GoTo DoIt

ForIs:
Word1 = WordsBetween("for", "is", Statement)
Word2 = WordsBetween("is", ".", Statement)
GoTo DoIt

IsFor:
Word1 = WordsBetween("", "is", Statement)
Word2 = WordsBetween("for", ".", Statement)
GoTo DoIt

GetMeans:
If Statement Like "*" + Chr(34) + "*" + Chr(34) + "*" Then GoTo GM
If Statement Like "*'*' means *" Then GoTo GM
If Statement Like "* means '*'*" Then GoTo GM
If Statement Like "* * means *" Then Exit Sub
If Statement Like "* means * *" Then Exit Sub
GM:
Word1 = WordsBetween("", "means", Statement)
Word2 = WordsBetween("means", ".", Statement)
GoTo DoIt



GetMeansV2:
Word1 = WordsBetween("", "means", Statement)
Word2 = WordsBetween("same as", ".", Statement)
GoTo DoIt



GetEquals:
OneWayMode = True
'If InMessage Like "*now and then*" Then Stop
Word1 = WordsBetween("", "=", Statement)
Word2 = WordsBetween("=", ".", Statement)
GoTo DoIt

GetEqualsV2:
OneWayMode = True
Word1 = WordsBetween("", "equals", Statement)
Word2 = WordsBetween("equals", ".", Statement)




DoIt:
'Call ChangeToPerson(Word1)
'Call ChangeToSubject(Word1)
'Call ChangeYourForMyEtc(Word1)
'Call ChangeToPerson(Word2)
'Call ChangeToSubject(Word2)
'Call ChangeYourForMyEtc(Word2)


If Word2 Like "'*'" Then Call AddMatchWord(Word1, Word2): Exit Sub
If Word2 Like "* and *" Then GoTo GetWords
If Word2 Like "* or *" Then GoTo GetWords
If Word2 Like "*,*" Then GoTo GetWords


Call AddMatchWord(Word1, Word2)
If OneWayMode Then Exit Sub
Call AddMatchWord(Word2, Word1)
Exit Sub





GetWords:
Word = Word1: GoSub RemoveBits: Word1 = Word
Word = Word2: GoSub RemoveBits: Word2 = Word

Dim Words(100)
'Call ExtractWords(Word2, Words())
GoSub ExtractEm

N = 0
Do
If Words(N) = "" Then GoTo Done
If Words(N) Like "[!a-z]*" Then GoTo Skip
Call AddMatchWord(Word1, Words(N))
If OneWayMode = False Then Call AddMatchWord(Words(N), Word1)

Skip:
N = N + 1
Loop

Done:
Exit Sub
 




ExtractEm:
p = 0
N = 0
EELoop1:
If p >= Len(Word2) Then Return
P2 = InStr(p + 1, Word2, ",")
If P2 = 0 Then P2 = Len(Word2) + 1
Word = Trim(Mid(Word2, p + 1, P2 - (p + 1)))
If P2 = Len(Word2) + 1 Then If Word Like "* and *" Then GoTo SplitAnd
If P2 = Len(Word2) + 1 Then If Word Like "* or *" Then GoTo SplitOr
Words(N) = Word
p = P2
N = N + 1: If N = 47 Then Return
GoTo EELoop1


' No longer used
SplitAnd:
p = InStr(1, Word, " and ")
Words(N) = Mid(Word, 1, p - 1)
Words(N + 1) = Mid(Word, p + 5)
Return

SplitOr:
p = InStr(1, Word, " or ")
Words(N) = Mid(Word, 1, p - 1)
Words(N + 1) = Mid(Word, p + 4)
Return





DoType2:
Call ChangeToSubject(Word1)
Call ChangeToPerson(Word1)
Call ChangeYourForMyEtc(Word1)
Call ChangeToSubject(Word2)
Call ChangeToPerson(Word2)
Call ChangeYourForMyEtc(Word2)



Word = Word1: GoSub RemoveBits: Word1 = Word
Word = Word2: GoSub RemoveBits: Word2 = Word
Call AddMatchWord(Word1, Word2)
If OneWayMode = False Then Call AddMatchWord(Word2, Word1)
Exit Sub



RemoveBits:
If Word Like "[!a-z]*[!a-z]" Then Word = Mid(Word, 2, Len(Word) - 2)
If Word Like "A *" Then Word = Right(Word, Len(Word) - 2)
If Word Like "An *" Then Word = Right(Word, Len(Word) - 3)
If Word Like "The *" Then Word = Right(Word, Len(Word) - 4)
Return

End Sub


Private Sub XisY()
If InMessage Like "The * is *" Then GoTo CheckFurther
If InMessage Like "An * is *" Then GoTo CheckFurther
If InMessage Like "A * is *" Then GoTo CheckFurther
Exit Sub

CheckFurther:
If InMessage Like "* is not[!a-z]*" Then Exit Sub
Statement = InMessage

Call ChangeToPerson(Statement)
Call ChangeToSubject(Statement)
Call ChangeYourForMyEtc(Statement)

If Statement Like "* is a *" Then GoTo IsA
If Statement Like "* is an *" Then GoTo IsAn

Exit Sub


NameIs:
Word1 = WordsBetween("", "name is", Statement)
Word2 = WordsBetween("name is", ".", Statement)
If Word1 Like "*'s" Then Word1 = Left(Word1, Len(Word1) - 2)
GoTo DoType2

IsA:
Word1 = WordsBetween("", "is a", Statement)
Word2 = WordsBetween("is a", ".", Statement)
GoTo DoType2

IsAn:
Word1 = WordsBetween("", "is an", Statement)
Word2 = WordsBetween("is an", ".", Statement)


DoType2:
Word = Word1: GoSub RemoveBits: Word1 = Word
Word = Word2: GoSub RemoveBits: Word2 = Word

Call AddMatchWord(Word1, Word2)
Exit Sub


RemoveBits:
If Word Like "[!a-z]*[!a-z]" Then Word = Mid(Word, 2, Len(Word) - 2)
If Word Like "A *" Then Word = Right(Word, Len(Word) - 2)
If Word Like "An *" Then Word = Right(Word, Len(Word) - 3)
If Word Like "The *" Then Word = Right(Word, Len(Word) - 4)
Return




End Sub


