Attribute VB_Name = "VariousFunctions"
Public UsersMessage
Public UsersMessageTime
Public BotsMessage
Public BotsMessageTime
Public InternalMessage
Public InternalMessageTime

Public BotsMessagePartTime  ' The time of the message part
Public BotsMessagePart      ' The bots message is read/typed out in parts. This contains the current part.


Public TimeOfLastUserActivity
Public TimeOfLastBotActivity

Public SwearingAllowed
Public QuietMode
Public RamblingAllowed
Public InferringAllowed

Public WontSay

Public MultiAnswerMode

Public ReadingAllowed
Public BotReading
Public BotReadingPosition
Public BotReadingFile       ' Name of the factfile that the bot is reading
Public BotReadingFullpath   ' Full path of the factfile that the bot is reading

Public SaveMemoryAllowed

Public UsersName
Public Subject
Public BareSubject 'This is the subject without the 'an,a,the' words attached.
Public ItemSearchMethod
Public Const InOrder = 2
Public Const AnyOrder = 1
Public Const WordForWord = 0
Public ItemSearchStyle
Public Const Careless = 1
Public Const Strict = 0

Public WordsData

Public AllWords()
Public AllWordsAmount

Public Script

Public ShutdownTime

Public MemoryPosition
Public Memory

Public GeneralKnowledge
Public PersonalKnowledge
Public Facts


Public Const VarBufferSize = 50
Public VarName(VarBufferSize)
Public VarValue(VarBufferSize)

Type CategoriesStructure
WordsAmount(100) As Integer
Words(100, 100) As String
End Type
Public CurrentCategories As CategoriesStructure
Public Categories


Public Things


Public EventsAmount
Public EventItem(32)
Public EventRemindInterval(32)
Public EventTime(32)
Public EventRemindStartTime(32)
Public EventLastMessageTime(32)



' Data related to the 'FindRelevantStatements' routine
Public Statements(100)  ' The found statements
Public Relevances(100)  ' The sorted relevances of the found statements
Public Indexes(100)     ' Sorted indexes to the found statements (Highest relevance first)
Public MatchedWords(100, 20) ' All the words that had been matched (for debug purposes)

Public RelevanceAmount  ' The actual relevance the most relevant statement has.

' These constants determine the search mode.
Public Const MostRelevant = 0   ' Predicts the ideal relevance a statement needs to have and halts searching if found. (This can help to speed things up a little when searching for a single response to a query)
Public Const AllRelevant = 1    ' Returns all relevant statements


Public RecNumber       ' The number of relevant statements found


Public MatchWordsAmount
Public MatchWords






Public NewMessageAvailable


Public Const Past = -1
Public Const Present = 0
Public Const Future = 1

Public Const MessageHistorySize = 5000
Public MessageHistory

Dim AltWordsV2(1010)
Private PreviousAltWord


Type AlternativesStructure
WordNumber(100) As Integer
WordsAmount(100) As Integer
Words(30, 170) As String
End Type
Public CurrentAlternatives As AlternativesStructure
Public AlternativeWords


Public WontSayData


Option Compare Text

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Type POINTAPI
    X As Long
    Y As Long
    End Type

 Public Sub DontSayMessage()

If BotsMessage Like "." Then Exit Sub
WontSayData = WontSayData + BotsMessage

Dim Msg(15)
Static MsgNumber
Msg(0) = "Ok, I will not say it again."
Msg(1) = "I will not say that."
Msg(2) = "I wont say that."
Msg(3) = "Ok, no problem."
Msg(4) = "I wont."
Msg(5) = "Right."
Msg(6) = "Okay."
Msg(7) = "Forgotten already!."
Msg(8) = "Is there anything I can say?."
Msg(9) = "Would I ever?."
Msg(10) = "As if!."
Msg(11) = "Promise."
Msg(12) = "Okay."
Msg(13) = "Right."

If MsgNumber < 12 Then MsgNumber = MsgNumber + 1 Else MsgNumber = 12 + Int(Rnd * 2)

OutMessage = Msg(MsgNumber)
Call SetBotsMessage(Msg(MsgNumber))



End Sub


' Deletes the item from memory at <Position>
Public Function DeleteItem(MemPos)

Debug.Print "Memory position:- "; MemoryPosition
Debug.Print "Memory:-"; Mid(Memory, MemoryPosition, 20)

GetEnd:
ItemEnd = InStr(MemPos, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If ItemEnd <= 1 Then DeleteItem = "": Exit Function


GetStart:
ItemStart = InStrRev(Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), MemPos, vbBinaryCompare)
If ItemStart = 0 Then ItemStart = 1: GoTo GotIt
ItemStart = ItemStart + 4



GotIt:
MemPos = ItemStart
DeleteItem = Mid(Memory, ItemStart, ItemEnd - ItemStart)


Memory = Left(Memory, MemPos - 1) + Right(Memory, Len(Memory) - (MemPos + Len(DeleteItem)) - 3)

MemoryPosition = ItemStart



End Function

Public Sub FindRelevantStatements(Message, SearchMode)
    Dim DebugWords(20)

    TryNumber = 1

    OriginalMessage = Message

    RecNumber = 0

GetSubjectEtc:
    Call ChangeToPerson(Message)
    Call ChangeToSubject(Message)
    Call ChangeYourForMyEtc(Message)


    Item = Message



    If Message Like "*information about *" Then Item = WordsBetween(" about ", "", Message): GoTo SetTriv
    If Message Like "*information on *" Then Item = WordsBetween("information on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]data on *" Then Item = WordsBetween("data on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]low down on *" Then Item = WordsBetween("low down on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]run down on *" Then Item = WordsBetween("run down on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]dirt on *" Then Item = WordsBetween("dirt on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]facts on *" Then Item = WordsBetween("facts on", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]gossip on *" Then Item = WordsBetween("gossip on ", "", Message): GoTo SetTriv
    If Message Like "*knowledge of *" Then Item = WordsBetween("knowledge of", "", Message): GoTo SetTriv
    If " " + Message Like "*[!a-z]tell *[!a-z]about *" Then Item = WordsBetween("about", "", Message): GoTo SetTriv
    If Message Like "*[!a-z]know*[!a-z]about*" Then Item = WordsBetween("about", "", Message): GoTo SetTriv
    
    
    If Message Like "do i know *" Then Call ReplaceWords("do I know", "", Item)


SetTriv:
    TrivialWords = Array("is", "do", "does", "what", "why", "which", "just", "who", "by", "of", "it", "a", "and", "an", "as", "be", "the", "there", "that", "to", "was", "are", "not", "did", "had", "for", "this", "am", "would", "can", "could", "will", "wont", "may", "might", "been", "are", "should", "very", "")


    If Item = "" Then Exit Sub

    CheckName = False
    If " " + Item Like "*[!a-z]" + UsersName + "[!a-z]*" Then CheckName = True

    
    Call GetMatchWords(Item)

 
    Debug.Print " "
    Debug.Print Message
    Debug.Print "Subject:- " + Subject

    Message = LCase(Message)



' This will count the amount of good juicy words there are within the users message
' If there aren't any then the message will be added to a previous message until there is a decent query

    N = 0
    GoodWordCount = 1
GELoop1:
    If CurrentAlternatives.Words(N, 0) = "" Then GoTo CheckGoodWordCount
    W = 0
GELoop2:
    If TrivialWords(W) = "" Then GoTo GoodWord
    If CurrentAlternatives.Words(N, 0) Like TrivialWords(W) Then GoodWordCount = GoodWordCount + 0.2: GoTo GELoop3
    W = W + 1
    GoTo GELoop2
GoodWord:
    GoodWordCount = GoodWordCount + 1
GELoop3:
    N = N + 1
    GoTo GELoop1



CheckGoodWordCount:
    If GoodWordCount >= 1 Then GoTo DoMonsterSearch

    If TryNumber = 4 Then GoTo DoMonsterSearch

    TryNumber = TryNumber + 1

    If TryNumber = 2 Then Message = BotsMessage + Message: GoTo GetSubjectEtc
    If TryNumber = 3 Then Message = LastInMessage + Message: GoTo GetSubjectEtc

 
DoMonsterSearch:
    GoSub GetSubjectAlternativeNumber


    N = 0
    RestartItemSearch


Loop0:
    DoEvents: If NewMessageAvailable Then GoTo Done
    N = 0
    Statement = PreviousMemoryItem()
    If Statement = "" Then GoTo SortRelevantStatements
    
    If Statement Like "*=*" Then GoTo Loop0

CheckQuest:
    If Statement Like "*[?]." Then GoTo Loop0
    If Statement Like Message Then GoTo Loop0
    

    OriginalStatement = Statement
    If Statement Like "If * then *" Then GoSub GetCondition
 


    DW = 0
    MatchCount = 0
    WordsAmount = 0
    Relevance = 0

    Statement2 = LCase(" " + Statement + " ")

'    If Statement Like "*tom wants the*" Then If InMessage Like "* want[?]*" Then Stop


Loop2:
    If CurrentAlternatives.Words(N, 0) = "" Then GoTo CheckOtherWords
    For W = 0 To CurrentAlternatives.WordsAmount(N) - 1
    CurrentWord = CurrentAlternatives.Words(N, W)
 
'    If " " + Statement Like "*[!a-z]" + CurrentWord + "[!a-z]*" Then GoTo SetMatchCount
    P = InStr(1, Statement2, CurrentWord, vbBinaryCompare)
    If P <> 0 Then GoTo CheckSides


GetNextVariation:
    Next W


GetNextWord:
    WordsAmount = WordsAmount + 1
    N = N + 1
    GoTo Loop2


CheckSides:
    Value = 1
    If CurrentWord = "=" Then GoTo SetMatchCount
    If Mid(Statement2, P - 1, 1) Like "[a-z]" Then GoTo GetNextVariation
    If Mid(Statement2, P + Len(CurrentWord), 1) Like "[!a-z]" Then GoTo SetMatchCount
    If Len(CurrentWord) > 4 Then GoTo SetMatchCount
    If Mid(Statement2, P + Len(CurrentWord), 1) Like "[a-z]" Then Value = Len(CurrentWord) / 10



SetMatchCount:
    If W = 0 Then Dword = "CA: " + CurrentWord: GoSub AddDebugWord Else Dword = "CA: " + CurrentAlternatives.Words(N, 0) + "(" + CurrentWord + ")": GoSub AddDebugWord
    If W > 0 Then Value = Value - 0.1
    If CurrentWord Like "what" Then GoTo GetNextWord
    TW = 0
    Do
    If TrivialWords(TW) = "" Then GoTo AddIt
    If CurrentWord Like TrivialWords(TW) Then Value = 0.2: GoTo AddIt
    TW = TW + 1
    Loop
AddIt:
    MatchCount = MatchCount + Value
    GoTo GetNextWord



CheckOtherWords:

CheckUserName:
    If CheckName Then If " " + Statement Like "*[!a-z]you*" Then Dword = "Name-check: You (" + UsersName + ")": GoSub AddDebugWord: MatchCount = MatchCount + 0.7


SetRelevance:
'    If MatchCount < Int(RecNumber / 20) + 0.9 Then GoTo Loop0
    
    If MatchCount < 1 Then GoTo Loop0


    If MatchCount > WordsAmount Then MatchCount = WordsAmount
    

    
    Relevance = (100 / GoodWordCount) * MatchCount



    SizeDifference = Len(Statement) - Len(Item)
'    If sizedifference < 0 Then sizedifference = (Not sizedifference) + 1
    If SizeDifference < 0 Then SizeDifference = 0
    Relevance = Relevance - (SizeDifference / 200)
    If Relevance < 0 Then GoTo Loop0



'    If Subject Like "me" Then GoTo RecordStatementAndRelevance
'    If Subject Like "I" Then GoTo RecordStatementAndRelevance
'   If Subject Like UsersName Then If " " + Statement Like "*[!a-z]you[!a-z]*" Then GoTo CheckStatementSubject


    
    If Subject = "" Then GoTo RecordStatementAndRelevance


    If " " + Statement Like "*[!a-z]" + Subject + "[!a-z]*" Then Dword = "Contains subject": GoSub AddDebugWord: Relevance = Relevance + 20: GoTo ContainsSubject

    If SubjectAlternativeNumber = -1 Then GoTo RecordStatementAndRelevance
    N = SubjectAlternativeNumber
    For W = 0 To CurrentAlternatives.WordsAmount(N) - 1
    If Statement Like "*[!a-z]" + CurrentAlternatives.Words(N, W) + "[!a-z]*" Then Dword = "Contains subject indirectly": GoSub AddDebugWord: Relevance = Relevance + 9: GoTo ContainsSubject
    If Statement Like "*[!a-z]" + CurrentAlternatives.Words(N, W) + "*" Then Dword = "Contains word similar to subject": GoSub AddDebugWord: Relevance = Relevance + (3 + Len(CurrentAlternatives.Words(N, W))): GoTo RecordStatementAndRelevance
    Next W

    GoTo RecordStatementAndRelevance

ContainsSubject:
    Temp2 = Subject
    Temp = BareSubject
    Call GetSubject(Statement)
    StatementBareSubject = LCase(BareSubject)
    StatementSubject = Subject
    BareSubject = Temp
    Subject = Temp2
    If StatementBareSubject = "" Then GoTo RecordStatementAndRelevance

    Dword = "Statement Subject:- " + StatementSubject: GoSub AddDebugWord

    If StatementSubject Like Subject Then Relevance = Relevance + 10: Dword = "Subjects Match": GoSub AddDebugWord: GoTo RecordStatementAndRelevance

CheckSub2:
    If SubjectAlternativeNumber = -1 Then GoTo RecordStatementAndRelevance
    N = SubjectAlternativeNumber
    For W = 0 To CurrentAlternatives.WordsAmount(N) - 1
    If StatementSubject = CurrentAlternatives.Words(N, W) Then Relevance = Relevance + 8: Dword = "Subjects match indirectly": GoSub AddDebugWord: GoTo RecordStatementAndRelevance
    Next W




RecordStatementAndRelevance:
    If Relevance < 0 Then Relevance = 0
    If Relevance > 100 Then Relevance = 100
'    TTT = Timer
'TWLoop:
'    If Timer < (TTT - 0.1) Then GoTo TWLoop


' Now check to see if we have a highly relevant statement, if we have then don't bother doing anymore
' searches, just reply with this one.

RecStats:
    If SearchMode = AllRelevant Then GoTo RecStats_Original
    If Relevance < 95 Then GoTo RecStats_Original

    Statements(RecNumber) = OriginalStatement
    Relevances(0) = Relevance
    Indexes(0) = RecNumber
    GoSub RecordDebugInfo
    RecNumber = RecNumber + 1
    GoTo Done


RecStats_Original:
    MemoryPercentage = ((100 / Len(Memory)) * MemoryPosition)
    Relevance = Relevance - ((100 - MemoryPercentage) / 1000)

    If Relevance < (40 + (RecNumber / 2)) Then GoTo Loop0
    Statements(RecNumber) = OriginalStatement
    Relevances(RecNumber) = Relevance
    Indexes(RecNumber) = RecNumber
    GoSub RecordDebugInfo
    RecNumber = RecNumber + 1
    If RecNumber = 99 Then Debug.Print "Overflowed with relevant statements!!!": GoTo SortRelevantStatements
    GoTo Loop0




SortRelevantStatements:
    If RecNumber < 2 Then GoTo Done
Loop4:
    Change = False
    For N = 0 To RecNumber - 2
    If Relevances(N) < Relevances(N + 1) Then Temp = Relevances(N): Relevances(N) = Relevances(N + 1): Relevances(N + 1) = Temp: Temp = Indexes(N): Indexes(N) = Indexes(N + 1): Indexes(N + 1) = Temp: Change = True
    Next N
    If Change = True Then GoTo Loop4



Done:
    Message = OriginalMessage
    Exit Sub


AddDebugWord:
    If DW = 19 Then Return
    DebugWords(DW) = Dword: DebugWords(DW + 1) = "": DW = DW + 1
    Return



RecordDebugInfo:
    W = 0
RSLoop:
    If DebugWords(W) = "" Then MatchedWords(RecNumber, W) = "": Return
    MatchedWords(RecNumber, W) = DebugWords(W)
    W = W + 1
    GoTo RSLoop



GetCondition:
    Condition = WordsBetween("If", "then", Statement)
    Statement = Condition
    Return
    

GetSubjectAlternativeNumber:
    SubjectAlternativeNumber = -1
    N = 0
GNWLoop1:
    If CurrentAlternatives.Words(N, 0) = "" Then Return
    If CurrentAlternatives.Words(N, 0) Like Subject Then If CurrentAlternatives.WordsAmount(N) > 1 Then SubjectAlternativeNumber = N: Return
    N = N + 1
    GoTo GNWLoop1




End Sub



Public Sub AddMatchWord(Word1, Word2)

'remove quotes etc

If Word1 = "" Then GoTo Error
If Word2 = "" Then GoTo Error

If Word1 Like "[!a-z]*[!a-z]" Then Word1 = Mid(Word1, 2, Len(Word1) - 2)
If Word2 Like "[!a-z]*[!a-z]" Then Word2 = Mid(Word2, 2, Len(Word2) - 2)

Word1 = LCase(Word1)
Word2 = LCase(Word2)

If Word1 Like "* *" Then Word1 = "'" + Word1 + "'"
If Word2 Like "* *" Then Word2 = "'" + Word2 + "'"


FindKey:
P = InStr(1, MatchWords, ">" + Word1 + "=", vbBinaryCompare)
If P = 0 Then GoTo AddNewKey
WordsEnd = InStr(P, MatchWords, Chr(13) + Chr(10), vbBinaryCompare)
If Mid(MatchWords, P, WordsEnd - P) Like "*^" + Word2 + "^*" Then Exit Sub
Words = Mid(MatchWords, P, WordsEnd - P)
MatchWords = Replace(MatchWords, Words, Words + Word2 + "^", 1, 1, vbBinaryCompare)
Exit Sub


AddNewKey:
MatchWords = MatchWords + ">" + Word1 + "=" + "^" + Word2 + "^" + Chr(13) + Chr(10)

Exit Sub

Error:
Debug.Print "AddMatchWord() Error!!"


End Sub


Public Sub AddFullStop(Message)
If Message = "" Then Exit Sub
Message = Trim(Message)
If Message Like "*." Then Exit Sub
Message = Message + "."
End Sub


Public Function GetCommand(Message)
GetCommand = ""
Commands = Array("Say", "Talk", "Tell", "Quit", "Exit", "Terminate", "Shutdown", "Use", "Abort", "Stop", "Show", "Hide", "Display", "Select", "Please", "Shut", "Shutup", "Fly", "use", "Load", "Enable", "Disable", "Run", "Start", "Execute", "Jump", "Delete", "Make", "Launch", "Fly", "Walk", "Go", "Edit", "swim", "")
Loop1:
If Commands(N) = "" Then Exit Function
If Message Like Commands(N) + " *" Then GoTo Done
N = N + 1
GoTo Loop1

Done:
If Message Like Commands(N) + " means *" Then Exit Function
If Message Like Commands(N) + " is another *" Then Exit Function
If Message Like Commands(N) + " is the *" Then Exit Function
If Message Like Commands(N) + " is not *" Then Exit Function
GetCommand = Commands(N)


End Function

' Adds the thing (noun) to the "Things" array
Public Sub AddThing(Thing)
If Thing <> "" Then GoTo DoIt
Exit Sub

DoIt:
If Things = "" Then GoTo AddThing
P = InStr(1, Things, Thing + Chr(13) + Chr(10), vbBinaryCompare)
If P = 0 Then GoTo AddThing
Exit Sub


AddThing:
Things = Things + Thing + Chr(13) + Chr(10)



End Sub

Public Sub ArrayTest()
Dim At() As String
Blah = Array("AB", "AB", "DAVE", "AB", "AB")



At = Filter(Blah, "DAVE")


End Sub

' This will change words like 'Your' for 'My' and 'My' for 'Your' etc.
' It wont change any of the words that are within quotes.
Public Sub ChangeYourForMyEtc(Statement)
If Statement = "" Then Exit Sub
Message = Statement
Statement = ""
Dim Words(100)
Call BreakUpSentence(Message, Words())


MiscWords = Array("and", "when", "if", "from", "or", "then", "where", "a", "about", "how", "I", "[.,!]", "")

N = 0
Loop1:
If Words(N) = "" Then GoTo Done


If Words(N) = Chr(34) Then
If QuotesOpen = True Then QuotesOpen = False: GoTo CheckWords
QuotesOpen = True: GoTo CheckWords
End If


If Words(N) = "'" Then
If QuotesOpen = True Then QuotesOpen = False: GoTo CheckWords
QuotesOpen = True: GoTo CheckWords
End If



CheckWords:
If QuotesOpen = True Then GoTo GetNext
If Words(N) Like "yours" Then Words(N) = "mine": GoTo GetNext
If Words(N) Like "mine" Then Words(N) = "yours": GoTo GetNext
If Words(N) Like "you" Then GoTo CheckYou
If Words(N) Like "I" Then Words(N) = "you": GoTo GetNext
If Words(N) Like "me" Then Words(N) = "you": GoTo GetNext
If Words(N) Like "are" Then If Words(N + 1) Like "you" Then Words(N) = "am": GoTo GetNext
If Words(N) Like "are" Then If N > 0 Then If Words(N - 1) Like "I" Then Words(N) = "am": GoTo GetNext
If Words(N) Like "am" Then If N > 0 Then If Words(N - 1) Like "[!0-9]*" Then Words(N) = "are": GoTo GetNext
If Words(N) Like "my" Then Words(N) = "your": GoTo GetNext
If Words(N) Like "your" Then Words(N) = "my"
GetNext:
Statement = Statement + " " + Words(N)
N = N + 1
GoTo Loop1

Done:
Call RemoveUnwantedSpaces(Statement)
Call Capitalise(Statement)
Exit Sub


CheckYou:
W = 0
CWLoop:
If MiscWords(W) = "" Then Words(N) = "I": GoTo GetNext
If Words(N + 1) Like MiscWords(W) Then Words(N) = "me": GoTo GetNext
W = W + 1
GoTo CWLoop





End Sub

Public Sub GetMatchWords(Message)


Statement = Message
Call TieAlternatives(Statement)
Static Words(1000)

Call BreakUpSentence(Statement, Words())

RemoveCrapStuff:
N = 0
C = 0
RPLoop:
If Words(N) = "" Then Words(C) = "": GoTo SetBase
If Words(N) = "?" Then GoTo Skip
If Words(N) = "." Then GoTo Skip
If Words(N) = "," Then GoTo Skip
If Words(N) = "!" Then GoTo Skip
If Words(N) = "-" Then GoTo Skip
If Words(N) = Chr(34) Then GoTo Skip
Words(C) = Words(N)
C = C + 1
Skip:
N = N + 1
GoTo RPLoop

' set the base words
SetBase:
N = 0
Loop0:
Call ChangeAll("-", " ", Words(N))
CurrentAlternatives.Words(N, 0) = LCase(Words(N))
CurrentAlternatives.WordsAmount(N) = 1
If Words(N) = "" Then GoTo CheckEm
N = N + 1: If N = 29 Then CurrentAlternatives.Words(N, 0) = "": GoTo CheckEm
GoTo Loop0



CheckEm:
If MatchWords = "" Then Exit Sub
W = 0
Loop1:
If CurrentAlternatives.Words(W, 0) = "" Then Exit Sub
KeyWord = CurrentAlternatives.Words(W, 0)
GoSub GetOtherVariations
P = InStr(1, MatchWords, ">" + KeyWord + "=", vbBinaryCompare)
If P > 0 Then GoSub GetEm: GoTo GetNext
If CurrentAlternatives.Words(W, 1) = "" Then GoTo GetNext
KeyWord = CurrentAlternatives.Words(W, 1)
P = InStr(1, MatchWords, ">" + KeyWord + "=", vbBinaryCompare)
If P > 0 Then GoSub GetEm
GetNext:
W = W + 1
GoTo Loop1




GetEm:
WordsEnd = InStr(P, MatchWords, Chr(13) + Chr(10), vbBinaryCompare)
StartPos = P + Len(KeyWord) + 3
Loop4:
If StartPos >= WordsEnd Then Return
EndPos = InStr(StartPos, MatchWords, "^", vbBinaryCompare)
MatchWord = Mid(MatchWords, StartPos, EndPos - StartPos)
StartPos = EndPos + 1
If CurrentAlternatives.WordsAmount(W) > 168 Then Debug.Print "MATCHWORDS OVERFLOW!!": Return
NewMatchWord = MatchWord
GoSub AddMatch
GoSub GetOtherVariations
GoTo Loop4




GetOtherVariations:
    Amount = CurrentAlternatives.WordsAmount(W)
    MatchWord = CurrentAlternatives.Words(W, Amount - 1)
    CurrentAlternatives.Words(W, Amount) = ""
    If MatchWord Like "*'s" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 2): GoSub AddMatch: Return
    If MatchWord Like "*??ing" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 3): GoTo CheckDouble
    If MatchWord Like "*???y" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 1): GoTo CheckDouble
    If MatchWord Like "*???iest" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 4): GoTo CheckDouble
    If MatchWord Like "*??est" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 3): GoTo CheckDouble
    If MatchWord Like "*???ed" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 2): GoTo CheckDouble
    If MatchWord Like "*??s" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 1): GoSub AddMatch: Return
    If MatchWord Like "*???er" Then NewMatchWord = Left(MatchWord, Len(MatchWord) - 2): GoTo CheckDouble
    Return
CheckDouble:
    If Right(NewMatchWord, 1) = Mid(NewMatchWord, Len(NewMatchWord) - 1, 1) Then NewMatchWord = Left(NewMatchWord, Len(NewMatchWord) - 1)
    GoSub AddMatch
    Return




GetOtherVariations_Old:
    Amount = CurrentAlternatives.WordsAmount(W)
    MatchWord = CurrentAlternatives.Words(W, Amount - 1)
    MatchWord2 = MatchWord
    If MatchWord Like "*'s" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 2): GoSub AddMatch: Return
    If MatchWord Like "*??ing" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 3): GoTo CheckIt
    If MatchWord Like "*???y" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 1): GoTo CheckIt
    If MatchWord Like "*???iest" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 4): GoTo CheckIt
    If MatchWord Like "*??est" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 3): GoTo CheckIt
    If MatchWord Like "*??s" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 1): GoTo CheckIt
    If MatchWord Like "*???er" Then MatchWord2 = Left(MatchWord, Len(MatchWord) - 2)


CheckIt:
    If Len(MatchWord2) > 4 Then NewMatchWord = MatchWord2: GoSub AddMatch: Return
    NewMatchWord = MatchWord2 + "ing": GoSub AddMatch
    NewMatchWord = MatchWord2 + "iest": GoSub AddMatch
    NewMatchWord = MatchWord2 + "est": GoSub AddMatch
    NewMatchWord = MatchWord2 + "y": GoSub AddMatch
    NewMatchWord = MatchWord2 + "er": GoSub AddMatch
    If Right(MatchWord2, 1) = Mid(MatchWord2, Len(MatchWord2) - 1, 1) Then NewMatchWord = Left(MatchWord2, Len(MatchWord2) - 1): MatchWord2 = NewMatchWord: GoSub AddMatch
    NewMatchWord = MatchWord2 + "s": GoSub AddMatch
    Return



AddMatch:
    Amount = CurrentAlternatives.WordsAmount(W)
    If Amount > 168 Then Return
    CurrentAlternatives.Words(W, Amount) = LCase(NewMatchWord): Amount = Amount + 1
    CurrentAlternatives.WordsAmount(W) = CurrentAlternatives.WordsAmount(W) + 1
    Return




End Sub

Public Sub GetStoredMatchWords()

Label = "* MatchWords:"


WordsPosition = InStr(1, Memory, Label + Chr(13) + Chr(10), vbBinaryCompare)
If WordsPosition = 0 Then Exit Sub
EndPosition = InStr(WordsPosition, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If EndPosition = 0 Then EndPosition = Len(Memory)

EndPosition = EndPosition + 2


ActualWordsPosition = WordsPosition + Len(Label) + 2
MatchWords = Mid(Memory, ActualWordsPosition, EndPosition - ActualWordsPosition)

UpperSection = Left(Memory, WordsPosition - 1)
LowerSection = Right(Memory, Len(Memory) - (EndPosition - 1))

Memory = UpperSection + LowerSection


End Sub

Public Sub GetWontSayData()


Label = "* WontSay:"

WordsPosition = InStr(1, Memory, Label + Chr(13) + Chr(10), vbBinaryCompare)
If WordsPosition = 0 Then Exit Sub
EndPosition = InStr(WordsPosition, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If EndPosition = 0 Then EndPosition = Len(Memory)

EndPosition = EndPosition + 2


ActualWordsPosition = WordsPosition + Len(Label) + 2
WontSayData = Mid(Memory, ActualWordsPosition, EndPosition - ActualWordsPosition)

UpperSection = Left(Memory, WordsPosition - 1)
LowerSection = Right(Memory, Len(Memory) - (EndPosition - 1))

Memory = UpperSection + LowerSection



End Sub


Public Sub LoadMemory()

On Error Resume Next

MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"

Open MyPath + "Memory.txt" For Input As #1
If Err.Number <> 0 Then Err.Clear: GoTo CheckDat
Size = LOF(1)
If Size > 100 Then Size = 100
Memory = Input(Size, #1)
Close #1
' 34563 is the code that is needed at the start of plain ascii memory file if I am to load it in.
P = InStr(1, Memory, "* 34563 *", vbBinaryCompare)
If P = 0 Then Close #1: GoTo CheckDat
Open MyPath + "Memory.txt" For Input As #1
Size = LOF(1)
Memory = Input(Size, #1)
Close #1
Mid(Memory, P, 9) = "* 44444 *"


GoTo DoRoughFormat


' Load an encrypted ascii memory file.
CheckDat:
Open MyPath + "Memory.Dat" For Input As #1
If Err.Number <> 0 Then GoTo WarnNoMemoryFile
Size = LOF(1)
Memory = Input(Size, #1)
Close #1




If Len(Memory) < 200 Then GoTo Decrypt

If InStr(1, Memory, Chr(149), vbBinaryCompare) Then GoTo Decrypt

GoTo WarnAboutMemoryFormat


Decrypt:
Memory = Replace(Memory, Chr(142), "a", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(143), "e", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(144), "i", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(145), "o", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(146), "u", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(147), "s", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(148), "t", 1, -1, vbBinaryCompare)
Memory = Replace(Memory, Chr(149), " ", 1, -1, vbBinaryCompare)



DoRoughFormat:

Call GetStoredMatchWords
Call GetWontSayData


GoSub FormatMemory





Exit Sub












FormatMemory:
If Len(Memory) = 0 Then Return
'Remove any excess return codes from Memory
Memory = Replace(Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10), Chr(13) + Chr(10) + Chr(13) + Chr(10))

'Remove any unwanted spaces from the ends of the statements within memory.
Memory = Replace(Memory, " " + Chr(13) + Chr(10) + Chr(13) + Chr(10), Chr(13) + Chr(10) + Chr(13) + Chr(10))

' make sure all statements within memory have a fullstop on the end.
P = 1
Loop1:
P = InStr(P, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If P = 0 Then GoTo CheckMemEnd
If Mid(Memory, P - 1, 1) <> "." Then Memory = Left(Memory, P - 1) + "." + Right(Memory, Len(Memory) - (P - 1))
P = P + 4
GoTo Loop1

CheckMemEnd:
P = Len(Memory)
CMLoop:
C = Mid(Memory, P, 1)
P = P - 1
If P = 0 Then Return
If C = Chr(13) Then GoTo CMLoop
If C = Chr(10) Then GoTo CMLoop
Memory = Left(Memory, P + 1) + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Return




'-----------------------------------------

WarnAboutMemoryFormat:
Msg = "The memory file is in the wrong format."   ' Define message.
Style = vbOKOnly + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Error!!"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
'If Response = vbOK Then End
End

WarnNoMemoryFile:
Msg = "Can't find a memory file within the current folder:- " + App.Path ' Define message.
Style = vbOKOnly + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Error!!"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
End


End Sub


Public Sub SaveNewKnowledge()



Err.Clear
On Error Resume Next
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "General.FCT" For Append As #1
'If Err.Number <> 0 Then Err.Clear: GoTo FileError
Print #1, GeneralKnowledge
'If Err.Number <> 0 Then Err.Clear: GoTo FileError
Close #1



Err.Clear
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "Personal.FCT" For Append As #1
'If Err.Number <> 0 Then Err.Clear: GoTo FileError
Print #1, PersonalKnowledge
'If Err.Number <> 0 Then Err.Clear: GoTo FileError
Close #1

Exit Sub


FileError:
Msg = "There was a problem trying to save a fact file!" ' Define message.
Style = vbOKOnly + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Error!!"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)



End Sub

Public Sub SetBotsMessage(Message)

TimeOfLastBotActivity = (Date + Time)

'If Timer < (BotsMessageTime + Len(BotsMessage) / 50) Then Exit Sub

BotsMessage = Message
BotsMessageTime = Timer

End Sub

Public Sub SetUsersMessage(Message)
UsersMessage = Message
UsersMessageTime = Timer
NewMessageAvailable = True
TimeOfLastUserActivity = (Date + Time)
End Sub

Public Sub SetInternalMessage(Message)
InternalMessage = Message
InternalMessageTime = Timer

End Sub

Public Sub StoreMatchWords()

Memory = Memory + "* MatchWords:" + Chr(13) + Chr(10)
Memory = Memory + MatchWords + Chr(13) + Chr(10)


End Sub

Public Sub StoreWontSayData()

Memory = Memory + "* WontSay:" + Chr(13) + Chr(10)
Memory = Memory + WontSayData + Chr(13) + Chr(10)


End Sub


' This tidies the users message, corrects any mistakes etc.
Public Sub TidyMessage(Message)
If Message = "" Then Exit Sub

Call RemoveUnwantedSpaces(Message)


'Capitalise the first letter.
Call Capitalise(Message)

' Add a full stop to the end
'If Message Like "*." Then Exit Sub
'Message = Message + "."
Call AddFullStop(Message)

End Sub

'Adds the specified item to Bob's memory.
Public Sub AddMemoryItem(Item)
Memory = Memory + Item + Chr(13) + Chr(10) + Chr(13) + Chr(10)
End Sub



' This will compare two strings and return True if they match
' The first string may contain the '*' wildcard and references categories.
' For example, lets say "place=outside,inside,above" is within the Category List
' Then we had called this function with the following strings:
' String1="The * is <place>"
' String2="The cat is outside"
' The function would then return a True result because outside is a place.

Public Function CategoryCompare(String1, String2)
    Static PreviousString1
    Static MatchWord(50)
    Dim Cat(20)
    If String1 <> PreviousString1 Then GoSub ExtractWordsFromString1: PreviousString1 = String1
    CategoryCompare = False

    N = 0
M3_CheckMatch:
M3_Loop3:
    Word = ExtractWord(String2)
    If Word = "" Then Exit Function
    If MatchWord(N) Like "*<*>*" Then GoTo CompareWithCategory
    If MatchWord(N) = "*" Then GoTo M3_CheckMore
    If Word Like MatchWord(N) Then GoTo M3_CheckMore

CheckForWildCard:
    If N = 0 Then GoTo M3_Loop3 Else If MatchWord(N - 1) = "*" Then GoTo M3_Loop3
    Exit Function
    
M3_CheckMore:
    N = N + 1
    If MatchWord(N) = "" Then CategoryCompare = True: Exit Function
    GoTo M3_CheckMatch



CompareWithCategory:
    Position = 1
    Category = WordsBetween("<", ">", MatchWord(N))
FindCategory:
    Position = InStr(Position, Categories, Category + "=")
    If Position = 0 Then Exit Function
    If Position > 1 Then If Mid(Categoryies, Position - 1, 1) Like "[a-z]" Then Position = Position + 1: GoTo FindCategory
    Position = InStr(Position, Categories, "=") + 1
    C = 0
    CatWord = ""
    LeftPart = WordsBetween("", "<", MatchWord(N))
    RightPart = WordsBetween(">", "", MatchWord(N))
ExtractCategoryItems:
    Char = Mid(Categories, Position, 1)
    Position = Position + 1
    If Char = "," Then Cat(C) = LeftPart + CatWord + RightPart: C = C + 1: CatWord = "": GoTo ExtractCategoryItems
    If Char = "." Then Cat(C) = LeftPart + CatWord + RightPart: Cat(C + 1) = "": GoTo CW_CheckMatch
    CatWord = CatWord + Char
    GoTo ExtractCategoryItems

CW_CheckMatch:
    C = 0
CW_Loop1:
    Match = Cat(C)
    If Match = "" Then If N = 0 Then GoTo CW_GetNextWord Else If MatchWord(N - 1) = "*" Then GoTo CW_GetNextWord Else Return
    If Word Like Match Then GoTo M3_CheckMore
    C = C + 1
    GoTo CW_Loop1
CW_GetNextWord:
    Word = ExtractWord(String2)
    If Word = "" Then Exit Function
    GoTo CW_CheckMatch





ExtractWordsFromString1:
    TemporaryString = String1
    Word = ""
    N = 0
    Position = 1
EW_Loop1:
    C = Mid(TemporaryString, Position, 1)
    Position = Position + 1
    If C Like "[a-z]" Then Word = Word + C: GoTo EW_Loop1
    If C Like "[0-9]" Then Word = Word + C: GoTo EW_Loop1
    If C = "'" Then Word = Word + C: GoTo EW_Loop1
    If C = "<" Then Word = Word + C: GoTo EW_Loop1
    If C = ">" Then Word = Word + C: GoTo EW_Loop1
    If C = "*" Then Word = Word + C: GoTo EW_Loop1
    GoSub StoreWord
    If Position <= Len(TemporaryString) Then GoTo EW_Loop1
    MatchWord(N) = ""
    Return

StoreWord:
    MatchWord(N) = Word: N = N + 1: Word = ""
    Return
 
 
 
 End Function


' This will take a message like:- "I am going out for a meal on my Birthday"
' and convert it to:- "I am going out for a meal on 1/10/71"
' It can only convert the event if there is a memory item that refers to the event's date.
' Like:- "My birthday is on the 1st of October."

Public Sub ConvertEventToDate(Message)
If Message = "" Then Exit Sub
NewMessage = Message
    Call ChangeToPerson(NewMessage)
    Call ChangeToSubject(NewMessage)
    Call ChangeYourForMyEtc(NewMessage)


    RestartItemSearch
    ItemSearchMethod = WordForWord
    ItemSearchStyle = Strict
Loop1:
    Statement = NextItemContaining("is")
    If Statement = "" Then Exit Sub
    
Check_Stage2:
    If Statement Like "*#/*#/#*" Then GoTo Check_Stage3
    For I = 1 To 12
    If Statement Like "* " + MonthName(I, True) + "*" Then GoTo Check_Stage3
    Next I
    GoTo Loop1


Check_Stage3:
    If Statement Like "It is * on the *" Then TheEvent = WordsBetween("It is", "on the", Statement): TheDate = WordsBetween("on the", "", Statement): GoTo CheckIt
    If Statement Like "on the * it is *" Then TheEvent = WordsBetween("it is", "", Statement): TheDate = WordsBetween("on the", "it is", Statement)
    If Statement Like "* is on the *" Then TheEvent = WordsBetween("", "is on the", Statement): TheDate = WordsBetween("on the", "", Statement): GoTo CheckIt
    If Statement Like "* starts on the *" Then TheEvent = WordsBetween("", "starts on the", Statemet): TheDate = WordsBetween("starts on the", "", Statement): GoTo CheckIt
    GoTo Loop1


CheckIt:
    If NewMessage Like "*" + TheEvent + "*" Then Call ReplaceWords(TheEvent, TheDate, NewMessage): Message = NewMessage
    If TheEvent Like "christmas day" Then TheEvent = "christmas"
    If NewMessage Like "*" + TheEvent + "*" Then Call ReplaceWords(TheEvent, TheDate, NewMessage): Message = NewMessage
    GoTo Loop1



End Sub

' This will convert a fraction operation into a more common maths format
' Example:
' String="two thirds of nine"
' ConvertFractions(String)
' String becomes "((2/3*9)"
'

Private Sub ConvertFractions2(Statement)
If Statement Like "* of *" Then GoTo Check
Exit Sub


Check:
Numbers1 = Array("zeroth", "oneth", "half", "third", "forth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth", "eleventh", "twelfth", "thirteenth", "fourteenth", "fifteenth", "sixteenth", "seventeenth", "eighteenth", "nineteenth", "twentieth", "thirtieth", "fortieth", "fiftieth", "sixtieth", "seventieth", "eightieth", "ninetieth", "twentieth", "hundredth", "thousandth", "millionth", "billionth")

If Statement Like "*quarter of*" Then ReplaceWords "quarter of", "forth of", Statement
If Statement Like "* quarters of *" Then ReplaceWords "quarters of", "forth of", Statement
If Statement Like "* of a *" Then ReplaceWords "of a", "of", Statement
If Statement Like "*halves of*" Then ReplaceWords "halves of", "half of", Statement


Dim Words(100)
GoSub ExtractValues


W = 0
Loop10:
If Words(W) = "" Then GoTo Done
For N = 0 To 31
If Words(W) Like Numbers1(N) Then GoSub ChangeIt: GoTo CheckNext
If Words(W) Like Numbers1(N) + "s" Then GoSub ChangeIt: GoTo CheckNext
Next N
CheckNext:
W = W + 1
GoTo Loop10
Done:
GoSub BuildConvertedStatement
Exit Sub



BuildConvertedStatement:
W = 0
Statement = ""
Do
If Words(W) = "" Then Statement = Trim(Statement): Return
If Words(W) = "!none!" Then GoTo SkipIt
Statement = Statement + Words(W) + " "
SkipIt:
W = W + 1
Loop


ExtractValues:
Word = ""
Position = 1
Loop1:
 C = Mid(Statement, Position, 1)
 Position = Position + 1
 If C Like "[a-z]" Then Word = Word + C: GoTo Loop1
 If C Like "[0-9]" Then Word = Word + C: GoTo Loop1
 If C = "-" Then Word = Word + C: GoTo Loop1
 If C = "/" Then Word = Word + C: GoTo Loop1
 If C = "+" Then Word = Word + C:: GoTo Loop1
 If C = "*" Then Word = Word + C: GoTo Loop1
 If C Like "[()]" Then Word = Word + C: GoTo Loop1
 
StoreWord:
 If Word = "" Then GoTo CheckEnd
 Words(N) = Word: N = N + 1: Word = ""
CheckEnd:
 If Position <= Len(Statement) Then GoTo Loop1
 Words(N) = ""
Return







ChangeIt:
S = N
If N > 20 Then If N < 30 Then S = ((S - 20) * 10) + 20: GoTo DoIt
If N = 30 Then S = 1000
If N = 31 Then S = 1000000
If N = 32 Then S = 1000000000
DoIt:
If W = 0 Then GoSub Merge2: Return
If Words(W - 1) Like "*[0-9]*" Then GoSub Merge3 Else GoSub Merge2
Return


Merge2:
Words(W) = "((1/" + Trim(Str(N)) + ")" + "*" + Words(W + 2) + ")"
Words(W + 1) = "!none!"
Words(W + 2) = "!none!"
Return

Merge3:
Words(W) = "((" + Words(W - 1) + "/" + Trim(Str(N)) + ")" + "*" + Words(W + 2) + ")"
Words(W - 1) = "!none!"
Words(W + 1) = "!none!"
Words(W + 2) = "!none!"
Return




End Sub



Public Sub ConvertFractionWordsToNumbers(Statement)

Check:
Numbers = Array("third", "3rd", "fourth", "4th", "fifth", "5th", "sixth", "6th", "seventh", "7th", "eighth", "8th", "ninth", "9th", "tenth", "10th", "twelfth", "12th", "thirteenth", "13th", "fourteenth", "14th", "fifteenth", "15th", "sixteenth", "16th", "seventeenth", "17th", "eighteenth", "18th", "nineteenth", "19th", "twentieth", "20th", "thirtieth", "30th", "fourtieth", "40th", "fiftieth", "50th", "sixtieth", "60th", "seventieth", "70th", "eightieth", "80th", "ninetieth", "90th", "hundredth", "100th", "thousandth", "1000th", "millionth", "1000000th", "billionth", "1000000000th", "")


N = 0
Loop1:
If Numbers(N) = "" Then Exit Sub
If Statement + " " Like "*" + Numbers(N) + "[!a-z]*" Then ReplaceWords Numbers(N), Numbers(N + 1), Statement
If Statement + " " Like "*" + Numbers(N) + "s[!a-z]*" Then ReplaceWords Numbers(N) + "s", Numbers(N + 1) + "s", Statement
N = N + 2
GoTo Loop1



End Sub


Public Sub ConvertFractions(Statement)
If Statement Like "* of *" Then GoTo Check
If Statement Like "* an *" Then GoTo Check
Exit Sub



Check:
Statement = WordsToNumbers(Statement)

GoSub ConvertVarious:


Loop1:
If Statement Like "*#rd of #*" Then Text = "rd of ": GoTo CheckMore
If Statement Like "*#th of #*" Then Text = "th of ": GoTo CheckMore
If Statement Like "*#rds of #*" Then Text = "rds of ": GoTo CheckMore
If Statement Like "*#ths of #*" Then Text = "ths of ": GoTo CheckMore
Exit Sub


CheckMore:
P = 1
CM_Loop1:
P = InStr(P, Statement, Text)
If P = 0 Then Exit Sub
C = Mid(Statement, P - 1, 1)
If C Like "[0-9]" Then GoTo CM_CheckRight
GoTo CM_Loop1

CM_CheckRight:
C = Mid(Statement, P + Len(Text), 1)
If C Like "[!0-9]" Then GoTo CM_Loop1


GetValue2:
Value1 = "1"
Value2 = ""
P2 = P - 1
GV2_Scan:
Value2 = Mid(Statement, P2, 1) + Value2
P2 = P2 - 1
If P2 = 0 Then GoTo GetValue3
If Mid(Statement, P2, 1) = "." Then GoTo GV2_Scan
If Mid(Statement, P2, 1) Like "[0-9]" Then GoTo GV2_Scan




GetValue1:
If Mid(Statement, P2, 1) = " " Then GoTo GV1_GetEnd
GoTo GetValue3

GV1_GetEnd:
P2 = P2 - 1
If P2 = 0 Then GoTo GetValue3
If Mid(Statement, P2, 1) = " " Then GoTo GV1_GetEnd
If Mid(Statement, P2, 1) Like "[!0-9]" Then P2 = P2 + 1: GoTo GetValue3

GV1_GetIt:
Value1 = ""
GV1_Scan:
Value1 = Mid(Statement, P2, 1) + Value1
P2 = P2 - 1
If P2 = 0 Then GoTo GetValue3
If Mid(Statement, P2, 1) = "." Then GoTo GV1_Scan
If Mid(Statement, P2, 1) Like "[0-9]" Then GoTo GV1_Scan




GetValue3:
Value3 = ""
P = P + Len(Text)
GV3_Scan:
Value3 = Value3 + Mid(Statement, P, 1)
P = P + 1
If P = Len(Statement) + 1 Then GoTo ReplaceIt
If Mid(Statement, P, 1) Like "[0-9]" Then GoTo GV3_Scan


ReplaceIt:
LeftString = Left(Statement, P2)
RightString = Right(Statement, (Len(Statement) - P) + 1)
Statement = LeftString + "(" + Value1 + "*(" + Value3 + "/" + Value2 + "))" + RightString
Exit Sub





ConvertVarious:
If Statement Like "*halves of #" Then Call ReplaceWords("halves of", "2th of", Statement)
If Statement Like "*halves of an #" Then Call ReplaceWords("halves of an", "2th of", Statement)
If Statement Like "*halves of a #" Then Call ReplaceWords("halves of a", "2th of", Statement)
If Statement Like "*half of an #*" Then Call ReplaceWords("half of an", "2th of", Statement)
If Statement Like "*half of #*" Then Call ReplaceWords("half of", "2th of", Statement)
If Statement Like "*half an #*" Then Call ReplaceWords("half an", "2th of", Statement)
If Statement Like "*half of a #*" Then Call ReplaceWords("half of a", "2th of", Statement)
If Statement Like "*half a #*" Then Call ReplaceWords("half a", "2th of", Statement)
If Statement Like "*quarter of #*" Then Call ReplaceWords("quarter of", "4th of", Statement)
If Statement Like "*quarter of an #*" Then Call ReplaceWords("quarter of an", "4th of", Statement)
If Statement Like "*quarter of a #*" Then Call ReplaceWords("quarter of a", "4th of", Statement)
Return








End Sub




Public Function AlternativeWordV2(Word, W)
If Word = "" Then Exit Function
AlternativeWordV2 = ""
If Word = PreviousAltWord Then GoTo AlreadyGotEm


PreviousAltWord = Word
P = 1
Loop0:
P = InStr(P, WordsData, Chr(13) + Chr(10) + Word, 0)
If P = 0 Then AltWordsV2(0) = "": Exit Function
If Mid(WordsData, P + Len(Word) + 2, 1) = "," Then GoTo Found
P = P + Len(Word) + 2
GoTo Loop0

Found:
EndPosition = InStr(P + 2, WordsData, Chr(13) + Chr(10), 0)

NewWords = Mid(WordsData, P, EndPosition - P)


GoSub ExtractWords

AlreadyGotEm:
AlternativeWordV2 = AltWordsV2(W)



Exit Function

ExtractWords:
    AltWordsV2(0) = ""
    N = 0
    P1 = InStr(1, NewWords, ",")
    If P1 = 0 Then Return
Loop1:
    P2 = InStr(P1 + 1, NewWords, ",")
    If P2 = 0 Then AltWordsV2(N) = "": Return
    AltWordsV2(N) = Mid(NewWords, P1 + 1, (P2 - P1) - 1)
    P1 = P2
    N = N + 1: If N = 1000 Then AltWordsV2(1000) = "": Return
    GoTo Loop1
    

End Function

Public Sub RemoveUnnecessaryWords(Message)
Call ReplaceWords("please", "", Message)

End Sub

Public Sub AddMessageToHistory(Message)
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

End Sub



Public Sub AddVariable(Name, Value)


Static LastAddedVariableNumber

For N = 0 To VarBufferSize - 1
If Name Like VarName(N) Then VarValue(N) = Val(Value): Exit Sub
Next N


If LastAddedVariableNumber = VarBufferSize Then LastAddedVariableNumber = 0
VarName(LastAddedVariableNumber) = Name
Value = Trim(Value)
VarValue(LastAddedVariableNumber) = Val(Value)
LastAddedVariableNumber = LastAddedVariableNumber + 1




End Sub


' This will break a sentence up into words and store the words into an array
Public Sub BreakUpSentence(Sentence, ExtractedWords())

ArraySize = UBound(ExtractedWords)
N = 0
CPosition = 1
Word = ""
Loop1:
 C = Mid(Sentence, CPosition, 1)
 If C = "!" Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = Chr(34) Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = "'" Then If Word = "" Then Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = "'" Then If Mid(Sentence + " ", CPosition + 1, 1) Like "[!a-z]" Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = "." Then If Mid(Sentence + " ", CPosition + 1, 1) = " " Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C Like "[(),]" Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = "?" Then GoSub GetWord: Word = C: GoSub GetWord: GoTo GetNextCharacter
 If C = " " Then GoSub GetWord: GoTo GetNextCharacter
 Word = Word + C
GetNextCharacter:
 CPosition = CPosition + 1
 If CPosition <= Len(Sentence) Then GoTo Loop1
 GoSub GetWord
 ExtractedWords(N) = ""
 Exit Sub


GetWord:
 If N = ArraySize - 1 Then Return
 If Word = "" Then Return
 ExtractedWords(N) = Word: N = N + 1: Word = ""
 Return
 

End Sub


' This will calculate a past or future time and date by adding the specified parameters to the current time and date.
' Example:
' > Print Date
' 1/8/99
' > NewTime=CalcDate("1","","","","","")
' > Print NewTime
' Monday the 1st of August, 2000. 1:30 AM
Public Function CalcDate(TheYear, TheMonth, TheDay, TheHour, TheMinute, TheSecond)


GoSub GetYear
GoSub GetMonth
GoSub GetDay
GoSub GetHour
GoSub GetMinute
GoSub GetSecond
GoSub AddDateAndTime
Exit Function


AddDateAndTime:
'NewTime = TimeSerial(Hour(Time) + TheHour, Minute(Time) + TheMinute, Second(Time) + TheSecond) + Date

'TheTime = Format(NewTime, "hh:mm:ss AMPM")

'NewDate = Format(NewTime, "d/m/yyyy")

'thedate = DateSerial(Year(NewDate) + TheYear, Month(NewDate) + TheMonth, Day(NewDate) + TheDay)

TheDate = Date + Time
TheDate = DateAdd("yyyy", TheYear, TheDate)
TheDate = DateAdd("m", TheMonth, TheDate)
TheDate = DateAdd("d", TheDay, TheDate)
TheDate = DateAdd("h", TheHour, TheDate)
TheDate = DateAdd("n", TheMinute, TheDate)
TheDate = DateAdd("s", TheSecond, TheDate)

TheYear = Format(TheDate, "yyyy")
TheMonth = Format(TheDate, "mmmm")
TheDay = Format(TheDate, "dddd")
TheDayNumber = Format(TheDate, "d")
TheDayNumber = Date2to2ndEtc(TheDayNumber)
CalcDate = "Time: " + Format(TheDate, "h:mm:ss AMPM") + "." 'TheTime > TheDate
CalcDate = CalcDate + " Day: " + TheDay + "."
CalcDate = CalcDate + " Month: " + TheMonth + "."
CalcDate = CalcDate + " Day of month: " + TheDayNumber + "."
CalcDate = CalcDate + " Year: " + TheYear + "."
CalcDate = CalcDate + " Date: " + Format(TheDate, "d/m/yyyy") + "."
Return


GetYear:
If TheYear = "" Then TheYear = 0: Return
Number = Val(SuperSum(TheYear))
LowerRange = -1500
UpperRange = 6000
GoSub CheckError
TheYear = Number
Return

GetMonth:
If TheMonth = "" Then TheMonth = 0: Return
Number = Val(SuperSum(TheMonth))
LowerRange = -25000:
UpperRange = 94000
GoSub CheckError
TheMonth = Number
Return

GetDay:
If TheDay = "" Then TheDay = 0: Return
Number = Val(SuperSum(TheDay))
LowerRange = -600000
UpperRange = 2800000
GoSub CheckError
TheDay = Number
Return

GetHour:
If TheHour = "" Then TheHour = 0: Return
Number = Val(SuperSum(TheHour))
LowerRange = -10000000
UpperRange = 60000000
GoSub CheckError
TheHour = Number
Return

GetMinute:
If TheMinute = "" Then TheMinute = 0: Return
Number = Val(SuperSum(TheMinute))
LowerRange = -900000000
UpperRange = 2000000000
GoSub CheckError
TheMinute = Number
Return

GetSecond:
If TheSecond = "" Then TheSecond = 0: Return
Number = Val(SuperSum(TheSecond))
LowerRange = -900000000
UpperRange = 2000000000
GoSub CheckError
TheSecond = Number
Return

CheckError:
If Number > UpperRange Then GoTo Error
If Number < LowerRange Then GoTo Error
Return

Error:
CalcDate = "Error"
Exit Function

End Function




' This will place the next sequence of alternative words into <statement>
' The words are taken from the alternative-words-list created by the AlternativesAmount Function.
'
' Another procedure related to this one is:- AddAlternatives
'

Public Sub CycleAlternatives(Statement)
Statement = ""



IncrementAltNumber:
If CurrentAlternatives.Words(N, 0) = "" Then GoTo Loop2
If CurrentAlternatives.WordNumber(N) = CurrentAlternatives.WordsAmount(N) Then CurrentAlternatives.WordNumber(N) = 0: GoTo GetNext
CurrentAlternatives.WordNumber(N) = CurrentAlternatives.WordNumber(N) + 1
GoTo Loop2
GetNext:
N = N + 1: GoTo IncrementAltNumber



Loop2:
If CurrentAlternatives.Words(W, 0) = "" Then GoTo Done
CW = CurrentAlternatives.WordNumber(W)
Word = CurrentAlternatives.Words(W, CW)
Statement = Statement + Word + " "
W = W + 1
GoTo Loop2


Done:
Call RemoveUnwantedSpaces(Statement)


End Sub



Public Function Date2to2ndEtc(Tday)
If Tday = "1" Then Date2to2ndEtc = "1st": Exit Function
If Tday = "2" Then Date2to2ndEtc = "2nd": Exit Function
If Tday = "3" Then Date2to2ndEtc = "3rd": Exit Function
If Tday = "21" Then Date2to2ndEtc = "21st": Exit Function
If Tday = "22" Then Date2to2ndEtc = "22nd": Exit Function
If Tday = "23" Then Date2to2ndEtc = "23rd": Exit Function
If Tday = "31" Then Date2to2ndEtc = "31st": Exit Function
Date2to2ndEtc = Tday + "th"

End Function

Public Function DateWithin(Statement)

Message = WordsToNumbers(Statement)




GetMonth:
If Message Like "*next month[!a-z]*" Then TheDate = DateAdd("m", 1, Date): GoTo GetYear
If Message Like "*last month[!a-z]*" Then TheDate = DateAdd("m", -1, Date): GoTo GetYear
If Message Like "*in a month[!a-z]*" Then TheDate = DateAdd("m", 1, Date): GoTo GetYear

If Message Like "*in *# months*" Then
Amount = Val(WordsBetween("in", "months", Message))
TheDate = DateAdd("m", Amount, Date)
GoTo GetYear
End If

If Message Like "*next *# month*" Then
Amount = Val(WordsBetween("next", "month", Message))
TheDate = DateAdd("m", Amount, Date)
GoTo GetYear
End If

GetYear:





End Function

Public Sub ExtractNumbers(Statement, Numbers())

N = 0
Numbers(0) = ""
Position = 1

FindNumber:
Num = ""
C = Mid(Statement, Position, 1)
If C = "" Then Exit Sub
If C Like "[0-9]" Then GoTo GetNumber
Position = Position + 1
GoTo FindNumber

GetNumber:
Num = Num + C
Position = Position + 1
C = Mid(Statement, Position, 1)
If C Like "[0-9]" Then GoTo GetNumber

Numbers(N) = Num
Numbers(N + 1) = ""
N = N + 1
GoTo FindNumber

End Sub

Public Function ExtractTense(Statement)
' The order of words in this array could prove to be important.
TenseWords = Array("am", "is", "will", "did", "are", "was", "do", "can", "would", "went", "may", "might", "could", "should", "shall", "were", "does", "")
ExtractTense = ""

Loop1:
If TenseWords(N) = "" Then Exit Function
If Statement Like "* " + TenseWords(N) + " *" Then GoTo GetIt
If Statement Like TenseWords(N) + " *" Then GoTo GetIt
If Statement + " " Like "* " + TenseWords(N) + "[!a-z]*" Then GoTo GetIt
N = N + 1
GoTo Loop1

GetIt:
ExtractTense = TenseWords(N)
Call ReplaceWords(ExtractTense, "", Statement)


End Function


' This will extract the first word from a statement.
'
' Example:
' Message="Hello world"
' Word=ExtractWord(Message)
' Print Word
' Hello
' Print Message
' world

Public Function ExtractWord(Statement)
Statement = Trim(Statement)
If Statement = "" Then Exit Function
Position = 1
Loop0:
If Position > Len(Statement) Then Statement = "": ExtractWord = "": Exit Function
C = Mid(Statement, Position, 1)
If C Like "[0-9]" Then GoTo GetWord
If C Like "[!a-z]" Then Position = Position + 1: GoTo Loop0

GetWord:
Word = ""
Loop1:
 C = Mid(Statement, Position, 1)
 Position = Position + 1
 If C = "(" Then GoTo Loop1
 If C Like "[a-z]" Then Word = Word + C: GoTo Loop1
 If C Like "[0-9]" Then Word = Word + C: GoTo Loop1
 If C = "." Then If Word Like "*#" Then If Mid(Statement, Position, 1) Like "#*" Then Word = Word + C: GoTo Loop1
 If C = ":" Then If Word Like "*#" Then If Mid(Statement, Position, 1) Like "#*" Then Word = Word + C: GoTo Loop1
ExtractIt:
ExtractWord = Word
If Position > Len(Statement) + 1 Then Statement = "": Exit Function
Statement = Right(Statement, Len(Statement) - (Position - 2))
Statement = LTrim(Statement)
End Function



' This will replace all the words like:- "I","My" and "Me" with the user's name.

Public Sub ChangeToPerson(Statement)

If Statement = "" Then Exit Sub
Message = Statement
Statement = ""
Dim Words(100)
Call BreakUpSentence(Message, Words())


N = 0
Loop1:
If Words(N) = "" Then GoTo Done

If Words(N) = Chr(34) Then
If QuotesOpen = True Then QuotesOpen = False: GoTo CheckWords
QuotesOpen = True: GoTo CheckWords
End If



If Words(N) = "'" Then
If QuotesOpen = True Then QuotesOpen = False: GoTo CheckWords
QuotesOpen = True: GoTo CheckWords
End If



CheckWords:
If QuotesOpen = True Then GoTo GetNext
If Words(N) Like "my" Then Words(N) = UsersName + "'s": GoTo GetNext
If Words(N) Like "I" Then If Words(N + 1) = "am" Then Words(N) = UsersName: Words(N + 1) = "is": GoTo GetNext
If Words(N) Like "I" Then If Words(N + 1) = "have" Then Words(N) = UsersName: Words(N + 1) = "has": GoTo GetNext
If Words(N) Like "am" Then If Words(N + 1) = "I" Then Words(N) = "is": Words(N + 1) = UsersName: GoTo GetNext
If Words(N) Like "I" Then If Words(N + 1) Like "[a-z]*" Then Words(N) = UsersName: GoTo GetNext
If Words(N) Like "me" Then Words(N) = UsersName
GetNext:
Statement = Statement + " " + Words(N)
N = N + 1
GoTo Loop1

Done:
Call RemoveUnwantedSpaces(Statement)
Call Capitalise(Statement)
Exit Sub






End Sub


' This will try to extract the subject from <message>
' Example:
'
' Message="My cat is outside.
' GetSubject(Message)
'
' Result:
' Subject="My cat"
Public Sub GetSubject(Message)
If Message = "" Then Exit Sub

Subject = ""
BareSubject = ""
If Message Like "* *" Then GoTo HasWords
If Message Like "Which*" Then GoTo Done
If Message Like "Why*" Then GoTo Done
If Message Like "Where*" Then GoTo Done
If Message Like "When*" Then GoTo Done
If Message Like "What*" Then GoTo Done
If Message Like "Who*" Then GoTo Done
If Message Like "How*" Then GoTo Done
If Message Like "did*" Then GoTo Done
If Message Like "yes[!a-z]*" Then GoTo Done
If Message Like "no[!a-z]*" Then GoTo Done
If Message Like "maybe[!a-z]*" Then GoTo Done
If Message Like "because[!a-z]*" Then GoTo Done
If Message Like "perhaps[!a-z]*" Then GoTo Done
If Message Like "[a-z0-9]*" Then NewSubject = Message: GoTo Done
Exit Sub



HasWords:
'If Message Like "* means *" Then Word = WordsBetween("", "means", Message): If Word Like "* *" Then GoTo CheckRest Else NewSubject = Word: GoTo Done
If Message Like "* means *" Then NewSubject = WordsBetween("", "means", Message): GoTo Done
If Message Like "*'s * *" Then GoTo CheckOthers


CheckRest:
If Message Like "*subject*[?]*" Then Exit Sub
If Message Like "Who is it[?]*" Then NewSubject = "My name": GoTo Done
If Message Like "Who am I speaking to[?]*" Then NewSubject = "My name": GoTo Done
If Message Like "Who am I talking to[?]*" Then NewSubject = "My name": GoTo Done
If Message Like "Is it you *[?]*" Then GoTo PossibleName
If Message Like "Is that you *[?]*" Then GoTo PossibleName
If Message Like "* the *est [a-z]*[?]*" Then NewSubject = "the " + WordsBetween("the ", "?", Message): GoTo Done
If " " + Message Like "*[!a-z]talk about *" Then NewSubject = WordsBetween("talk about", "", Message): GoTo Done
If " " + Message Like "*[!a-z]tell me about *" Then NewSubject = WordsBetween("me about", "", Message): GoTo Done
If Message Like "* you know *[!a-z]about *" Then NewSubject = WordsBetween("about", "", Message): GoTo Done

GoSub CheckForThings: If NewSubject <> "" Then GoTo Done


CheckFurther:
If Message Like "* a *[?]*" Then NewSubject = "a " + WordsBetween(" a ", "?", Message): GoTo CheckOthers
If Message Like "* a *" Then NewSubject = "a " + WordsBetween(" a ", ".", Message): GoTo CheckOthers
If Message Like "*if * goes*" Then NewSubject = WordsBetween("If", "goes", Message): GoTo Done
If Message Like "Why do you *[?]*" Then GoTo Type3
If Message Like "Why are you *[?]*" Then NewSubject = "The reason you are " + WordsBetween("you", "?", Message): GoTo Done
If Message Like "* to the *" Then NewSubject = WordsBetween(" to ", ".", Message): GoTo Done
'If Message Like "*There is a *" Then Subject = "The " + WordsBetween("is a", ".", Message): GoTo Done
'If Message Like "*There was a *" Then Subject = "The " + WordsBetween("was a", ".", Message): GoTo Done
If Message Like "*There are no * in*" Then NewSubject = "The " + WordsBetween("are no", "in", Message): GoTo Done
If Message Like "do I like *[?]*" Then NewSubject = WordsBetween(" like ", "?", Message): GoTo Done
If Message Like "do you like *[?]*" Then NewSubject = WordsBetween(" like ", "?", Message): GoTo Done
If Message Like "* has *" Then GoTo GetHas
If Message Like "What does * do[?]*" Then NewSubject = WordsBetween("What does", "do", Message): GoTo Done
If Message Like "Can I *[?]*" Then NewSubject = WordsBetween(" I ", "?", Message): GoTo Done
If Message Like "Do I have *[?]*" Then NewSubject = WordsBetween("have", "?", Message): GoTo Done
If Message Like "* about *" Then NewSubject = WordsBetween(" about ", "", Message): GoTo Done
If Message Like "What * is it[?]*" Then NewSubject = WordsBetween("What", "is it", Message): GoTo Done
GoTo CheckOthers

 



GetHas:
Words = WordsBetween("has", "", Message)
P = InStrRev(Words, " ", -1, vbBinaryCompare)
If P = 0 Then Item = Words Else Item = Mid(Words, P + 1)
NewSubject = WordsBetween("", "has", Message) + "'s " + Item
GoTo Done



CheckOthers:
Word = Array("are", "has", "is", "was", "were", "could", "will", "might", "may", "can", "")

' This section will find the subject in questions like this:
' Where are the dogs? = the dogs
' Where is my cat? = your cat
Part1:
N = 0
Loop1:
If Word(N) = "" Then GoTo MaybePerson
If Message Like "* " + Word(N) + " *" Then GoTo GuessSubject
N = N + 1
GoTo Loop1




GuessSubject:
Item = WordsBetween("", " " + Word(N) + " ", Message)
GoSub CheckIfBad
If NewSubject <> "" Then GoTo Done
Item = WordsBetween(" " + Word(N) + " ", "", Message)
If Right(Item, 2) = "?." Then Item = Left(Item, Len(Item) - 2)
If Right(Item, 1) = "." Then Item = Left(Item, Len(Item) - 1)
GoSub CheckIfBad
GoTo Done


CheckIfBad:
Item2 = " " + Item + " "
If Item2 Like "*[!a-z]what[!a-z]*" Then Return
If Item2 Like "*subject*" Then Return
If Item2 Like "*[!a-z]It[!a-z]*" Then Return
If Item2 Like "*There*" Then Return
If Item2 Like "*The[!a-z]*" Then Return
If Item2 Like "*[!a-z]that*" Then Return
If Item2 Like "*[!a-z]She[!a-z]*" Then Return
If Item2 Like "*[!a-z]He[!a-z]*" Then Return
If Item2 Like "*[!a-z]we[!a-z]*" Then Return
If Item2 Like "*[!a-z]they[!a-z]*" Then Return
If Item2 Like "*[!a-z]you[!a-z]*" Then Return
If Item2 Like "*[!a-z]going[!a-z]*" Then Return
If Item2 Like "*Why*" Then Return
If Item2 Like "*Where*" Then Return
If Item2 Like "*When*" Then Return
If Item2 Like "*What*" Then Return
If Item2 Like "*Who*" Then Return
If Item2 Like "*How*" Then Return

NewSubject = Item
Return


MaybePerson:
If " " + Message Like "*[!a-z]I[!a-z]*" Then NewSubject = "I": GoTo Done
If " " + Message Like "*[!a-z]you[!a-z]*" Then NewSubject = "you"
GoTo Done




PossibleName:
If Message Like "* who *" Then Exit Sub
If Message Like "* that *" Then Exit Sub
NewSubject = "My name"
GoTo Done

Type3:
Item = WordsBetween("do you", "?", Message)
NewSubject = "The reason why you " + Item




Done:
If NewSubject = "" Then Exit Sub
Subject = NewSubject
RemoveEndCrap:
If Right(Subject, 1) Like "[.?!]" Then Subject = Left(Subject, Len(Subject) - 1): GoTo RemoveEndCrap
If BareSubject = "" Then BareSubject = Subject
Exit Sub




CheckForThings:
If Message = "" Then Return
Dim IWords(100)
Statement = Message
Call ExtractWords(Statement, IWords())
N = 0
CTLoop1:
If IWords(N) = "" Then Return
P = InStr(1, Things, Chr(10) + IWords(N) + Chr(13) + Chr(10), vbBinaryCompare)
If P <> 0 Then NewSubject = IWords(N): GoTo AddPreword
N = N + 1
GoTo CTLoop1

AddPreword:
BareSubject = NewSubject
If N = 0 Then Return
If NewSubject = "I" Then Return
If NewSubject = "me" Then Return
If NewSubject = "you" Then Return
If IWords(N - 1) Like "*'s" Then NewSubject = IWords(N - 1) + " " + NewSubject: Return
If IWords(N - 1) Like "a" Then NewSubject = "a " + NewSubject: Return
If IWords(N - 1) Like "an" Then NewSubject = "an " + NewSubject: Return
If IWords(N - 1) Like "the" Then NewSubject = "the " + NewSubject: Return
Return



End Sub


' This will change words like 'Your' for 'My' and 'My' for 'Your' etc.
'
Public Sub ChangeYourForMyEtc_Old(Statement)


ReplaceWords "I am", "ZyouZ are", Statement
ReplaceWords "You are", "ZIZ am", Statement
If Statement Like "*.I.*" Then GoTo S1
ReplaceWords "I", "ZyouZ", Statement
S1:
ReplaceWords "you", "ZmeZ", Statement
ReplaceWords "me", "ZyouZ", Statement
ReplaceWords "Your", "ZmyZ", Statement
ReplaceWords "My", "ZYourZ", Statement

ReplaceWords "ZYourZ", "Your", Statement
ReplaceWords "ZmyZ", "my", Statement
ReplaceWords "ZIZ", "I", Statement
ReplaceWords "ZYouZ", "you", Statement
ReplaceWords "ZmeZ", "me", Statement


MiscWords = Array("must", "should", "do", "could", "would", "will", "can", "wont", "was", "had", "did", "may", "might", "shall", "not", "want", "")

Loop1:
If MiscWords(N) = "" Then Exit Sub
If Statement Like "me *" Then Call ReplaceCharacters("me ", "I ", Statement)
Call ReplaceWords("me " + MiscWords(N), "I " + MiscWords(N), Statement)
Call ReplaceWords(MiscWords(N) + " me", MiscWords(N) + " I", Statement)
N = N + 1
GoTo Loop1



End Sub

   
' This will add a question mark to the end of <message> if it is a question.
'

Public Sub AddQuestionMark(Message)
If Message Like "*[?]*" Then Exit Sub
QuestioningWord = Array("What", "Where", "When is", "When will", "When could", "When did", "When may", "When can", "When was", "When should", "Who", "Why", "Will", "Do you", "Was", "How", "Can", "Show", "Give", "Are", "Is", "Does", "Have", "Which", "")
I = 0
Loop2:
If QuestioningWord(I) = "" Then Exit Sub
If Message Like QuestioningWord(I) + " *" Then GoTo AddIt
I = I + 1
GoTo Loop2

AddIt:
Message = Trim(Message)

Loop1:
If Right(Message, 1) = "." Then Message = Left(Message, Len(Message) - 1): GoTo Loop1
Message = Message + "?."

End Sub


' The word 'an' is used instead of 'a' when it precedes a word beginning with a vowel.
' This will replace any typing mistakes like "a animal" with "an animal"
Public Sub CorrectAWithAn(Message)

Vowel = Array("a", "e", "i", "o", "u")
N = 0
Lp1:
If Message Like "* a " + Vowel(N) + "*" Then ReplaceCharacters " a " + Vowel(N), " an " + Vowel(N), Message: GoTo Lp1
N = N + 1: If N < 5 Then GoTo Lp1

End Sub


' This works similar to the ItemLike function.
' The only difference is that this cares about the order of the words.
' Example:
' If there was a memory item containing this statement: "My cat is outside"
' then you used this function to search for this: "cat outside" the above statement would be returned.
' but it wouldn't be returned if you search for this "outside cat"


Public Function ItemLikeInOrder(TheWords)
ItemLikeInOrder = ""
    Dim Words(20)
    NextV2 = ""

    Call ExtractWords(TheWords, Words())
    If Words(0) = "" Then Exit Function



    N = 1
    MainWord = Words(0)
Loop0:
    If Words(N) = "" Then GoTo CreateWordsToSearch
    WordsToSearch = WordsToSearch + "*[!a-z]" + Words(N)
    If Len(Words(N)) > Len(MainWord) Then MainWord = Words(N)
    N = N + 1
    GoTo Loop0




CreateWordsToSearch:
    N = 0
    WordsToSearch = ""
Loop1:
    If Words(N) = "" Then WordsToSearch = WordsToSearch + "*": GoTo SearchWord
    WordsToSearch = WordsToSearch + "*[!a-z]" + Words(N)
    If ItemSearchStyle = Careless Then GoTo GetNext
    WordsToSearch = WordsToSearch + "[!a-z]"
GetNext:
    N = N + 1
    GoTo Loop1



SearchWord:
    Statement = ItemWith(MainWord)
    If Statement = "" Then Exit Function
    If " " + Statement + " " Like WordsToSearch Then GoTo Done
    GoTo SearchWord


Done:
    ItemLikeInOrder = Statement
    Exit Function



End Function


' Finds a memory item containing specified words
' This function also allows the use of wildcards.
' Example:
'
' Result=NextItem_WildCards(apples * oranges)
' Print Result
' Tom has four apples and three oranges

Public Function NextItem_WildCards(TheWords)
    NextItem_WildCards = ""
    Dim Parts(50)

    If TheWords = "" Then Exit Function

    GoSub GetLargestPart

SearchPart:
    Statement = ItemWith_WildCards(LargestPart)
    If Statement = "" Then Exit Function
    If " " + Statement + " " Like "*" + TheWords + "*" Then GoTo Done
    GoTo SearchPart
Done:
    NextItem_WildCards = Statement
    Exit Function
    


GetLargestPart:
    GoSub ExtractWordParts
    N = 1
    LargestPart = Parts(0)
Loop0:
    If Parts(N) = "" Then Return
    If Len(Parts(N)) > Len(LargestPart) Then LargestPart = Parts(N)
    N = N + 1
    GoTo Loop0
    Return



ExtractWordParts:
    N = 0
    StartPosition = 1
Loop3:
    EndPosition = InStr(StartPosition, TheWords, "*")
    If EndPosition = StartPosition Then StartPosition = StartPosition + 1: GoTo Loop3
    If EndPosition = 0 Then Parts(N) = Mid(TheWords, StartPosition): Return
    Parts(N) = Mid(TheWords, StartPosition, EndPosition - StartPosition)
    StartPosition = EndPosition + 1
    N = N + 1
    GoTo Loop3


End Function


Public Function ItemWith(Item)
    ItemWith = ""
    If MemoryPosition <= 1 Then Exit Function
    Position = MemoryPosition
Loop1:
    Position = InStrRev(Memory, Item, Position - 1)
    If Position = 0 Then ItemWith = "": Exit Function
    If Position = 1 Then GoTo CheckEnd
    If Mid(Memory, Position - 1, 1) Like "[a-z]" Then GoTo Loop1
    
CheckEnd:
    If ItemSearchStyle = Careless Then GoTo Found
    If Mid(Memory, Position + Len(Item), 1) Like "[a-z]" Then GoTo Loop1

Found:
    MemoryPosition = Position
    ItemWith = LocalItem(MemoryPosition)
    
'    P = InStr(position, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10))
'    If (P + 4) >= Len(Memory) Then Exit Function
'    If Mid(Memory, P + 4, 1) = Chr(13) Then Exit Function
    





End Function



Public Function ItemWith_WildCards(Item)
    ItemWith_WildCards = ""
    If MemoryPosition <= 1 Then Exit Function
    Position = MemoryPosition
Loop1:
    Position = InStrRev(Memory, Item, Position - 1)
    If Position = 0 Then ItemWith_WildCards = "": Exit Function
Found:
    MemoryPosition = Position
    ItemWith_WildCards = LocalItem(MemoryPosition)
End Function

' This will return the next memory item (the item stored after the current one)
Public Function NextMemoryItem()
NextMemoryItem = ""
If MemoryPosition = -1 Then MemoryPosition = 1: GoTo GetIt
If MemoryPosition >= Len(Memory) Then Exit Function
Position = InStr(MemoryPosition, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If Position = 0 Then Exit Function

MemoryPosition = Position + 4



GetIt:
EndPos = InStr(MemoryPosition, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If EndPos = 0 Then Exit Function
NextMemoryItem = Mid(Memory, MemoryPosition, EndPos - MemoryPosition)
MemoryPosition = EndPos

End Function



Public Sub OperatorWordsToSymbols(Message)

Call ChangeAll(" Plus ", "+", Message)
Call ChangeAll(" Minus ", "-", Message)
Call ChangeAll(" Multiplied by ", "*", Message)
Call ChangeAll(" Times ", "*", Message)
Call ChangeAll(" Divided by ", "/", Message)
Call ChangeAll(" Add ", "+", Message)
Call ChangeAll(" Take away ", "-", Message)



GoSub InsertMultiplySymbols
GoSub ConvertDivides
GoSub ConvertPercentage

Exit Sub




InsertMultiplySymbols:
If Message Like "*[0-9] [0-9]*" Then GoTo InsertMulti
Return
P = 0
InsertMulti:
P = InStr(P + 1, Message, " ")
If P = 0 Then Return
If P = 1 Then GoTo InsertMulti
If Mid(Message, P - 1, 1) Like "[0-9]" Then If Mid(Message, P + 1, 1) Like "[0-9]" Then Mid(Message, P, 1) = "*"
GoTo InsertMulti






ConvertDivides:
Call ReplaceWords("within", "in", Message)
Call ReplaceWords("an", "", Message)
Call ReplaceWords("a", "", Message)
If Message Like "*[0-9)] into [0-9(]*" Then Word = "into": GoTo ConvertIt
If Message Like "*[0-9)] in [0-9(]*" Then Word = "in": GoTo ConvertIt
If Message Like "*[0-9)] are there in [0-9(]*" Then Word = "are there in": GoTo ConvertIt
If Message Like "*[0-9)] are in [0-9(]*" Then Word = "are in": GoTo ConvertIt
Return
ConvertIt:
P = 0
CILoop:
P = InStr(P + 1, Message, " " + Word + " ") + 1
If P = 0 Then Return
If P = 1 Then GoTo CILoop


GetValue1:
V1Pos = P - 2
Value1 = ""
V1Loop:
If V1Pos = 0 Then GoTo CheckV1
C = Mid(Message, V1Pos, 1)
V1Pos = V1Pos - 1
If C Like "[0-9]" Then Value1 = C + Value1: GoTo V1Loop
If C Like "[.()+/*-]" Then Value1 = C + Value1: GoTo V1Loop
CheckV1:
If Value1 = "" Then GoTo CILoop


GetValue2:
V2Pos = P + Len(Word) + 1
Value2 = ""
V2Loop:
If V2Pos = Len(Message) + 1 Then GoTo CheckV2
C = Mid(Message, V2Pos, 1)
V2Pos = V2Pos + 1
If C = "." Then If V2Pos = Len(Message) + 1 Then GoTo CheckV2
If C Like "[0-9]" Then Value2 = Value2 + C: GoTo V2Loop
If C Like "[.()+/*-]" Then Value2 = Value2 + C: GoTo V2Loop
CheckV2:
If Value2 = "" Then GoTo CILoop


SwitchEm:
NewMessage = Left(Message, V1Pos) + " " + Value2 + "/" + Value1 + Right(Message, Len(Message) - (V2Pos - 1))
Message = NewMessage
Return





ConvertPercentage:
Call ChangeAll(" percent of ", "% of ", Message)
If Message Like "*[0-9)]% of [0-9(]*" Then GoTo ConvPerc
Return
ConvPerc:
P = 0
PCILoop:
P = InStr(P + 1, Message, "% of ")
If P = 0 Then Return
If P = 1 Then GoTo CILoop


GetPValue1:
V1Pos = P - 1
Value1 = ""
PV1Loop:
If V1Pos = 0 Then GoTo PCheckV1
C = Mid(Message, V1Pos, 1)
V1Pos = V1Pos - 1
If C Like "[0-9]" Then Value1 = C + Value1: GoTo PV1Loop
If C Like "[.()+/*-]" Then Value1 = C + Value1: GoTo PV1Loop
PCheckV1:
If Value1 = "" Then GoTo PCILoop


GetPValue2:
V2Pos = P + 5
Value2 = ""
PV2Loop:
If V2Pos = Len(Message) + 1 Then GoTo PCheckV2
C = Mid(Message, V2Pos, 1)
V2Pos = V2Pos + 1
If C = "." Then If V2Pos = Len(Message) + 1 Then GoTo PCheckV2
If C Like "[0-9]" Then Value2 = Value2 + C: GoTo PV2Loop
If C Like "[.()+/*-]" Then Value2 = Value2 + C: GoTo PV2Loop
PCheckV2:
If Value2 = "" Then GoTo PCILoop


SwitchPValues:
NewMessage = Left(Message, V1Pos) + " (" + Value2 + "/100)*" + Value1 + Right(Message, Len(Message) - (V2Pos - 1))
Message = NewMessage
Return















End Sub

' This will return the previous memory item. (the memory item stored above the current one)
Public Function PreviousMemoryItem()
PreviousMemoryItem = ""
If MemoryPosition <= 1 Then Exit Function
Position = InStrRev(Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), MemoryPosition, vbBinaryCompare)
If Position = 0 Then Exit Function
Loop1:
If (Position - 1) <= 0 Then Exit Function
If Asc(Mid(Memory, Position, 1)) < 32 Then Position = Position - 1: GoTo Loop1
MemoryPosition = Position - 1

GetItem:
PreviousMemoryItem = LocalItem(MemoryPosition)

End Function


Public Sub QuestionToAnswer(Question)
If Question = "" Then Exit Sub
If Question Like "*[?]*" Then GoTo SwitchIt
Exit Sub

SwitchIt:

Words = Array("are", "will", "was", "can", "would", "should", "must", "may", "could", "have", "has", "did", "do", "is", "am", "")

N = 0
Loop1:
If Words(N) = "" Then Exit Sub
If Question Like Words(N) + " *" Then GoSub SwitchWords: Exit Sub
N = N + 1
GoTo Loop1



SwitchWords:

Call ChangeAll("?", "", Question)


Word1 = ExtractWord(Question)
Word1 = LCase(Word1)
Word2 = ExtractWord(Question)
Call Capitalise(Word2)
If Word2 Like "the" Then Word2 = Word2 + " " + ExtractWord(Question)
If Word2 Like "*'s" Then Word2 = Word2 + " " + ExtractWord(Question)


Question = Word2 + " " + Word1 + " " + Question
Call AddFullStop(Question)
Call ChangeYourForMyEtc(Question)
Return



End Sub


Public Sub QuestionToNegativeAnswer(Question)
If Question = "" Then Exit Sub
If Question Like "*[?]*" Then GoTo SwitchIt
Exit Sub

SwitchIt:
Words = Array("will", "was", "can", "would", "should", "must", "may", "could", "have", "has", "did", "do", "is", "am", "")

Call ChangeYourForMyEtc(Question)
Call ChangeAll("?", "", Question)

N = 0
Loop1:
If Words(N) = "" Then Exit Sub
If Question Like Words(N) + " *" Then GoSub SwitchWords: Exit Sub
N = N + 1
GoTo Loop1


SwitchWords:
Word1 = ExtractWord(Question)
If Question Like "* not *" Then GoTo CapIt
Word1 = Word1 + " not"

CapIt:
Word1 = LCase(Word1)
Word2 = ExtractWord(Question)
Call Capitalise(Word2)
If Word2 Like "the" Then Word2 = Word2 + " " + ExtractWord(Question)
If Word2 Like "*'s" Then Word2 = Word2 + " " + ExtractWord(Question)

Question = Word2 + " " + Word1 + " " + Question
Call AddFullStop(Question)
Return

End Sub


Public Sub RemoveFullStop(Message)
Loop1:
If Right(Message, 1) = "." Then Message = Left(Message, Len(Message) - 1): GoTo Loop1

End Sub

' This removes excess/unwanted spaces from <message>.
Public Sub RemoveUnwantedSpaces(Message)
Message = Trim(Message)

'Loop1:
'TempMessage = Message
'ReplaceCharacters "  ", " ", Message
'If Message <> TempMessage Then GoTo Loop1

Position = 1
Loop2:
C = Mid(Message, Position, 2)
If C = "  " Then GoTo GetNext

If C = " ," Then C = ",": Position = Position + 1: GoTo GetIt
If C = " ." Then C = ".": Position = Position + 1: GoTo GetIt
If C = " ?" Then C = "?": Position = Position + 1: GoTo GetIt
If C = " !" Then C = "!": Position = Position + 1
GetIt:
NewMessage = NewMessage + Left(C, 1)
GetNext:
Position = Position + 1
If Position <= Len(Message) Then GoTo Loop2

Message = NewMessage


RemoveSpacesNearQuotes:
QuotesOpen = False
Position = 1
Loop3:
If Mid(Message, Position, 1) = Chr(34) Then GoSub FoundQuote
If Mid(Message, Position, 1) = "'" Then GoSub CheckQuote
Position = Position + 1: If Position > Len(Message) Then Exit Sub
GoTo Loop3



CheckQuote:
If Position = 1 Then GoSub FoundQuote: Return
If Position = Len(Message) Then GoSub FoundQuote: Return
If Mid(Message, Position - 1, 3) Like "[a-z0-9]'[a-z0-9]" Then Return
GoSub FoundQuote
Return




FoundQuote:
If QuotesOpen = True Then QuotesOpen = False Else QuotesOpen = True

If QuotesOpen = True Then If Mid(Message + ".", Position + 1, 1) = " " Then Message = Left(Message, Position) + Right(Message, Len(Message) - Position - 1): Return
If QuotesOpen = False Then If Mid(Message, Position - 1, 1) = " " Then Message = Left(Message, Position - 2) + Right(Message, Len(Message) - Position + 1): Return

Return









End Sub

' This will replace words like 'it' and 'he' with the current subject
Public Sub ChangeToSubject(Message)


If Subject <> "" Then
If Message Like "*" + Subject + "*" Then Exit Sub ' no need to change to subject if the subject is already present.
If Subject = UsersName Then Exit Sub
    ReplaceWords "She", Subject, Message
    ReplaceWords "he", Subject, Message
    


    Word = Subject
    If Subject Like "A *" Then Word = Right(Subject, Len(Subject) - 2) + "s"
    ReplaceWords "they", Word, Message



DoIts:
    If Message Like "It is * that *" Then Exit Sub  ' It is estimated that blah blah...
    If Message Like "It has * that *" Then Exit Sub
    If Message Like "*it take the *" Then Exit Sub
    If Message Like "*it take for *" Then Exit Sub
    If Message Like "* will it be in *" Then Exit Sub
    Call ReplaceWords("it", Subject, Message)
End If



End Sub




' Converts numbers from:- '101' to 'one hundred and one'
' Note:- it can't cope with numbers bigger than 9 digits.
Public Function NumbersToWords(Statement)
If Statement = "" Then Exit Function
OldStatement = Statement
If Statement Like "*##########*" Then GoTo AllDone
If Statement Like "*[0-9]*" Then GoTo Convert
GoTo AllDone


Convert:
Number = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
Tens = Array("twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")

' First convert '.' into 'point'
If Statement Like "*.#*" Then GoTo FindPoint Else GoTo DoNumbers

FindPoint:
N = 1
FindLoop:
Characters = Mid(Statement, N, 2)
If Characters Like ".[0-9]" Then GoTo ChangePoint
N = N + 1
GoTo FindLoop

ChangePoint:
EndBit = Mid(Statement, N + 1)
Statement = Left(Statement, N - 1) + " point " + EndBit
GoTo Convert


' Now to convert the numbers
DoNumbers:
' Find the first number
Loop1:
ValueToReplace = ""
N = 1
Do
C = Mid(Statement, N, 1)
N = N + 1
Loop Until C Like "[0-9]"

' Extract the number and place it into 'ValueToReplace'
Loop2:
ValueToReplace = ValueToReplace + C
C = Mid(Statement, N, 1)
N = N + 1
If C Like "[0-9]" Then GoTo Loop2


' If it is '0' then replace it with 'Zero' and check to see if there are anymore numbers.
Value = Val(ValueToReplace)
CN = "": If Value = 0 Then CN = "Zero": GoTo Done

' Convert the number, three digits at a time.
Value = "00000000000" + ValueToReplace
N3 = Mid(Value, Len(Value) - 8, 3)
GoSub GetThree
If C3 <> "" Then CN = C3 + " million"
N3 = Mid(Value, Len(Value) - 5, 3)
GoSub GetThree
If C3 <> "" Then CN = CN + C3 + " thousand"
N3 = Mid(Value, Len(Value) - 2, 3)
GoSub GetThree
If C3 <> "" Then
If CN <> "" Then CN = CN + " and"
CN = CN + C3
End If

Done:
CN = LTrim(CN)
' Replace the old numeric version with the worded version.
ReplaceCharacters ValueToReplace, CN, Statement

' Go and convert the rest of the numbers, if any.
If Statement Like "*[0-9]*" Then GoTo Loop1



' Ok, all done, now quit.
AllDone:
NumbersToWords = Statement
Statement = OldStatement

Exit Function

' This sub will convert three digits like:- '123' into 'one hundred and twenty three'
GetThree:
C3 = ""
V = Val(Left(N3, 1))
If V > 0 Then C3 = " " + Number(V) + " hundred"
V = Val(Right(N3, 2))
If V = 0 Then Return
If C3 <> "" Then C3 = C3 + " and"
If V <= 19 Then C3 = C3 + " " + Number(V): Return
V = Val(Mid(N3, 2, 1))
C3 = C3 + " " + Tens(V - 2)
V = Val(Right(N3, 1))
If V = 0 Then Return
C3 = C3 + " " + Number(V)
Return

End Function


' This will evaluate all the variables within <statement>
' Bob=2
' Tom=3
' Test="Bob Tom"
' ConvertVariables(Test)
' Test becomes: "3 2"
Public Sub ConvertVariables(Statement)
  
    
    Call GetVariables
    For N = 0 To VarBufferSize - 1
    If VarName(N) = "" Then GoTo GetNext
    If Statement Like "*" + VarName(N) + "*" Then V = Trim(Str(VarValue(N))): Call ReplaceWords(VarName(N), V, Statement)
    If Statement Like "*" + VarName(N) + "s*" Then V = Trim(Str(VarValue(N))): Call ReplaceWords(VarName(N) + "s", V, Statement)
    If Statement Like "*" + VarName(N) + "es*" Then V = Trim(Str(VarValue(N))): Call ReplaceWords(VarName(N) + "es", V, Statement)
GetNext:
    Next N
End Sub


Public Sub GetMouse(X, Y)
    Dim pa As POINTAPI
    A = GetCursorPos(pa)
    X = pa.X
    Y = pa.Y
End Sub

'
Public Sub GetVariables()
'    VarName = Array("Pigs", "Cows", "Dogs", "Cats")
'    VarValue = Array(10, 20, 30, 40)
'    VarBufferSize = 4

End Sub









Public Sub SaveAllMemory()

Call StoreMatchWords
Call StoreWontSayData

GoSub EncryptMemory


On Error GoTo FileError
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "Memory.Dat" For Output As #1
Print #1, Memory
Close #1

Exit Sub


EncryptMemory:
Memory = Replace(Memory, "a", Chr(142), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "e", Chr(143), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "i", Chr(144), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "o", Chr(145), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "u", Chr(146), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "s", Chr(147), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, "t", Chr(148), 1, -1, vbBinaryCompare)
Memory = Replace(Memory, " ", Chr(149), 1, -1, vbBinaryCompare)
Return







FileError:
Msg = "There was a problem trying to save my memory!" ' Define message.
Style = vbOKOnly + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Error!!"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)



End Sub


' This will add a variable to the variables list
' Examples:
' SetVariable "Cows=20"
' SetVariable "Pigs=20"
' SetVariable "Animals=Pigs"
'
' If there are no errors then <statement> will be returned with a null string.

Public Sub SetVariable(Statement)
If Statement Like "[a-z]*=*" Then GoTo DoIt
Exit Sub



DoIt:
Call ChangeYourForMyEtc(Statement)
Value = WordsBetween("=", "", Statement)
Result = SuperSum(Value)
If Result = "" Then Exit Sub
If Result Like "*[a-z]*" Then Exit Sub
Name = WordsBetween("", "=", Statement)
Call AddVariable(Name, Result)
Statement = ""

End Sub

' Extracts the date from <Message>
Public Function SuperDate(Message)
If Message = "" Then Exit Function

If Message Like "*#/#*" Then GoTo Maybe

For I = 1 To 12
If Message Like "* days in " + MonthName(I, True) + "*" Then Exit Function
If Message Like "* days of " + MonthName(I, True) + "*" Then Exit Function
If Message Like "*" + MonthName(I, True) + "*" Then GoTo Maybe
Next

Exit Function



Maybe:
MyDate = WordsToNumbers(Message)

Call ReplaceWords("first", "1", MyDate)
Call ReplaceWords("twenty second", "22", MyDate)
Call ReplaceWords("thirty second", "32", MyDate)
Call ReplaceWords("second", "2", MyDate)

Call ReplaceWords("1st", "1", MyDate)
Call ReplaceWords("2nd", "2", MyDate)
Call ReplaceWords("3rd", "3", MyDate)
Call ChangeAll("th ", " ", MyDate)
Call ChangeAll(".", "-", MyDate)
Call ChangeAll(",", " ", MyDate)
Call ReplaceWords(" o'clock", ":00", MyDate)

ExtractedDate = ""

Loop1:
MonthNumber = 0
Word = ExtractWord(MyDate)
If Word <> "" Then GoTo CheckIfDate
MyDate = ExtractedDate

If IsDate(MyDate) Then GoTo Done
MyDate = "1 " + MyDate
If IsDate(MyDate) Then GoTo Done
SuperDate = ""
Exit Function
Done:
SuperDate = Format(MyDate, "ttttt d/mm/yyyy")
Exit Function



CheckIfDate:
If Word Like "*[/-]*[/-]*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#:#*" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#pm" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like "*#am" Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Val(Word) Then Number = Val(Word): ExtractedDate = ExtractedDate + Str(Number) + " ": GoTo Loop1

GetMonthName:
If Word Like MonthName(MonthNumber + 1, True) Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
If Word Like MonthName(MonthNumber + 1, False) Then ExtractedDate = ExtractedDate + Word + " ": GoTo Loop1
MonthNumber = MonthNumber + 1: If MonthNumber = 12 Then GoTo Loop1
GoTo GetMonthName



End Function

' This will return the sum of a statement like:
' Statement="2+5*3"
' Result=SimpleSum(statement)
' Note:- It can't handle brackets.
Public Function SimpleSum(Statement)
    If Statement Like "*[+*/-]*" Then GoTo Calculate
    SimpleSum = Statement
    Exit Function
    
Calculate:

    Mode = "+"


    P = 1

    SimpleSum = 0
Loop1:
    Value = ""
    If P > Len(Statement) Then SimpleSum = LTrim(Str(SimpleSum)): Exit Function
    C = Mid(Statement, P, 1)
    If C = "." Then GoTo Loop2
    If C Like "[0-9]" Then GoTo Loop2
    If C Like "[+*/-]" Then Mode = C
    P = P + 1: GoTo Loop1


Loop2:
    C = Mid(Statement, P, 1)
    If C = "." Then Value = Value + C: GoTo GetNextDigit
    If C Like "[0-9]" Then Value = Value + C Else GoTo DoIt
GetNextDigit:
    P = P + 1: If P > Len(Statement) Then GoTo DoIt
    GoTo Loop2

DoIt:
    If Mode = "+" Then SimpleSum = SimpleSum + Val(Value)
    If Mode = "-" Then SimpleSum = SimpleSum - Val(Value)
    If Mode = "/" Then
    If Val(Value) = 0 Then SimpleSum = "Error:- Divide by zero": Exit Function
    SimpleSum = SimpleSum / Val(Value)
    End If
    If Mode = "*" Then SimpleSum = SimpleSum * Val(Value)
    GoTo Loop1


End Function


' This will strip all punctuation from a statement, leaving only letters, numbers and spaces.
Public Sub StripPunctuation(Statement)
NS = ""
Position = 1
Loop1:
If Position > Len(Statement) Then Statement = NS: Exit Sub
C = Mid(Statement, Position, 1)
Position = Position + 1
If C Like "#" Then NS = NS + C: GoTo Loop1
If C Like "[a-z]" Then NS = NS + C: GoTo Loop1
If C = " " Then NS = NS + C: GoTo Loop1
GoTo Loop1

End Sub

'This will return the sum of a statement.
'The statement must be a string.
'The statement may include brackets.
'Example:
' Statement="8+2+(6*3)"
' Result=Sum(Statement)
' Result is:- 28
Public Function Sum(Statement)


Statement = Statement + ")"
Statement = "(" + Statement

Sum = 0
Loop1:
If Statement Like "*[()]*" Then GoTo ProcessBrackets
GoTo BracketsDone

ProcessBrackets:

' Find an open bracket
FindOpen:
    P = Len(Statement)
FOLp1:
    C = Mid(Statement, P, 1)
    If C = "(" Then Op = P: GoTo FindClose
    P = P - 1: If P = 0 Then Statement = "(" + Statement: P = 1: GoTo FindClose
    GoTo FOLp1


' Find close bracket
FindClose:
FCLp1:
    If Mid(Statement, P, 1) = ")" Then CP = P: GoTo Extract
    P = P + 1: If P > Len(Statement) Then GoTo CantFindClose
    GoTo FCLp1


CantFindClose:
    Mid(Statement, Op, 1) = " "
    GoTo BracketsDone

Extract:
    SubStatement = Mid(Statement, Op + 1, (CP - Op) - 1)
    Mid(Statement, Op, 1) = "T"
    Mid(Statement, CP, 1) = "T"
    Result = SimpleSum(SubStatement)
    If Result Like "*Divide by zero*" Then Sum = Result: Exit Function
    ReplaceCharacters "T" + SubStatement + "T", Result, Statement

    GoTo Loop1
BracketsDone:
    Sum = Statement



End Function

Public Function Capitalise(Message)
If Message = "" Then Exit Function
FirstLetter = Left(Message, 1)
FirstLetter = UCase(FirstLetter)
Message = Right(Message, Len(Message) - 1)
Message = (FirstLetter + Message)
End Function

' This will replace all <Words1> in <statement> with <Words2>
Public Sub ChangeAll(Words1, Words2, Statement)
Do
Temp = Statement
ReplaceCharacters Words1, Words2, Statement
If Statement = Temp Then Exit Sub
Loop
End Sub

' This will replace Characters1 within the string Text with Characters2
Public Sub ReplaceCharacters(Characters1, Characters2, Text)
Position = InStr(1, Text, Characters1, 1)
If Position = 0 Then Exit Sub
Text2 = Left(Text, Position - 1)
Text3 = Mid(Text, Position + Len(Characters1))
Text = Text2 + Characters2 + Text3

End Sub




' This will replace the <words1> within <Text> with <words2>
'
' Examples:
' Msg="My cat is outside"
' ReplaceWords "cat","rabbit",Msg
' Msg Becomes:
' Msg="My rabbit is outside"
'
' Msg="My catflap is broke"
' ReplaceWords "cat","rabbit",Msg
' Msg Remains:
' Msg="My catflap is broke"
' The 'cat' isnt changed this time because it isnt a single lone word
'
Public Sub ReplaceWords(Words1, Words2, Text)
If Text = "" Then Exit Sub
If Words1 = "" Then Exit Sub
Position = 1
Loop1:
Position = InStr(Position, Text, Words1)
If Position = 0 Then GoTo Done
If Position = 1 Then GoTo CheckEnd
If Mid(Text, Position - 1, 1) Like "[a-z]" Then Position = Position + 1: GoTo Loop1
CheckEnd:
If Mid(Text, Position + Len(Words1), 1) Like "[a-z]" Then Position = Position + 1: GoTo Loop1
Text2 = Left(Text, Position - 1)
Text3 = Mid(Text, Position + Len(Words1))
If Words2 <> "" Then GoTo P1
C = Right(Text2, 1) + Left(Text3, 1)
If C Like " [!a-z]" Then If C Like " [!0-9]" Then Text2 = Left(Text2, Len(Text2) - 1): GoTo P1

P1:
Text = Text2 + Words2 + Text3
Position = Position + Len(Words2)
GoTo Loop1
Done:
Text = Trim(Text)
End Sub



' This will return a value indicating the current tense of a statement.
' The values returned:
'  1 - Future
'  0 - Present
' -1 - Past
'
' Example:
' > Message="Bob will go out"
' > Print Tense(Message)
' 1
'

Public Function Tense(Statement)

Tense = Present
If Statement Like "*[!a-z]is[!a-z]*" Then Tense = Present
If Statement Like "is[!a-z]*" Then Tense = Present
If Statement Like "isn't" Then Tense = Present
If Statement Like "*do*" Then Tense = Present
If Statement Like "*will*" Then Tense = Future
If Statement Like "*did*" Then Tense = Past
If Statement Like "*was*" Then Tense = Past
If Statement Like "*can*" Then Tense = Present
If Statement Like "*would*" Then Tense = Past
If Statement Like "*went*" Then Tense = Past
If Statement Like "*may*" Then Tense = Present
If Statement Like "*might*" Then Tense = Past
If Statement Like "*could*" Then Tense = Past
If Statement Like "*should*" Then Tense = Past
If Statement Like "*shall*" Then Tense = Present
If Statement Like "*were*" Then Tense = Past
If Statement Like "*does*" Then Tense = Present
If Statement Like "*are*" Then Tense = Present


End Function



Public Sub TestDim(Blah, Text)
If Blah > 2 Then GoTo Continue

'Dim words(10)
Static Words(10)
Words(1) = "Hello"
Words(2) = "There"
Continue:
Text = Words(1) + " " + Words(2)

End Sub

Private Sub TieAlternatives(Statement)

If MatchWords = "" Then Exit Sub
Statement = " " + Statement
P = 0
Loop1:
P = InStr(P + 1, MatchWords, ">'", vbBinaryCompare)
If P = 0 Then GoTo Done
P2 = InStr(P + 2, MatchWords, "=", vbBinaryCompare)
If P2 = 0 Then GoTo Loop1
P = P + 2: P2 = P2 - 1
Words = Mid(MatchWords, P, P2 - P)
If Statement Like "*[!a-z]Words[!a-z]*" Then GoTo TieEm
GoTo Loop1


TieEm:
TiedWords = Words
Call ChangeAll(" ", "-", TiedWords)
Call ChangeAll(Words, TiedWords, Statement)
GoTo Loop1

Done:
Statement = Trim(Statement)





End Sub

' This will scan a statement for numbers like "one hundred and twenty" and change them into numbers like "120"
Public Function WordsToNumbers(Statement)
If Statement = "" Then Exit Function

OriginalStatement = Statement


GoSub ConvertNumbers



Call ConvertFractionWordsToNumbers(Statement)

' Convert single numbers like 'three' into '3'
Num0to19 = Array("zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
For N = 19 To 0 Step -1
Loop1:
Temp = Statement
ReplaceWords Num0to19(N), CStr(N), Statement
If Temp <> Statement Then GoTo Loop1
Next N

'Convert numbers like 'twenty' to '20'
'Also, add a y to the end of the conveted numbers for later processing
Tens = Array("Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
For N = 0 To 7
Loop2:
Temp = Statement
ReplaceCharacters Tens(N), CStr((N * 10) + 20) + "y", Statement
If Temp <> Statement Then GoTo Loop2
Next N


' Now any number with y on the end and followed by a single number will be merged together.
' So '20y 1' will become '21'.

Statement = Statement + " "
Loop3:
For N = 1 To Len(Statement) - 1
If Mid(Statement, N, 2) = "0y" Then GoTo Found
Next N
GoTo Done
Found:
If Mid(Statement, N + 2, 2) Like "[0-9] " Then
ReplaceCharacters "0y", "", Statement
GoTo Loop3
End If
If Mid(Statement, N + 2, 3) Like " [0-9][!0-9]" Then
ReplaceCharacters "0y ", "", Statement
GoTo Loop3
End If
ReplaceCharacters "0y", "0", Statement
GoTo Loop3
Done:
Statement = Left(Statement, Len(Statement) - 1)

' This section will swap various worded numbers
ChangeAll "a hundred", "100", Statement
    ' Change any single ' hundred' words into '00'
    ' So:- '1 hundred' becomes '100'
ChangeAll " hundred", "00", Statement
    ' Do the same as above but with a thousand
ChangeAll "a thousand", "1000", Statement
ChangeAll " thousand", "000", Statement
    ' As above but with a million
ChangeAll "a million", "1000000", Statement
ChangeAll " million", "000000", Statement
    ' Now a billion
ChangeAll "a billion", "1000000000", Statement
ChangeAll " billion", "000000000", Statement
ChangeAll "a trillion", "1000000000000", Statement
ChangeAll "trillion", "000000000000", Statement


' Now for the hard bit.
' This will convert numbers like '100 and 20' into '120'
 
' Start by extracting the '<number> and <number>' part
Loop4:
Temp = Statement
Position = 1
Loop5:
Position = InStr(Position, Statement, "00 and ")
If Position = 0 Then GoTo Finished
If Mid(Statement, Position + 7, 1) Like "![0-9]" Then Position = Position + 1: GoTo Loop5
P = Position - 1
Loop6:
If Mid(Statement, P, 1) Like "[0-9]" Then
If P > 1 Then P = P - 1: GoTo Loop6
End If
CStart = P
P = Position + 8
Loop7:
If Mid(Statement, P, 1) Like "[0-9]" Then
P = P + 1: If P < Len(Statement) Then GoTo Loop7
End If
CEnd = P
Calc = Mid(Statement, CStart, CEnd - CStart)
' Calc now contains the '<number> and <number>' part
' now add the two numbers together
Value1 = Val(Calc)
VT = Mid(Statement, Position + 7)
Value2 = Val(VT)
Total = Str(Value1 + Value2)
' Swap the '<number> and <number>' part with the result
ReplaceCharacters Calc, Total, Statement
' Ok, its obvious we have processed something so go and check to see if there are anymore.
If Statement <> Temp Then GoTo Loop4


Finished:
WordsToNumbers = Statement
Statement = OriginalStatement
Exit Function






ConvertNumbers:
If Statement + " " Like "*s[!a-z]*" Then GoTo ConvEm
Return
ConvEm:
Numbers1 = Array("zeros", "ones", "twos", "threes", "fours", "fives", "sixes", "sevens", "eights", "nines", "tens", "elevens", "twelves", "thirteens", "fourteens", "fifteens", "sixteens", "seventeens", "eighteens", "nineteens", "twenties", "thirties", "forties", "fifties", "sixties", "seventies", "eighties", "nineties", "hundreds", "thousands", "millions", "billions", "trillions")
For N = 0 To 32
If " " + Statement + " " Like "*[!a-z]" + Numbers1(N) + "[!a-z]*" Then GoSub Convert
Next N
Return

Convert:
Numbers2 = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred", "thousand", "million", "billion", "trillion")
Call ReplaceWords(Numbers1(N), Numbers2(N), Statement)
Return













End Function



' This will place all the words within a string into an array.
' Example:
' Call ExtractWords(String,MyArray())
Public Sub ExtractWords(Item, ExtractedWords())
Word = ""
Position = 1
Loop1:
 C = Mid(Item, Position, 1)
 Position = Position + 1
 If C Like "[a-z]" Then Word = Word + C: GoTo Loop1
 If C Like "[0-9]" Then Word = Word + C: GoTo Loop1
 If C = "'" Then Word = Word + C: GoTo Loop1

StoreWord:
 If Word = "" Then GoTo CheckOperators
 ExtractedWords(N) = Word: N = N + 1: Word = ""
 
CheckOperators:
' added the "()" on 31/8/01 (might effect other procedures)
 If C Like "[()+-/*=]" Then ExtractedWords(N) = C: N = N + 1
 If Position <= Len(Item) Then GoTo Loop1
 ExtractedWords(N) = ""



End Sub

' This returns the next memory item that contains all the words within the specified sentence
' Note:- It doesn't care about the order of the words



Public Function ItemLike(Sentence)
ItemLike = ""
    Dim Words(20)
    NextV2 = ""

    Call ExtractWords(Sentence, Words())
    If Words(0) = "" Then Exit Function
        
'    If Words(1) = "" Then GoTo SearchWord
'    If Len(Words(0)) = 1 Then Temp = Words(1): Words(1) = Words(0): Words(0) = Temp


    N = 1
    MainWord = Words(0)
Loop0:
    If Words(N) = "" Then GoTo SearchWord
    If Len(Words(N)) > Len(MainWord) Then MainWord = Words(N)
    N = N + 1
    GoTo Loop0




SearchWord:
    Statement = ItemWith(MainWord)
    If Statement = "" Then Exit Function
    N = 0
Loop1:
    If Words(N) = "" Then GoTo Done

    If ItemSearchStyle = Careless Then GoTo JustCheckLeftSide
    If Statement Like "*[!a-z]" + Words(N) + "[!a-z]*" Then N = N + 1: GoTo Loop1
    If Statement Like Words(N) + "[!a-z]*" Then N = N + 1: GoTo Loop1
    If Statement Like "*[!a-z]" + Words(N) Then N = N + 1: GoTo Loop1
    GoTo SearchWord

JustCheckLeftSide:
    If Statement Like "*[!a-z]" + Words(N) + "*" Then N = N + 1: GoTo Loop1
    If Statement Like Words(N) + "*" Then N = N + 1: GoTo Loop1
    GoTo SearchWord


Done:
    ItemLike = Statement
    Exit Function



End Function


' Returns a memory-item from within Bob's memory that contains the specified item.
' The search method must be specified before using it.
'
' Example:
' SearchMethod=ExactMatch
' Item=NextItemContaining("colour red")
' This would return an item containing:- "colour red"
'
' SearchMethod=AnyOrder
' Item=NextItemContaining("colour red")
' This would return an item containing:- "colour" and "red" or "red" and "colour"
Public Function NextItemContaining(Item)
    NextItemContaining = ""
    If ItemSearchMethod = AnyOrder Then NextItemContaining = ItemLike(Item): Exit Function
    If ItemSearchMethod = WordForWord Then NextItemContaining = ItemWith(Item): Exit Function
    If ItemSearchMethod = InOrder Then NextItemContaining = ItemLikeInOrder(Item): Exit Function


End Function



Public Sub CollectThings(Thing, Items())
    If Thing = "" Then Exit Sub
Look:
    RestartItemSearch
    ItemSearchMethod = WordForWord
    ItemSearchStyle = Careless

    N = 0
L1: Statement = NextItemContaining(Thing)
    If Statement = "" Then GoTo Finished
    If Statement Like "* is a " + Thing + "[!a-z]*" Then GoTo GetIsA
    If Statement Like "* is an " + Thing + "[!a-z]*" Then GoTo GetIsAn
    If Statement Like "* are " + Thing + "s" + "[!a-z]*" Then GoTo GetAre
    If Statement Like "* are all " + Thing + "s" + "[!a-z]*" Then GoTo GetAre
    GoTo L1

GetIsA:
    Items(N) = WordsBetween("", "is a", Statement)
    N = N + 1
    GoTo L1
    

GetIsAn:
    Items(N) = WordsBetween("", "is an", Statement)
    N = N + 1
    GoTo L1
    
GetAre:
    Call ReplaceCharacters(" and ", ",", Statement)
    GoSub ExtractList
    If List = "" Then GoTo L1
    Dim Words(50)
    Call ExtractWords(List, Words())
    W = 0
L2: If Words(W) = "" Then GoTo L1
    If Words(W) Like "[a-z]*" Then Items(N) = Words(W): N = N + 1
    W = W + 1
    GoTo L2


ExtractList:
    List = ""
    P = InStr(1, Statement, ",")
    If P = 0 Then Return
    P2 = InStrRev(Statement, " ", P)
    If P2 = 0 Then P2 = 1
    P3 = InStr(P, Statement, " ")
    If P3 = 0 Then Return
    List = Mid(Statement, P2, P3 - P2)
    Return





Finished:
    Items(N) = ""



End Sub

' This will get the statement that is within the area of <position> from Memory.
' It extracts the statement from memory and adds it to the end of memory. <<--- This feature was later remmed out
' This is so that it will be faster to access statement the next time. <<------      "                      "
'
Public Function LocalItem(Position)


GetEnd:
ItemEnd = InStr(Position, Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), vbBinaryCompare)
If ItemEnd <= 1 Then LocalItem = "": Exit Function


GetStart:
ItemStart = InStrRev(Memory, Chr(13) + Chr(10) + Chr(13) + Chr(10), Position, vbBinaryCompare)
If ItemStart = 0 Then ItemStart = 1: GoTo GotIt
ItemStart = ItemStart + 4


GotIt:
Position = ItemStart


LocalItem = Mid(Memory, ItemStart, ItemEnd - ItemStart)
MemoryPosition = ItemStart


'FullStopPosition = InStr(position + 1, Memory, ".")
'LocalSentence = Mid(Memory, position + 1, FullStopPosition - position)

' replace any return codes within the sentence with spaces and then trim the leading spaces
'For L = 1 To Len(LocalSentence)
'If Asc(Mid(LocalSentence, L, 1)) < 32 Then Mid(LocalSentence, L, 1) = " "
'Next L
'LocalSentence = Trim(LocalSentence)

'there is no way.




' I realised that bringing all accessed data to the forefront of memory wasn't going to work so I took these two lines out.
'Memory = Left(Memory, position - 1) + Right(Memory, Len(Memory) - (position + Len(LocalItem)) - 3)
'Call AddMemoryItem(LocalItem)

End Function


' This will reset the memory searchers start position
Public Sub RestartItemSearch()
MemoryPosition = Len(Memory)
End Sub


'This will return the sum of a statement.
' Possible statements:
' Result=SuperSum("Eight plus two plus two hundred")
' Result=SuperSum("10+(6*3)")
' Result=SuperSum("Two plus (three times three)")
' Result=SuperSum("Two tens")
' Result=SuperSum("two thirds of nine?")
Public Function SuperSum(Statement)


Message = WordsToNumbers(Statement)

Call ConvertVariables(Message)
If Message Like "*[0-9]*" Then GoTo CarryOn
SuperSum = ""
Exit Function


CarryOn:
Call ChangeAll(" and a half", ".5", Message)

Call ConvertFractions(Message)

Call OperatorWordsToSymbols(Message)



SuperSum = Message

' Note: I have moved these commented lines further down.
'If Message Like "*#[+-/*]#*" Then GoTo Calc
'Exit Function
'Calc:

Loop1:
If Message Like "*[[]*]*" Then
ReplaceCharacters "[", "(", Message
ReplaceCharacters "]", ")", Message
GoTo Loop1
End If
Loop2:
If Message Like "*{*}*" Then
ReplaceCharacters "{", "(", Message
ReplaceCharacters "}", ")", Message
GoTo Loop2
End If
Loop3:
If Message Like "*<*>*" Then
ReplaceCharacters "<", "(", Message
ReplaceCharacters ">", ")", Message
GoTo Loop3
End If


Statement = ""
For P = 1 To Len(Message)
C = Mid(Message, P, 1)
If C Like "[()+/*-]" Then Statement = Statement + C
If P < Len(Message) Then If C = "." Then Statement = Statement + C
If C Like "#" Then Statement = Statement + C
Next P


If Statement Like "*#+-#*" Then GoTo Calc
If Statement Like "*#[+-/*]#*" Then GoTo Calc
If Statement Like "*#*" Then SuperSum = Statement: Exit Function
SuperSum = ""
Exit Function



Calc:
SuperSum = Sum(Statement)

Exit Function

DivideError:
SuperSum = "Error:- Divide by zero!"
Exit Function




End Function

' This will extract the words within the <Text> between <word1> and <word2>.
' Example:
' > Text="The rain in spain."
' > Print WordsBetween("the "," in ",Text)
' rain
'
Public Function WordsBetween(Word1, Word2, Text)
If Word1 = "" Then Position = 1: GoTo Lp1
Position = InStr(1, Text, Word1)
If Position = 0 Then Position = 1
Position = Position + Len(Word1)
Lp1:
If Word2 = "" Then position2 = Len(Text) + 1: GoTo Lp2
position2 = InStr(Position, Text, Word2)
If position2 = 0 Then position2 = Len(Text) + 1
Lp2:
WordsBetween = Mid(Text, Position, position2 - Position)
WordsBetween = Trim(WordsBetween)
End Function


