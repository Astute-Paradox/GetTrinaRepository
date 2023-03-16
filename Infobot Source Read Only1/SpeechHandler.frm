VERSION 5.00
Begin VB.Form Speaker 
   Caption         =   "Speech Engines:"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "SpeechHandler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox ListWindow 
      Height          =   5325
      ItemData        =   "SpeechHandler.frx":0442
      Left            =   120
      List            =   "SpeechHandler.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Speaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Text to speech handler.

Public InMessage
Public OutMessage
Public MessageToSay

Private CurrentMessage
Private CurrentMessagePosition
Private MessagePart
Private PauseTime

Public NoEnginesAvailable
Public EngineActive
Private SelectError

Private EngineSelected
Private EngineNumber


Const Sapi4 = 1
Const Sapi5 = 2
Private EngineVersion

Dim TTS As TextToSpeech     ' Sapi 4 text to speech object
Attribute TTS.VB_VarHelpID = -1
Dim Voice As SpVoice        ' Sapi 5 text to speech object


Option Compare Text

Private LastUsersMessageTime
Private LastBotsMessageTime


Public Sub InitializeControl()
If EngineActive Then Exit Sub

If NoEnginesAvailable Then Exit Sub


' Try and initialize the sapi 5 TTS control
Err.Clear
On Error Resume Next
Set Voice = New SpVoice
If Err.Number <> 0 Then Err.Clear: GoTo TrySapi4
EngineActive = True: EngineVersion = Sapi5: Debug.Print "SAPI 5 in use."
Exit Sub


' Try and initialize the sapi 4 TTS control
TrySapi4:
Set TTS = New TextToSpeech
If Err.Number <> 0 Then Err.Clear: NoEnginesAvailable = True: Exit Sub
EngineActive = True: EngineVersion = Sapi4: Debug.Print "SAPI 4 in use."
Exit Sub



End Sub

Public Sub Speak(TheMessage)
If EngineActive = False Then Exit Sub

Message = TheMessage
Call ReplaceWords("Answerpad", "Answer pad", Message)
Message = Replace(Message, ",", ", ", 1, -1, vbBinaryCompare)


On Error Resume Next

If EngineVersion = Sapi5 Then GoTo SpeakWith5


TTS.StopSpeaking
TTS.Speak Message
If Err.Number <> 0 Then Debug.Print "TTS Speak Error!"
Exit Sub


SpeakWith5:
Voice.Speak Message, SVSFPurgeBeforeSpeak + SVSFlagsAsync
If Err.Number <> 0 Then Debug.Print "TTS Speak Error! (Sapi 5)"
Exit Sub




End Sub

Private Sub SayNextMessagePart()
Static LastSeenTime
Static LastSeenPart
If BotsMessagePartTime <> LastSeenTime Then GoTo SayIt
' The above line sometimes misses a new message if the new message appeared with the same second as the last
' The following line will help to catch it.
If BotsMessagePart <> LastSeenPart Then GoTo SayIt
Exit Sub


SayIt:
LastSeenTime = BotsMessagePartTime
LastSeenPart = BotsMessagePart
Call Speak(BotsMessagePart)
Exit Sub





End Sub

' This is the text to speech handler.
Public Sub Update()


Call InitializeControl

'Call SelectMyFavouriteEngine



OutMessage = ""
InMessage = ""


CheckMessages:
InMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime

'If BotsMessageTime <> LastBotsMessageTime Then Call Speak(BotsMessage): LastBotsMessageTime = BotsMessageTime
'If BotsMessageTime <> LastBotsMessageTime Then LastBotsMessageTime = BotsMessageTime

Call SayNextMessagePart

Call Msg_SelectSpeechEngine: If OutMessage <> "" Then GoTo Done
Call Msg_OpenSpeechEngine: If OutMessage <> "" Then GoTo Done

Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage): LastBotsMessageTime = BotsMessageTime



End Sub

Public Sub SelectMyFavouriteEngine()
If EngineActive = False Then Exit Sub
Static AlreadyDone

If AlreadyDone Then Exit Sub
AlreadyDone = True

On Error GoTo Oops

If EngineVersion = Sapi4 Then GoTo SelectSapi4Voice

Exit Sub



SelectSapi4Voice:
EnginesAmount = TTS.CountEngines
For I = 1 To EnginesAmount
EngineName = TTS.ModeName(I)
If EngineName Like "*Adult Male*" Then Call UseEngine(I): Exit Sub
Next I
Exit Sub


Oops:
Debug.Print "FAVE TTS ERROR!!"
End Sub




Private Sub Msg_SelectSpeechEngine()
If EngineActive = False Then Exit Sub
If EngineSelected Then EngineSelected = False: GoTo ActivateEngine
If InMessage Like "*show *speech engine*" Then ShowSpeechEngines: OutMessage = "Ok.": Exit Sub
If InMessage Like "*select *speech engine*" Then GoTo GetNumber
If InMessage Like "*use *speech engine*" Then GoTo GetNumber
If InMessage Like "Select #*" Then If Speaker.Visible Then GoTo GetNumber
If InMessage Like "*Hide *speech engine*" Then GoTo HideEm
If InMessage Like "*Close *speech engine*" Then GoTo HideEm
If Speaker.Visible Then GoTo CheckForNumberOnly
Exit Sub


HideEm:
Speaker.Visible = False
OutMessage = "Done."
Exit Sub


CheckForNumberOnly:
NewMessage = WordsToNumbers(InMessage)
If NewMessage Like "*[a-z]*" Then Exit Sub
If NewMessage Like "#*" Then InMessage = NewMessage: GoTo GetNumber
Exit Sub


GetNumber:
Message = InMessage
EngineNumber = SuperSum(Message)

If EngineNumber = "" Then OutMessage = "Enter the number of the speech engine.": Call ShowSpeechEngines: Exit Sub


EngineNumber = Val(EngineNumber)

ActivateEngine:
UseEngine EngineNumber
If SelectError Then OutMessage = "There was an error trying to select that engine.": Exit Sub

OutMessage = "Speech engine" + Str(EngineNumber) + " has been selected."



End Sub


Private Sub Msg_OpenSpeechEngine()
If EngineActive = False Then Exit Sub
If InMessage Like "*Open *speech*" Then GoTo OpenIt
If InMessage Like "*Open *engine*" Then GoTo OpenIt
Exit Sub

OpenIt:
On Error GoTo Error
TTS.LexiconDlg Interface.hWnd, "Test"

'Index = Interface.TextToSpeech1.Find("")
'Debug.Print Interface.TextToSpeech1.ProductName(Index)
'Debug.Print Interface.TextToSpeech1.ModeName(Index)
OutMessage = "Ok."
Exit Sub

Error:
OutMessage = "There seems to have beeen a problem trying to open the speech engine's options."

End Sub


Public Sub UseEngine(EngineNumber)
If EngineActive = False Then Exit Sub

SelectError = False
On Error Resume Next

If EngineVersion = Sapi5 Then GoTo UseSapi5Voice

TTS.Select EngineNumber
If Err.Number <> 0 Then Err.Clear: SelectError = True
Exit Sub


UseSapi5Voice:
 Set Voice.Voice = Voice.GetVoices().Item(EngineNumber - 1)
 If Err.Number <> 0 Then Err.Clear: SelectError = True
Exit Sub



End Sub


Public Sub ShowSpeechEngines()
If EngineActive = False Then Exit Sub

Err.Clear
On Error GoTo Oops

If EngineVersion = Sapi4 Then GoTo GetSapi4Voices

GetSapi5Voices:
    I = 1
    ListWindow.Clear
    Dim Token As ISpeechObjectToken
    For Each Token In Voice.GetVoices
        ListWindow.AddItem Str(I) + ". " + Chr(9) + Token.GetDescription
    I = I + 1
    Next
    GoTo ShowEm




GetSapi4Voices:
EnginesAmount = TTS.CountEngines
'Interface.TextToSpeech1.GeneralDlg Interface.hWnd, "Test"
'Index = Interface.TextToSpeech1.Find("")
ListWindow.Clear
For I = 1 To EnginesAmount
EngineName = TTS.ModeName(I)
ListWindow.AddItem Str(I) + ". " + Chr(9) + EngineName
Next I



ShowEm:
Speaker.Left = Interface.Left + ((Interface.Width - Speaker.Width) / 2)
Speaker.Top = Interface.Top + Interface.Height
If Interface.Top > (Screen.Height / 2) Then Speaker.Top = (Interface.Top - Speaker.Height)
Speaker.Show
Interface.UsersWindow.SetFocus

Exit Sub

Oops:

End Sub

Private Sub ListWindow_Click()
EngineNumber = ListWindow.ListIndex + 1
EngineSelected = True

End Sub


