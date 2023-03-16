VERSION 5.00
Begin VB.Form BrainCore 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   3360
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "BrainCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub HandleShutdown()
If IsDate(ShutdownTime) Then GoTo CheckShutdownTime
Exit Sub

CheckShutdownTime:
If Time > ShutdownTime Then If Time < (ShutdownTime + #12:00:30 AM#) Then GoTo Shutdown

Exit Sub


Shutdown:
Call SaveOptions
If SaveMemoryAllowed Then Call SaveAllMemory
Call SaveNewKnowledge

End




End Sub

Private Sub LoadOptions()

InitTime = Timer

On Error Resume Next
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "Options.Cfg" For Input As #1
Size = LOF(1)
Options = Input(Size, #1)
CloseIt:
Close #1



If Options Like "*MSAgent=On*" Then Call MSAgent.EnableMSAgent("") Else Call SetBotsMessage(" "): Call Speaker.Update

If Options Like "*SpeechRecognition=On*" Then Call SpeechRecognizer.EnableEngine
If Options Like "*AutoReturnTime*" Then GoSub GetReturnValue
GoSub GetFactFileInfo


Exit Sub


' Read the last fact-file that the bot was in the process of reading.
GetFactFileInfo:
PosWords = "ReadingPosition"
P = InStr(1, Options, PosWords)
If P = 0 Then Return
ValArea = Mid(Options, P + Len(PosWords) + 1, 15)
Value = Val(ValArea)

PathWords = "FactFilePath"
P = InStr(1, Options, PathWords)
If P = 0 Then Return
P2 = InStr(P, Options, Chr(13) + Chr(10))
If P2 = 0 Then P2 = Len(Options)
P = P + Len(PathWords) + 1
FullPath = Mid(Options, P, P2 - P)


Open FullPath For Input As #1
Size = LOF(1)
Facts = Input(Size, #1)
Close #1

BotReadingFile = Dir(FullPath)
BotReadingFullpath = FullPath
BotReadingPosition = Value
Return




GetReturnValue:
Pos = InStr(1, Options, "AutoReturnTime")
ReturnValue = Mid(Options, Pos + 15, 5)
Interface.AutoReturnTime = Val(ReturnValue)
Return







End Sub


Private Sub SaveOptions()
Options = ""
If MSAgent.AgentActive Then Options = Options + "MSAgent=On" + Chr(13) + Chr(10)
If SpeechRecognizer.UseSpeechRecognition Then Options = Options + "SpeechRecognition=On" + Chr(13) + Chr(10)

ReturnTime = Str(Interface.AutoReturnTime)
Options = Options + "AutoReturnTime=" + Trim(ReturnTime) + Chr(13) + Chr(10)
GoSub StoreFactFileInfo


On Error Resume Next
MyPath = App.Path: If MyPath Like "*[!\]" Then MyPath = MyPath + "\"
Open MyPath + "Options.Cfg" For Output As #1
Print #1, Options
Close #1
Exit Sub


StoreFactFileInfo:
If BotReadingPosition = 0 Then Return
If BotReadingFullpath = "" Then Return
Options = Options + "FactFilePath=" + BotReadingFullpath + Chr(13) + Chr(10)
Options = Options + "ReadingPosition=" + LTrim(Str(BotReadingPosition)) + Chr(13) + Chr(10)
Return




End Sub




Private Sub UpdateMenus()

Static PreAgentState
Static PreSRState
Static PreHistoryState
Static PreEasyPadState
Static PreSwearState
Static PreRambleState
Static PreReadState
Static PreSpeakState


If MSAgent.AgentActive <> PreAgentState Then PreAgentState = MSAgent.AgentActive: GoTo RebuildMenu
If SpeechRecognizer.Active <> PreSRState Then PreSRState = SpeechRecognizer.Active: GoTo RebuildMenu
If HistoryForm.Visible <> PreHistoryState Then PreHistoryState = HistoryForm.Visible: GoTo RebuildMenu
If EasyInput.Visible <> PreEasyPadState Then PreEasyPadState = EasyInput.Visible: GoTo RebuildMenu
If SwearingAllowed <> PreSwearState Then PreSwearState = SwearingAllowed: GoTo RebuildMenu
If RamblingAllowed <> PreRambleState Then PreRambleState = RamblingAllowed: GoTo RebuildMenu
If ReadingAllowed <> PreReadState Then PreReadState = ReadingAllowed: GoTo RebuildMenu
If Speaker.Visible <> PreSpeakState Then PreSpeakState = Speaker.Visible: GoTo RebuildMenu
Exit Sub


RebuildMenu:
Dim Cmd(10)

N = 0
If MSAgent.AgentActive Then Cmd(N) = "Disable Microsoft Agent" Else Cmd(N) = "Enable Microsoft Agent"
N = N + 1

If MSAgent.AgentActive Then Cmd(N) = "Select Agent character": N = N + 1

If MSAgent.AgentActive Then If MSAgent.Agent1.PropertySheet.Visible = False Then Cmd(N) = "Show Agent options": N = N + 1

If MSAgent.AgentActive = False Then If Speaker.Visible = False Then Cmd(N) = "Select Text to Speech engine": N = N + 1 Else If Speaker.Visible = True Then Cmd(N) = "Hide speech engines": N = N + 1

If SpeechRecognizer.Active Then Cmd(N) = "Disable speech recognition" Else Cmd(N) = "Enable speech recognition"
N = N + 1

If HistoryForm.Visible Then Cmd(N) = "Hide recent messages" Else Cmd(N) = "Show recent messages"
N = N + 1

If EasyInput.Visible Then Cmd(N) = "Hide Easy Pad" Else Cmd(N) = "Easy Pad"
N = N + 1

If SwearingAllowed Then Cmd(N) = "Do not swear" Else Cmd(N) = "Feel free to swear"
N = N + 1

If RamblingAllowed Then Cmd(N) = "Stop rambling" Else Cmd(N) = "Feel free to ramble"
N = N + 1

If ReadingAllowed Then If Facts <> "" Then Cmd(N) = "Abort reading"
N = N + 1


For N = 0 To 9
If Cmd(N) <> "" Then Interface.Cmd(N).Caption = Cmd(N): Interface.Cmd(N).Visible = True Else Interface.Cmd(N).Visible = False
Next N



End Sub



Private Sub Form_Load()

VersionNumber = LTrim(Str(App.Major)) + "." + LTrim(Str(App.Minor))
If App.Revision > 0 Then VersionNumber = VersionNumber + "." + LTrim(Str(App.Revision))
About1.Version.Caption = "Version " + VersionNumber
Splash.Version = "V" + VersionNumber

' The Open "file" For Input was previously Open "file" for Binary

'Open "WordsData.txt" For Input As #1
'Size = LOF(1)
'WordsData = Input(Size, #1)
'Close #1

StartTime = Timer

Splash.Show
Splash.Refresh

Call LoadMemory


Call LoadOptions

InitStartTime = Timer

Call SmartArse.InitialiseBot

Debug.Print "Init time:- "; Timer - InitStartTime


'Beep
'Splash.OkayShadow.Visible = True
'Splash.Okay.Visible = True
'Splash.Label1.Caption = "Click OK to start..."
'Splash.Refresh
'WaitLoop:
'DoEvents
'If Splash.Visible = True Then GoTo WaitLoop

Randomize
R = Int(Rnd * 5)
Dim Omsg(5)
Omsg(0) = "Shoot..."
Omsg(1) = "Fire away..."
Omsg(2) = "Ready for action..."
Omsg(3) = "I'm all yours.."
Omsg(4) = "Okay, lets do it..."
Splash.Label1.Caption = Omsg(R)


Timer1.Enabled = True
Interface.Top = (Splash.Top + Splash.Height)
Interface.Left = (Screen.Width / 2) - (Interface.Width / 2)
Call Interface.Update


Timer2.Enabled = True
Exit Sub





End Sub


Private Sub Timer1_Timer()




' Update the fact filer
Call FactFiler.Update

'Update the fact extractor
'Call FactExtractor.Update


Call Interface.Update


Call UpdateMenus


'Call Mirc.Update

Call SpeechRecognizer.Update

'Call Quake.Update


If MSAgent.AgentActive Then GoTo SkipSpeechHandler
Call Speaker.Update
SkipSpeechHandler:



'If MS-Agent is active then hide the Infobots output window.
If MSAgent.AgentActive Then Interface.BotsWindow.Visible = False Else Interface.BotsWindow.Visible = True
' Update the MS-Agent manager.
Call MSAgent.Update

'Update the Easy input handler.
Call EasyInput.Update

' Update the message-history viewer.
Call HistoryForm.Update


' Update the memory viewer.
Call MemoryForm.Update



Call HandleShutdown


If MSAgent.AgentActive Then If Speaker.Visible Then Speaker.Visible = False

' Enable the interface if not already visible
If Interface.Visible = False Then Interface.Show


If Interface.UsersWindow <> "" Then If Splash.Visible Then Unload Splash


End Sub


Private Sub Timer2_Timer()

Static LastMessage
' Update the main question handler.
'Static MQUMT
'If UsersMessageTime <> MQUMT Then SmartArse.InMessage = UsersMessage: NewMessageAvailable = True: MQUMT = UsersMessageTime


'If LastMessage <> BotsMessage Then If Timer < BotsMessageTime + (Len(BotsMessage) / 10) Then Exit Sub

Call SmartArse.Update

'If LastMessage <> BotsMessage Then If Timer < BotsMessageTime + (Len(BotsMessage) / 10) Then Exit Sub

'If SmartArse.OutMessage <> "" Then Call SetBotsMessage(SmartArse.OutMessage): LastMessage = BotsMessage




End Sub


