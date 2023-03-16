VERSION 5.00
Begin VB.Form MSAgent 
   Caption         =   "MS Agent"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "MSAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Action
Public InMessage
Public OutMessage
Public WithEvents Agent1 As Agent
Attribute Agent1.VB_VarHelpID = -1
Public AgentActive As Boolean
Public AgentName
Public Animation
Public MessageToSay
Option Compare Text



Public MyRequest

Private LastUsersMessageTime
Private LastBotsMessageTime


Private Sub DisableMSAgent()
If AgentActive = False Then Exit Sub

AgentActive = False
Agent1.Characters("Body").Hide
Agent1.Characters.Unload "Body"
Agent1.Connected = False



End Sub

Private Sub Msg_DisableAgent()
If InMessage Like "*Hide *" Then GoTo Check2
If InMessage Like "*Do not use *" Then GoTo Check2
If InMessage Like "*Turn off *" Then GoTo Check2
If InMessage Like "*Disable *" Then GoTo Check2
If InMessage Like "*Remove *" Then GoTo Check2
If InMessage Like "*Deactivate *" Then GoTo Check2
Exit Sub

Check2:
If InMessage Like "*agent*" Then GoTo DisableIt
If InMessage Like "*character*" Then GoTo DisableIt
Exit Sub

DisableIt:
Call DisableMSAgent

'OutMessage = "The character has been disabled."


End Sub


Private Sub Msg_MoveAgent()
If InMessage Like "*Move *" Then If AgentActive Then Word = "Move": GoTo CheckMore
If InMessage Like "*Go *" Then If AgentActive Then Word = "Go": GoTo CheckMore
If InMessage Like "*Fly *" Then If AgentActive Then Word = "Fly": GoTo CheckMore
If InMessage Like "*Float*" Then If AgentActive Then Word = "Float": GoTo CheckMore
Exit Sub

CheckMore:
If InMessage Like Word + " *" Then GoTo MoveIt
If InMessage Like "* you " + Word Then GoTo MoveIt
If InMessage Like "Please " + Word Then GoTo MoveIt
Exit Sub



MoveIt:
If InMessage Like "* other*side*" Then GoTo ChangeSide
If InMessage Like "* around*" Then GoTo ChangeSide
If InMessage Like "* about*" Then GoTo ChangeSide
S = 0
NewXPosition = Agent1.Characters("Body").Left
NewYPosition = Agent1.Characters("Body").Top
If InMessage Like "* middle*" Then GoSub GetCentrePosition: S = 1
If InMessage Like "* center*" Then GoSub GetCentrePosition: S = 1
If InMessage Like "* center*" Then GoSub GetCentrePosition: S = 1
If InMessage Like "* corner*" Then NewYPosition = (Screen.Height / Screen.TwipsPerPixelY) - Agent1.Characters("Body").Height: NewXPosition = (Screen.Width / Screen.TwipsPerPixelX) - Agent1.Characters("Body").Width: S = 1
If InMessage Like "* side*" Then NewXPosition = (Screen.Width / Screen.TwipsPerPixelX) - Agent1.Characters("Body").Width: S = 1
If InMessage Like "* up[!a-z]*" Then NewYPosition = 0: S = 1
If InMessage Like "* top[!a-z]*" Then NewYPosition = 0: S = 1
If InMessage Like "* down*" Then NewYPosition = (Screen.Height / Screen.TwipsPerPixelY) - Agent1.Characters("Body").Height: S = 1
If InMessage Like "* bottom[!a-z]*" Then NewYPosition = (Screen.Height / Screen.TwipsPerPixelY) - Agent1.Characters("Body").Height: S = 1
If InMessage Like "* left[!a-z]*" Then NewXPosition = 0: S = 1
If InMessage Like "* right[!a-z]*" Then NewXPosition = (Screen.Width / Screen.TwipsPerPixelX) - Agent1.Characters("Body").Width: S = 1
If S = 1 Then GoTo DoIt
Exit Sub

GetCentrePosition:
NewXPosition = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - (Agent1.Characters("Body").Width / 2)
NewYPosition = (Screen.Height / Screen.TwipsPerPixelY) / 2 - (Agent1.Characters("Body").Height / 2)
Return


ChangeSide:
CurrentXPosition = Agent1.Characters("Body").Left
NewYPosition = Agent1.Characters("Body").Top
NewXPosition = (Screen.Width / Screen.TwipsPerPixelX) - Agent1.Characters("Body").Width
If CurrentXPosition > (Screen.Width / Screen.TwipsPerPixelX) / 2 Then NewXPosition = 0


DoIt:
Agent1.Characters("Body").MoveTo NewXPosition, NewYPosition, 5

Replies = Array("Ok.", "Here I am.", "Done.", "I am here.", "Right.", "Your wish is my command.")
ReplyNumber = Int((Rnd * 5) + 1)
OutMessage = Replies(ReplyNumber)


End Sub


' This gives the MS Agent the ability to handle user messages like:
' Wave and say Hello.
' Look up and say up there.
' Look down.
' Say 'It is very cold outside.'
'
' If you just specify a statement for it to say but no action then
' it will try and automatically choose an appropriate action to perform
' while saying the specified statement.
' For example - if you tell it 'Say Hello' and don't specify an action
' for it to perform then the Wave action will be performed.
'
'

Private Sub GetActionAndMessageToSay()
Action = ""
If InMessage Like "If *" Then Exit Sub


GetAppropriateAction:
If MessageToSay Like "Hello*pleased to meet you*" Then Action = "Greet.": Exit Sub
If MessageToSay Like "hello[!a-z]*" Then Action = "Wave.": Exit Sub
If MessageToSay Like "Are you *there*[?]*" Then Action = "Get attention.": Exit Sub
If MessageToSay Like "Is *anybody *there*" Then Action = "Get attention.": Exit Sub
If MessageToSay Like "Yes[!a-z]*" Then Action = "Acknowledge.": Exit Sub
If MessageToSay Like "Yeah[!a-z]*" Then Action = "Acknowledge.": Exit Sub
If MessageToSay Like "Ok[!a-z]*" Then Action = "Acknowledge.": Exit Sub
If MessageToSay Like "OH NO[!a-z]*" Then Action = "Look alert.": Exit Sub
If MessageToSay Like "I *to announce *" Then Action = "Announce.": Exit Sub
If MessageToSay Like "*I don't know*" Then Action = "Look confused.": Exit Sub
If MessageToSay Like "Well done*" Then Action = "Congratulate.": Exit Sub
If MessageToSay Like "Congratulations*" Then Action = "Congratulate.": Exit Sub
If MessageToSay Like "I *congratulate*" Then Action = "Congratulate.": Exit Sub
If MessageToSay Like "I have no idea*" Then Action = "Decline.": Exit Sub
If MessageToSay Like "I don't have a clue*" Then Action = "Decline.": Exit Sub
If MessageToSay Like "I can do magic*" Then Action = "Do magic.": Exit Sub
If MessageToSay Like "I beg your pardon*" Then Action = "Ask user to repeat.": Exit Sub
If MessageToSay Like "Pardon." Then Action = "Ask user to repeat.": Exit Sub
If MessageToSay Like "What[?]*" Then Action = "Ask user to repeat.": Exit Sub
If MessageToSay Like "What did you say[?]*" Then Action = "Ask user to repeat.": Exit Sub
If MessageToSay Like "What was that[?]*" Then Action = "Ask user to repeat.": Exit Sub
If MessageToSay Like "Let me explain*" Then Action = "Explain.": Exit Sub
If MessageToSay Like "Up[!a-z]*" Then Action = "Look up.": Exit Sub
If MessageToSay Like "Down[!a-z]*" Then Action = "Look down.": Exit Sub
If MessageToSay Like "I am so happy*" Then Action = "Look pleased.": Exit Sub
If MessageToSay Like "I am pleased*" Then Action = "Look pleased.": Exit Sub
If MessageToSay Like "Hurray*" Then Action = "Look pleased.": Exit Sub
If MessageToSay Like "I am *busy*" Then Action = "Look busy.": Exit Sub
If MessageToSay Like "I am reading*" Then Action = "Read.": Exit Sub
If MessageToSay Like "Reading*" Then Action = "Read.": Exit Sub
If MessageToSay Like "I am *[!a-z]sad*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "I don't feel well*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "I feel bad[!a-z]*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "I feel terrible*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "I am not happy*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "That is awful*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "That* not nice*" Then Action = "Look sad.": Exit Sub
If MessageToSay Like "There is no need *that*" Then Action = "Look sad.": Exit Sub

End Sub




Private Sub PortrayBotsStatus()
If AgentActive = False Then Exit Sub
Static Reading
If BotReading Then GoTo StartIt Else GoTo StopIt
Exit Sub



StartIt:
If Reading = True Then Exit Sub
Agent1.Characters("Body").Play "Reading"
Reading = True
Exit Sub


StopIt:
If Reading = False Then Exit Sub
Agent1.Characters("Body").Stop
Agent1.Characters("Body").Play "ReadReturn"
Reading = False
Exit Sub


End Sub

Public Sub ShowAgentOptions()
If AgentActive Then Agent1.PropertySheet.Visible = True

End Sub


' This will open up MSAgents default properties window so the user can select the default agent.
Public Sub ShowAvailableAgents()
On Error GoTo Oops
If IsObject(Agent1) Then Agent1.ShowDefaultCharacterProperties

Oops:

End Sub


Public Sub Msg_PlayAnimation()
If AgentActive = False Then Exit Sub
If InMessage Like "Play *" Then Animation = WordsBetween("Play", ".", InMessage): GoTo PlayIt
Exit Sub

PlayIt:
On Error GoTo Done
Agent1.Characters("Body").Play Animation
OutMessage = "Done."

Done:
End Sub

Public Sub EnableMSAgent(Name)

If AgentActive Then Exit Sub

On Error GoTo HandleError
Set Agent1 = New Agent


If IsObject(Agent1) Then GoTo EnableIt
Exit Sub


EnableIt:
Agent1.Connected = True
If Name = "" Then Agent1.Characters.Load "Body" Else Agent1.Characters.Load "Body", Name + ".acs"


Agent1.Characters("Body").Top = (Screen.Height / Screen.TwipsPerPixelY) - Agent1.Characters("Body").Height
Agent1.Characters("Body").Left = (Screen.Width / 2) / Screen.TwipsPerPixelX - (Agent1.Characters("Body").Width / 2)



'The speak section will display the agent when necessary
'Agent1.Characters("Body").Show
AgentName = Agent1.Characters("Body").Name

AgentActive = True
Exit Sub


HandleError:
Debug.Print "Oops!"

End Sub


Public Sub Update()

Call PortrayBotsStatus

OutMessage = ""
InMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then LastUsersMessageTime = UsersMessageTime: InMessage = UsersMessage

If BotsMessageTime <> LastBotsMessageTime Then MessageToSay = BotsMessage: LastBotsMessageTime = BotsMessageTime


If InMessage = "" Then If MessageToSay = "" Then GoTo Done
Call Msg_AvailableAgents: If OutMessage <> "" Then GoTo Done
Call Msg_UseAgent: If OutMessage <> "" Then GoTo Done
Call Msg_AnimateThenSpeak: If OutMessage <> "" Then GoTo Done
If AgentActive = False Then GoTo Done
Call Msg_ShowOptions: If OutMessage <> "" Then GoTo Done
Call Msg_PlayAnimation: If OutMessage <> "" Then GoTo Done
Call Msg_MoveAgent: If OutMessage <> "" Then GoTo Done
Call Msg_DisableAgent: If OutMessage <> "" Then GoTo Done
Exit Sub


Done:
If OutMessage <> "" Then MessageToSay = OutMessage: Call SetBotsMessage(OutMessage): LastBotsMessageTime = BotsMessageTime



End Sub


Private Sub Msg_UseAgent()
If InMessage Like "*Show *" Then GoTo Check2
If InMessage Like "*Use *" Then GoTo Check2
If InMessage Like "*Enable *" Then GoTo Check2
If InMessage Like "*Display *" Then GoTo Check2
If InMessage Like "*View *" Then GoTo Check2
If InMessage Like "*Activate*" Then GoTo Check2
Exit Sub

Check2:
If InMessage Like "*option*" Then Exit Sub
If InMessage Like "*agent*" Then GoTo UseIt
If InMessage Like "*character*" Then GoTo UseIt
Exit Sub

UseIt:
If AgentActive Then OutMessage = "I am already active. To disable me, just tell me 'Disable agent'.": Exit Sub
Call EnableMSAgent("")
If AgentActive = False Then OutMessage = "There has been a problem trying to enable the agent.": Exit Sub

OutMessage = "Hello, I am pleased to meet you."

End Sub

Private Sub Msg_ShowOptions()
If InMessage Like "*Show*" Then GoTo Check2
If InMessage Like "*Display*" Then GoTo Check2
If InMessage Like "*View*" Then GoTo Check2
If InMessage Like "*see *" Then GoTo Check2
Exit Sub

Check2:
If InMessage Like "*Option*" Then GoTo Check3
If InMessage Like "*Properties*" Then GoTo Check3
If InMessage Like "*setting*" Then GoTo Check3
Exit Sub


Check3:
If InMessage Like "*agent*" Then GoTo ShowEm
If InMessage Like "*" + MSAgent.AgentName + "*" Then GoTo ShowEm
Exit Sub


ShowEm:
Call MSAgent.ShowAgentOptions
OutMessage = "Ok."
End Sub


Private Sub Msg_AvailableAgents()
If InMessage Like "*Select*" Then GoTo Check3
If InMessage Like "*Show*" Then GoTo Check2
If InMessage Like "*Display*" Then GoTo Check2
If InMessage Like "*View*" Then GoTo Check2
If InMessage Like "*see*" Then GoTo Check2
Exit Sub

Check2:
If InMessage Like "*available*" Then GoTo Check3
If InMessage Like "*selectable*" Then GoTo Check3
If InMessage Like "*default*" Then GoTo Check3
If InMessage Like "* all *" Then GoTo Check3
Exit Sub

Check3:
If InMessage Like "* agent*" Then GoTo ShowEm
If InMessage Like "* character*" Then GoTo ShowEm
Exit Sub

ShowEm:
If AgentActive = False Then Call EnableMSAgent("")
If AgentActive = False Then OutMessage = "There was a problem activating MSAgent.": Exit Sub

Call MSAgent.ShowAvailableAgents
OutMessage = "Ok."


End Sub


' This will animate MS Agent according to the message it is saying.
Private Sub Msg_AnimateThenSpeak()
If AgentActive = False Then Exit Sub
Call GetActionAndMessageToSay
On Error GoTo Done
If Action <> "" Then GoTo ProcessAction
If InMessage <> "" Then Action = InMessage: GoTo ProcessAction
GoSub Speak
Exit Sub


ProcessAction:
If Action Like "Wave[!a-z]*" Then GoSub Wave: GoTo ActionDone
If Action Like "Greet[!a-z]*" Then GoSub Greet: GoTo ActionDone
If Action Like "Get *attention*" Then GoSub GetAttention: GoTo ActionDone
If Action Like "Acknowledge*" Then GoSub Acknowledge: GoTo ActionDone
If Action Like "Look alert*" Then GoSub Alert: GoTo ActionDone
If Action Like "Announce*" Then GoSub Announce: GoTo ActionDone
If Action Like "Blink.*" Then GoSub Blink: GoTo ActionDone
If Action Like "*look confused*" Then GoSub Confused: GoTo ActionDone
If Action Like "Congratulate*" Then GoSub Congratulate: GoTo ActionDone
If Action Like "Decline*" Then GoSub Decline: GoTo ActionDone
If Action Like "Do *magic*" Then GoSub DoMagic: GoTo ActionDone
If Action Like "Show *magic*" Then GoSub DoMagic: GoTo ActionDone
If Action Like "Do *trick*" Then GoSub DoMagic: GoTo ActionDone
If Action Like "Show *trick*" Then GoSub DoMagic: GoTo ActionDone
If Action Like "Ask *to repeat*" Then GoSub Pardon: GoTo ActionDone
If Action Like "Explain." Then GoSub Explain: GoTo ActionDone
If Action Like "Point down*" Then Direction = "Down": GoSub Point: GoTo ActionDone
If Action Like "Point up*" Then Direction = "Up": GoSub Point: GoTo ActionDone
If Action Like "Point left*" Then Direction = "Left": GoSub Point: GoTo ActionDone
If Action Like "Point right*" Then Direction = "Right": GoSub Point: GoTo ActionDone
If Action Like "Gesture down*" Then Direction = "Down": GoSub Point: GoTo ActionDone
If Action Like "Gesture up*" Then Direction = "Up": GoSub Point: GoTo ActionDone
If Action Like "Gesture left*" Then Direction = "Left": GoSub Point: GoTo ActionDone
If Action Like "Gesture right*" Then Direction = "Right": GoSub Point: GoTo ActionDone
If Action Like "Look down*" Then Direction = "Down": GoSub Look: GoTo ActionDone
If Action Like "Look up*" Then Direction = "Up": GoSub Look: GoTo ActionDone
If Action Like "Look right*" Then Direction = "Right": GoSub Look: GoTo ActionDone
If Action Like "Look left*" Then Direction = "Left": GoSub Look: GoTo ActionDone
If Action Like "Look pleased*" Then GoSub Pleased: GoTo ActionDone
If Action Like "Look busy*" Then GoSub LookBusy: GoTo ActionDone
If Action Like "Read[!a-z]*" Then GoSub Read: GoTo ActionDone
If Action Like "Look sad*" Then GoSub LookSad: GoTo ActionDone
If Action Like "*be sad[!a-z]*" Then GoSub LookSad: GoTo ActionDone
If Action Like "Search*" Then GoSub Search: GoTo ActionDone
If Action Like "Listen*" Then GoSub Listen: GoTo ActionDone
If Action Like "Suggest*" Then GoSub Suggest: GoTo ActionDone
If Action Like "Look surprised*" Then GoSub LookSurprised: GoTo ActionDone
If Action Like "Think*" Then GoSub Think: GoTo ActionDone
If Action Like "Look uncertain*" Then GoSub LookUncertain: GoTo ActionDone
If Action Like "Write[!a-z]*" Then GoSub Writing: GoTo ActionDone
Exit Sub


Writing:
Agent1.Characters("Body").Play "Write"
GoSub Speak
Agent1.Characters("Body").Play "WriteContinued"
Agent1.Characters("Body").Play "WriteReturn"
Return



LookUncertain:
Animation = "Uncertain"
GoSub PlayAnim
Return


Think:
Animation = "Think"
GoSub PlayAnim
Return


LookSurprised:
Animation = "Surprised"
GoSub PlayAnim
Return


Suggest:
Animation = "Suggest"
GoSub PlayAnim
Return



Listen:
Agent1.Characters("Body").Play "StartListening"
Return


Search:
Animation = "Search"
GoSub PlayAnim
Return


LookSad:
Animation = "Sad"
GoSub PlayAnim
Return


Read:
Agent1.Characters("Body").Play "Read"
GoSub Speak
Agent1.Characters("Body").Play "ReadContinued"
Agent1.Characters("Body").Play "ReadContinued"
Agent1.Characters("Body").Play "ReadReturn"
Return


LookBusy:
GoSub Speak
Agent1.Characters("Body").Play "Process"
Return


Pleased:
Animation = "Pleased"
GoSub PlayAnim
Return



Look:
GoSub Speak
Agent1.Characters("Body").Play "Look" + Direction
Return


Point:
Animation = "Gesture" + Direction: GoSub PlayAnim
Return



Explain:
'If MessageToSay = "" Then OutMessage = "It's easy."
Agent1.Characters("Body").Play "Explain"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return


Pardon:
'If MessageToSay = "" Then OutMessage = "What did you say?."
Agent1.Characters("Body").Play "DontRecognize"
GoSub Speak
Return


DoMagic:
'If MessageToSay = "" Then OutMessage = "Not very magical,I know."
Agent1.Characters("Body").Play "DoMagic1"
Agent1.Characters("Body").Play "DoMagic2"
Agent1.Characters("Body").Play "Blink"
GoSub Speak
Return



Decline:
'If MessageToSay = "" Then OutMessage = "No thanks."
Agent1.Characters("Body").Play "Decline"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return



Congratulate:
'If MessageToSay = "" Then OutMessage = "You're great."
Agent1.Characters("Body").Play "Congratulate"
GoSub Speak
Agent1.Characters("Body").Play "Pleased"
Return


Confused:
'If MessageToSay = "" Then OutMessage = "I don't understand."
Agent1.Characters("Body").Play "Confused"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return



Blink:
'If MessageToSay = "" Then OutMessage = "I hope I didn't miss anything."
Agent1.Characters("Body").Play "Blink"
GoSub Speak
Return



Announce:
'If MessageToSay = "" Then OutMessage = "Done."
Agent1.Characters("Body").Play "Announce"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return


Alert:
'If MessageToSay = "" Then OutMessage = "I am always alert."
Agent1.Characters("Body").Play "Alert"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return


Acknowledge:
'If MessageToSay = "" Then OutMessage = "Ok."
Agent1.Characters("Body").Play "Acknowledge"
GoSub Speak
Return


GetAttention:
'If MessageToSay = "" Then OutMessage = "Wakey wakey."
Agent1.Characters("Body").Play "GetAttention"
Agent1.Characters("Body").Play "GetAttentionContinued"
Agent1.Characters("Body").Play "GetAttentionContinued"
GoSub Speak
Agent1.Characters("Body").Play "GetAttentionReturn"
Return

Greet:
'If MessageToSay = "" Then OutMessage = "How's that?"
Agent1.Characters("Body").Play "Greet"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return

Wave:
'If MessageToSay = "" Then OutMessage = "That was easy."
Agent1.Characters("Body").Play "Wave"
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return


PlayAnim:
Agent1.Characters("Body").Play Animation
GoSub Speak
Agent1.Characters("Body").Play "Blink"
Return


Speak:
Call ReplaceWords("Answerpad", "Answer pad", MessageToSay)
If Agent1.Characters("Body").Visible = False Then Agent1.Characters("Body").Show
If MessageToSay <> "" Then Agent1.Characters("Body").Speak MessageToSay: MessageToSay = ""
Return


ActionDone:
'If MessageToSay = "" Then OutMessage = "Ok."

Done:
End Sub


Private Sub Form_Load()


AgentActive = False
' Create an agent object
'On Error Resume Next
'Set Agent1 = New Agent



End Sub


