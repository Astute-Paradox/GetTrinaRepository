VERSION 5.00
Begin VB.Form SpeechRecognizer 
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "SpeechRecognizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OutMessage
Public InMessage
Private LastSpokenMessage
Public UseSpeechRecognition

Public Active

Public WithEvents Vdict1 As Vdict
Attribute Vdict1.VB_VarHelpID = -1
Option Compare Text

Dim WithEvents RecoContext As SpSharedRecoContext
Attribute RecoContext.VB_VarHelpID = -1
Dim Grammar As ISpeechRecoGrammar



Private LastUsersMessageTime
Private LastBotsMessageTime


Public Sub DisableEngine()
UseSpeechRecognition = False
Active = False

On Error GoTo Done
If IsObject(Vdict1) Then GoTo DisableSap4

If IsObject(Grammar) Then GoTo DisableSap5



DisableSap5:
Grammar.DictationSetState SGDSInactive
Exit Sub


DisableSap4:
Vdict1.Deactivate


Done:

End Sub


Public Sub EnableEngine()

    If Active Then Exit Sub
    On Error GoTo Done

'Try get sapi-5 engine
    If (RecoContext Is Nothing) Then
        Debug.Print "Initializing SAPI 5 reco context object..."
        Set RecoContext = New SpSharedRecoContext
        Set Grammar = RecoContext.CreateGrammar(1)
        Grammar.DictationLoad
    End If
    
    Grammar.DictationSetState SGDSActive
    UseSpeechRecognition = True
    Active = True
    Exit Sub

    
'Try get sapi-4 engine
    Set Vdict1 = New Vdict
    If IsObject(Vdict1) Then GoTo InitSap4
    Exit Sub
InitSap4:
    Debug.Print "Using sapi 4 recog engine..."
    Vdict1.Initialized = 1
    Vdict1.Mode = 32
    Vdict1.ActivateAndAssignWindow 0
    UseSpeechRecognition = True
    Active = True

Done:
End Sub

Public Sub Update()

InMessage = ""
OutMessage = ""
If UsersMessageTime <> LastUsersMessageTime Then InMessage = UsersMessage: LastUsersMessageTime = UsersMessageTime


If LastSpokenMessage <> "" Then Call SetUsersMessage(LastSpokenMessage)
LastSpokenMessage = ""
If OutMessage <> "" Then Call TidyMessage(OutMessage)


If InMessage Like "*Use speech recog*" Then GoTo EnableIt
If InMessage Like "*Enable speech recog*" Then GoTo EnableIt
If InMessage Like "*Turn on speech recog*" Then GoTo EnableIt
If InMessage Like "*Disable speech recog*" Then GoTo DisableIt
If InMessage Like "*Deactivate speech recog*" Then GoTo DisableIt
If InMessage Like "*Do not use speech recog*" Then GoTo DisableIt
If InMessage Like "*Switch off speech recog*" Then GoTo DisableIt
GoTo Done



EnableIt:
Call EnableEngine
OutMessage = "Speech engine activated."
GoTo Done



DisableIt:
Call DisableEngine
OutMessage = "Speech engine deactivated."


Done:
If OutMessage <> "" Then Call SetBotsMessage(OutMessage)



End Sub

Private Sub Form_Load()
'    On Error Resume Next
'    Set Vdict1 = New Vdict

End Sub

Private Sub RecoContext_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)



    LastSpokenMessage = Result.PhraseInfo.GetText


    ' Append the new text to the text box, and add a space at the end of the
    ' text so that it looks better
    
'    txtSpeech.SelStart = m_cChars
'    txtSpeech.SelText = strText & " "
'    m_cChars = m_cChars + 1 + Len(strText)






End Sub

Private Sub Vdict1_PhraseFinish(ByVal flags As Long, ByVal phrase As String)
LastSpokenMessage = phrase
End Sub


