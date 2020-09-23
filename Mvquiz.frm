VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmMVQuiz 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motor Vehicle Quiz with Animation"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   Icon            =   "Mvquiz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS TxSpeech 
      Height          =   465
      Left            =   4275
      OleObjectBlob   =   "Mvquiz.frx":030A
      TabIndex        =   18
      Top             =   75
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.OptionButton optAnswer4 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer3 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton optAnswer1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   555
      Left            =   7845
      TabIndex        =   0
      Top             =   5475
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image imgSpanner 
      Height          =   480
      Left            =   2400
      Picture         =   "Mvquiz.frx":0362
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3675
      Top             =   75
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label lblQuestions 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF0000&
      Caption         =   "Question:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF0000&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF0000&
      Caption         =   "4:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "3:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF0000&
      Caption         =   "2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblAnswer4 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnswer1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReStart 
         Caption         =   "Restart Quiz"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Character &Options"
      Begin VB.Menu mnuChange 
         Caption         =   "&Change Character"
         Begin VB.Menu mnuGenie 
            Caption         =   "&Genie the Genie"
         End
         Begin VB.Menu mnuMerlin 
            Caption         =   "&Merlin the Magician"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPeedy 
            Caption         =   "&Peedy the Parrot"
         End
         Begin VB.Menu mnuRobot 
            Caption         =   "&Robby the Robot"
         End
      End
      Begin VB.Menu mnuChangeVoice 
         Caption         =   "Change &Voice"
         Begin VB.Menu mnuVoice 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBalloons 
         Caption         =   "&Display Word Balloons"
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "Disable Sounds"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Character"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuRepeat 
      Caption         =   "Repeat &Question"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMVQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileNum As Integer
Dim Question(40) As String
Dim Answer(40, 4) As String
Dim Ans(4) As Integer
Dim Guess As Integer
Dim Score As Integer
Dim Questions As Integer
Dim Finalscore As String
Dim Level As String
Dim Response As Integer
Dim Search As Integer
Dim N As Integer
Dim FileLoc As String

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

With MyAgent
    If Button = 2 Then
        PopupMenu mnuOptions
    Else
        .StopAll
        .Play "Alert"
        If Sound = True Then .Speak "Don't touch me|Take your mouse off me|Leave me alone"
        .Play "RestPose"
    End If
End With

End Sub

Private Sub Agent1_DblClick(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

'MyAgent.StopAll
'If Sound = True Then MyAgent.Speak Answer(Questions, 1)

End Sub

Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

MyAgent.StopAll
MyAgent.Play "Blink"

End Sub

Private Sub Agent1_DragStart(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

MyAgent.StopAll
MyAgent.Play "Think"

End Sub

Private Sub Agent1_IdleStart(ByVal CharacterID As String)

Dim N As Integer

Randomize Timer
N = Int((7 * Rnd) + 1)
MyAgent.StopAll

With MyAgent
    Select Case N
        Case 1
            If Anim = "Robby" Then
                .Play "Idle3_1"
            Else
                .Play "LookDownBlink"
                .Play "LookDownBlink"
                .Play "LookDownBlink"
                .Play "LookDownReturn"
                .Stop
                .MoveTo 300, 700
                If Sound = True Then
                    .Speak "Man It's really dark ..inside your monitor!"
                    .MoveTo 300, 50
                    .MoveTo 400, 350
                    .Speak "Nice to be back!"
                    .Speak "Lets try again"
                End If
            End If

        Case 2
            If Anim = "Robby" Then
                .Play "LookDown"
                .Play "LookLeft"
                .Play "LookUp"
                .Play "LookRight"
                If Sound = True Then .Speak "Don't you know the answer!"
            Else
                .Play "LookDown"
                .Play "LookDownBlink"
                .Play "LookLeft"
                .Play "LookLeftBlink"
                .Play "LookUp"
                .Play "LookUpBlink"
                .Play "LookRight"
                .Play "LookRightBlink"
                If Search >= 1 Then
                    If Sound = True Then .Speak "Are you still having a problem with this question?"
                Else
                    If Sound = True Then .Speak "Are you having a problem with this question?"
                End If
                Search = Search + 1
            End If
        
        Case 3
            If Sound = True Then .Speak "I'll try and search for the correct answer"
            .Play "Search"
            If Sound = True Then
                .Speak "You're in luck, I think I've found it"
                .Speak Answer(Questions, 1)
            End If
        
        Case 4
            .Play "Idle1_1"
            .Play "Idle1_2"
            .Play "Idle1_3"
            .Play "Idle1_4"
            If Search >= 1 Then
                If Sound = True Then .Speak "It appears you are still struggling with this question!"
            Else
                If Sound = True Then .Speak "It appears you are struggling with this question!"
            End If
            Search = Search + 1
        
        Case 5
            If Search >= 1 Then
            If Sound = True Then .Speak "I'll try another search pattern for the correct answer"
            Else
            If Sound = True Then .Speak "I'll try to search for the correct answer"
            End If
            Search = Search + 1
            .Play "Search"
            If Sound = True Then .Speak "Sorry, no luck with the search"

        Case 6
            If Sound = True Then
                .Speak "I'll try a different search pattern for the correct answer"
                .Play "Search"
                .Speak "I think it might be. But I wouldn't place any bets on it"
                .Speak "because I'm not a mechanic"
            N = Int((4 * Rnd) + 1)
                .Speak Answer(Questions, N)
        End If
    
        Case 7
            .Play "DontRecognize"
            If Sound = True Then .Speak "I don't hear any input|I still don't hear any input|Have you fallen asleep"
        
        Case Else
            .Play "Idle3_1"
        
    End Select
End With

End Sub

Private Sub cmdContinue_Click()

With MyAgent
    .StopAll
    If Ans(Guess) = 1 Then Score = Score + 1
    If Ans(Guess) <> 1 Then
        If Sound = True Then
            .Speak "The correct answer is"
            .Speak Answer(Questions, 1)
            Sleep 5000
        Else
            MsgBox "Correct answer is:-" & vbCrLf & vbCrLf & Answer(Questions, 1), 64, "Correct Answer"
        End If
    End If
    lblScore.Caption = Score

    If Questions >= 40 Then Result: Exit Sub
    Runtime
End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKey1 Then optAnswer1.Value = True
If KeyCode = vbKeyNumpad1 Then optAnswer1.Value = True
If KeyCode = vbKey2 Then optAnswer2.Value = True
If KeyCode = vbKeyNumpad2 Then optAnswer2.Value = True
If KeyCode = vbKey3 Then optAnswer3.Value = True
If KeyCode = vbKeyNumpad3 Then optAnswer3.Value = True
If KeyCode = vbKey4 Then optAnswer4.Value = True
If KeyCode = vbKeyNumpad4 Then optAnswer4.Value = True

End Sub

Private Sub Form_Load()

On Error GoTo ErrHandler
Dim DirName As String

If App.PrevInstance = True Then
    MsgBox "This application is already running!", vbInformation + vbOKOnly, "Motor Vehicle Quiz is Running"
    End
End If
If Right(App.Path, 1) = "\" Then
    Ap = App.Path
Else
    Ap = App.Path & "\"
End If

CentreMe Me
Me.Show
RetVal = waveOutGetNumDevs()

If RetVal = 0 Then
    MsgBox "Your system cannot play Sound Files." & vbCrLf & vbCrLf & "So you won't hear any speech!", 48, "SoundCard Check"
End If
DirName = GetWindowsDir()
FileLoc = DirName & "Msagent\"
Success = IfFileExists(FileLoc & "Agentctl.dll")
If Success = False Then
    MsgBox "Msagent is not installed on this computer!" & vbCrLf & vbCrLf & "You can download them from:-" & vbCrLf & vbCrLf & "http://www.microsoft.com/msagent", 64, "Motor Vehicle Quiz"
    End
End If
FileLoc = DirName & "Msagent\Chars\"
Success = IfFileExists(FileLoc & "Merlin.acs")
If Success = False Then
    MsgBox "Merlin is not installed on this computer!" & vbCrLf & vbCrLf & "You my need MSagent as well." & vbCrLf & vbCrLf & "You can download them from:-" & vbCrLf & vbCrLf & "http://www.microsoft.com/msagent", 64, "Motor Vehicle Quiz"
    End
End If

ReDim SpeechID(TxSpeech.CountEngines)
mnuVoice(0).Caption = TxSpeech.ModeName(1)
SpeechID(0) = TxSpeech.ModeID(1)

For N = 1 To TxSpeech.CountEngines
    Load mnuVoice(N)
    mnuVoice(N).Visible = True
    mnuVoice(N).Caption = TxSpeech.ModeName(N)
    SpeechID(N) = TxSpeech.ModeID(N)
Next

StartSave
Anim = "Merlin"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H809   'Language ID = English
If RetVal = 0 Then
    MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
    mnuBalloons.Checked = vbChecked
Else
    MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    mnuBalloons.Checked = vbUnchecked
End If
MyAgent.MoveTo 400, 350
MyAgent.Show

optAnswer1.Visible = True
optAnswer2.Visible = True
optAnswer3.Visible = True
optAnswer4.Visible = True

optAnswer1.Value = False
cmdContinue.Visible = False
Sound = True
Init
Quiz
Runtime

Exit Sub
ErrHandler:
MsgBox Err.Description, 0, "Motor Vehicle Quiz"
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

With MyAgent
    .StopAll
    If Sound = True Then
        If Questions < 2 Then
            .Speak "You haven't started the quiz yet. Why are you leaving?"
        ElseIf Questions >= 2 And Questions < 40 Then
            .Speak "You haven't finished the quiz yet. You have only attempted " & Questions & " questions. Why are you leaving?"
        End If
        Sleep 7000
    End If
    .StopAll
    FinishSave
    Set MyAgent = Nothing
    Agent1.Characters.Unload Anim
    End
End With

End Sub

Private Sub mnuAbout_Click()

MyAgent.StopAll
frmAbout.Show 1

End Sub

Private Sub mnuChangeVoice_Click()

MyAgent.StopAll

End Sub

Private Sub mnuExit_Click()

With MyAgent
    .StopAll
    If Sound = True Then
        If Questions < 2 Then
            .Speak "You haven't started the quiz yet. Why are you leaving?"
        ElseIf Questions >= 2 And Questions < 40 Then
            .Speak "You haven't finished the quiz yet. You have only attempted " & Questions & " questions. Why are you leaving?"
        End If
        Sleep 7000
    End If
    .StopAll
    FinishSave
    Set MyAgent = Nothing
    Agent1.Characters.Unload Anim
    End
End With

End Sub

Public Sub Runtime()

Randomize Timer

Questions = Questions + 1

lblScore.Caption = Score
lblQuestions.Caption = Questions

Ans(1) = Int(Rnd * 4 + 1)
Do
    Ans(2) = Int(Rnd * 4 + 1)
Loop Until Ans(2) <> Ans(1)

Do
    Ans(3) = Int(Rnd * 4 + 1)
Loop Until Ans(3) <> Ans(1) And Ans(3) <> Ans(2)
Ans(4) = 10 - Ans(1) - Ans(2) - Ans(3)

lblQuestion.Caption = Question(Questions)

lblAnswer1.Caption = Answer(Questions, Ans(1))
lblAnswer2.Caption = Answer(Questions, Ans(2))
lblAnswer3.Caption = Answer(Questions, Ans(3))
lblAnswer4.Caption = Answer(Questions, Ans(4))

Guess = 0
optAnswer1.Value = False
optAnswer2.Value = False
optAnswer3.Value = False
optAnswer4.Value = False
cmdContinue.Visible = False
With MyAgent
    If Sound = True Then
        .Speak "Question " & Questions
        .Speak lblQuestion.Caption
        .Speak lblAnswer1.Caption
        .Speak "or"
        .Speak lblAnswer2.Caption
        .Speak "or"
        .Speak lblAnswer3.Caption
        .Speak "or"
        .Speak lblAnswer4.Caption
    End If
End With

End Sub

Private Sub mnuGenie_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Genie.acs")
If Success = False Then
    MyAgent.Play "Search"
    If Sound = True Then
        MyAgent.Speak "Sorry, Genie is not installed on this computer!"
        MyAgent.Speak "You can download him from:-"
        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Else
        MsgBox "Sorry, Genie is not installed on this computer!" & vbCrLf & "You can download him from http://www.microsoft.com/msagent/downloads.htm#character", 64, "Motor Vehicle Quiz"
    End If
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
Agent1.Characters.Unload Anim
Anim = "Genie"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H809 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = True
mnuPeedy.Checked = False
mnuRobot.Checked = False
mnuMerlin.Checked = False

End Sub

Private Sub mnuHide_Click()

With MyAgent
    mnuHide.Checked = Not mnuHide.Checked
    .StopAll
    Select Case mnuHide.Checked
        Case True
            Sound = False
            .Hide
        Case False
            Sound = True
            .Show
    End Select
End With

End Sub

Private Sub mnuInstructions_Click()

With MyAgent
    If Sound = True Then
        .StopAll
        .Speak "You will be shown 40 questions relating to Motor Vehicle studies."
        .Speak "Read each question carefully."
        .Speak "Keys:-"
        .Speak "You can use the Mouse to enter your choice."
        .Speak "You can also use the Numeric Keys to enter your choice, then press the Enter key."
        .Speak "Good Luck!!!"
    Else
        MsgBox "You will be shown 40 questions relating to Motor Vehicle studies.  Read each question carefully." & vbCrLf & vbCrLf & "Keys:-" & vbCrLf & "You can use the Mouse to enter your choice." & vbCrLf & "You can also use the Numeric Keys to enter your choice, then press the Enter key." & vbCrLf & vbCrLf & "Good Luck!!!", 64, "Instructions"
    End If
    optAnswer1.Value = False
    optAnswer2.Value = False
    optAnswer3.Value = False
    optAnswer4.Value = False
End With

End Sub

Private Sub mnuBalloons_Click()

mnuBalloons.Checked = Not mnuBalloons.Checked
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select

End Sub

Private Sub mnuMerlin_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Merlin.acs")
If Success = False Then
    MyAgent.Play "Search"
    If Sound = True Then
        MyAgent.Speak "Sorry, Merlin is not installed on this computer!"
        MyAgent.Speak "You can download him from"
        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Else
        MsgBox "Sorry, Merlin is not installed on this computer!" & vbCrLf & "You can download him from http://www.microsoft.com/msagent/downloads.htm#character", 64, "Motor Vehicle Quiz"
    End If
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
Agent1.Characters.Unload Anim
Anim = "Merlin"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H809 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = False
mnuRobot.Checked = False
mnuMerlin.Checked = True

End Sub

Private Sub mnuPeedy_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Peedy.acs")
If Success = False Then
    MyAgent.Play "Search"
    If Sound = True Then
        MyAgent.Speak "Sorry, Peedy is not installed on this computer!"
        MyAgent.Speak "You can download him from"
        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Else
        MsgBox "Sorry, Peedy is not installed on this computer!" & vbCrLf & "You can download him from http://www.microsoft.com/msagent/downloads.htm#character", 64, "Motor Vehicle Quiz"
    End If
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
Agent1.Characters.Unload Anim
Anim = "Peedy"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H809 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = True
mnuRobot.Checked = False
mnuMerlin.Checked = False

End Sub

Private Sub mnuRepeat_Click()

With MyAgent
    .StopAll
    If Sound = True Then
        .Speak "Question " & Questions
        .Speak lblQuestion.Caption
        .Speak lblAnswer1.Caption
        .Speak "or"
        .Speak lblAnswer2.Caption
        .Speak "or"
        .Speak lblAnswer3.Caption
        .Speak "or"
        .Speak lblAnswer4.Caption
    End If
End With

End Sub

Private Sub mnuReStart_Click()

MyAgent.StopAll

optAnswer1.Visible = True
optAnswer2.Visible = True
optAnswer3.Visible = True
optAnswer4.Visible = True
Init
Runtime

End Sub

Private Sub mnuRobot_Click()

MyAgent.StopAll
Success = IfFileExists(FileLoc & "Robby.acs")
If Success = False Then
    MyAgent.Play "Search"
    If Sound = True Then
        MyAgent.Speak "Sorry, Robby is not installed on this computer!"
        MyAgent.Speak "You can download him from"
        MyAgent.Speak "http://www.microsoft.com/msagent/downloads.htm#character"
    Else
        MsgBox "Sorry, Robby is not installed on this computer!" & vbCrLf & "You can download him from http://www.microsoft.com/msagent/downloads.htm#character", 64, "Motor Vehicle Quiz"
    End If
    Exit Sub
End If

MyAgent.Hide
Set MyAgent = Nothing
Agent1.Characters.Unload Anim
Anim = "Robby"
Agent1.Characters.Load Anim, Anim & ".acs"
Set MyAgent = Agent1.Characters(Anim)
MyAgent.AutoPopupMenu = False
MyAgent.LanguageID = &H809 'Language ID = English
MyAgent.MoveTo 400, 350
Select Case mnuBalloons.Checked
    Case False
        MyAgent.Balloon.Style = MyAgent.Balloon.Style And (Not BalloonOn)
    Case True
        MyAgent.Balloon.Style = MyAgent.Balloon.Style Or BalloonOn
End Select
MyAgent.Show
mnuGenie.Checked = False
mnuPeedy.Checked = False
mnuRobot.Checked = True
mnuMerlin.Checked = False

End Sub

Private Sub mnuSounds_Click()

mnuSounds.Checked = Not mnuSounds.Checked
MyAgent.StopAll
Select Case mnuSounds.Checked
    Case False
        Sound = True
        mnuRepeat.Enabled = True
    Case True
        Sound = False
        mnuRepeat.Enabled = False
End Select

End Sub

Private Sub mnuVoice_Click(Index As Integer)

On Error Resume Next

For N = 0 To TxSpeech.CountEngines
    mnuVoice(N).Checked = False
Next
mnuVoice(Index).Checked = True

With MyAgent
    .StopAll
    .TTSModeID = "{" & SpeechID(Index) & "}"
    .Speak "This is a test"
End With
Me.Caption = "Motor Vehicle Quiz with " & mnuVoice(Index).Caption

End Sub

Private Sub optAnswer1_Click()

Guess = 1
cmdContinue.Visible = True

End Sub

Private Sub optAnswer1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer1.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer1.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer2_Click()

Guess = 2
cmdContinue.Visible = True

End Sub

Private Sub optAnswer2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer2.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer2.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer3_Click()

Guess = 3
cmdContinue.Visible = True

End Sub

Private Sub optAnswer3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer3.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer3.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer4_Click()

Guess = 4
cmdContinue.Visible = True

End Sub

Public Sub Init()

Score = 0
Questions = 0
Search = 0

End Sub

Private Sub Result()

optAnswer1.Visible = False
optAnswer2.Visible = False
optAnswer3.Visible = False
optAnswer4.Visible = False
cmdContinue.Visible = False

If Score <= 9 Then Level = "Consider changing your profession."
If Score >= 10 And Score <= 19 Then Level = "You are a poor Mechanic."
If Score >= 20 And Score <= 27 Then Level = "You are a good Mechanic."
If Score >= 28 And Score <= 34 Then Level = "You are a good Technician."
If Score >= 35 And Score <= 40 Then Level = "You are a superb Technician."

Finalscore = "You scored " & Score & " out of " & Questions & "."

With MyAgent
    Select Case Score
        Case Is >= 31
            .Play "Announce"
        Case Is >= 20 And Score <= 30
            .Play "Congratulate"
        Case Is >= 10 And Score <= 19
            .Play "Congratulate_2"
        Case Is <= 9
            .Play "Confused"
    End Select
    If Sound = True Then
        .Speak Finalscore
        .Speak Level
        .Speak "Do you wish to Re-Run the Quiz?"
    Else
        MsgBox Finalscore & vbCrLf & vbCrLf & Level, 64, "Motor Vehicle Quiz Results"
    End If
    .Play "Writing"
    Response = MsgBox("Do you wish to Re-Run the Quiz.", 36, "Motor Vehicle Quiz")
    If Response = vbYes Then
        FinishSave
        .StopAll
        mnuReStart_Click
    Else
        FinishSave
        Set MyAgent = Nothing
        Agent1.Characters.Unload Anim
        End
    End If
End With

End Sub

Public Sub StartSave()

On Error GoTo ErrHandler

FileNum = FreeFile()
Open Ap & "Results.dat" For Append As #FileNum
Write #FileNum, "Start - " & Me.Caption, Format$(Date$, "dd/mm/yyyy"), Time$
Close #FileNum

Exit Sub
ErrHandler:
If Err = 61 Then
    MsgBox "This Disk is Full!" & vbCrLf & vbCrLf & "Put this program on a blank disk!", vbCritical + vbOKOnly, "Program Error"
    End
Else
    MsgBox Err.Description, 0, "Motor Vehicle Quiz"
End If

End Sub

Public Sub FinishSave()

On Error GoTo ErrHandler

If Questions < 40 Then Questions = Questions - 1
FileNum = FreeFile()
Open Ap & "Results.dat" For Append As #FileNum
Write #FileNum, "Finish - " & Me.Caption, Format$(Date$, "dd/mm/yyyy"), Time$, Score, Questions
Close #FileNum

Exit Sub
ErrHandler:
If Err = 61 Then
    MsgBox "This Disk is Full!" & vbCrLf & vbCrLf & "Put this program on a blank disk!", vbCritical + vbOKOnly, "Program Error"
    End
Else
    MsgBox Err.Description, 0, "Motor Vehicle Quiz"
End If

End Sub

Public Sub Quiz()

Question(1) = "When the engine is started from cold and is running on choke, the"
Answer(1, 1) = "crankcase lubricant is being contaminated by the fuel."
Answer(1, 2) = "mixture strength is too weak to give total burning of the fuel."
Answer(1, 3) = "engine must be allowed to reach its normal operating temperature before the choke is returned."
Answer(1, 4) = "ignition timing is automatically retarded to compensate for the increased burning rate."

Question(2) = "A crankcase is internally vented to"
Answer(2, 1) = "recirculate unburnt Hydrocarbons through the cylinders."
Answer(2, 2) = "pass oil mist into the cylinders to lubricate the pistons."
Answer(2, 3) = "circulate air to aid crankcase cooling."
Answer(2, 4) = "balance out cylinder pressures due to piston leakage."

Question(3) = "Raising the temperature of combustion tends to"
Answer(3, 1) = "raise the level of Oxides of Nitrogen in the exhaust."
Answer(3, 2) = "cause the engine to run on the rich side."
Answer(3, 3) = "increase the chance of incomplete burning of the fuel's Hydrogen content."
Answer(3, 4) = "produce more Carbon Dioxide in the exhaust gases."

Question(4) = "The quantity of Carbon Monoxide (CO) in the exhaust system is raised by"
Answer(4, 1) = "increasing the combustion temperature."
Answer(4, 2) = "advancing the ignition timing."
Answer(4, 3) = "poor exhaust manifold design."
Answer(4, 4) = "reducing the compression ratio."

Question(5) = "In an engine which lacks compression"
Answer(5, 1) = "there is a reduction in the combustion rate."
Answer(5, 2) = "there is a comparative reduction in the amount of fuel used."
Answer(5, 3) = "the ignition has to be retarded slightly from the manufacturer's specification."
Answer(5, 4) = "the percentage power loss is greatest at the higher engine speeds."

Question(6) = "A faulty radiator cap will"
Answer(6, 1) = "cause the engine to overheat."
Answer(6, 2) = "increase the boiling point of water."
Answer(6, 3) = "dangerously increase the cooling rate."
Answer(6, 4) = "increase fuel consumption."

Question(7) = "Increased fuel consumption can be caused by"
Answer(7, 1) = "a thermostat which opens at too low a temperature."
Answer(7, 2) = "the introduction of antifreeze into the cooling system."
Answer(7, 3) = "lowering the float level."
Answer(7, 4) = "advancing the ignition."

Question(8) = "The problem of dissociation means that"
Answer(8, 1) = "the mixture has to be enriched to give maximum engine power."
Answer(8, 2) = "the ignition timing has to be retarded at high engine speeds."
Answer(8, 3) = "the carburettor has to be fitted with an accelerator pump."
Answer(8, 4) = "the fuel has to doped with Tetra-ethyl-lead."

Question(9) = "A fixed choke carburettor has to be fitted with an idling by-pass system"
Answer(9, 1) = "because at low speeds an insufficient depression is felt in the choke."
Answer(9, 2) = "so that the mechanic can adjust the mixture strength over the range of engine speeds."
Answer(9, 3) = "because the main carburettor cannot maintain accurate metering at low engine speeds."
Answer(9, 4) = "to supplement the mixture strength at idling speed in order to overcome manifold precipitation."

Question(10) = "The voltmeter connected between the coil C.B. or negative terminal and earth with the ignition switched on will read battery voltage"
Answer(10, 1) = "when the contact breakers are open."
Answer(10, 2) = "only when the capacitor is removed."
Answer(10, 3) = "at all times."
Answer(10, 4) = "when the contact breakers are closed."

Question(11) = "The function of the capacitor in a coil ignition system is to"
Answer(11, 1) = "speed the collapse of the primary L.T. current."
Answer(11, 2) = "sustain the H.T. current."
Answer(11, 3) = "sustain the L.T. current."
Answer(11, 4) = "speed the collapse of the H.T. current."

Question(12) = "Retarded ignition timing would normally be corrected by"
Answer(12, 1) = "rotating the distributor body against the direction of rotor movement."
Answer(12, 2) = "rotating the distributor body in the direction of rotor movement."
Answer(12, 3) = "reducing the contact breaker gap."
Answer(12, 4) = "adjusting engine idle speed."

Question(13) = "Which of the following voltage and current figures would be reasonable in the secondary circuit of a coil ignition system"
Answer(13, 1) = "20,000 volts at 2.5 amps."
Answer(13, 2) = "20,000 volts at 12 amps."
Answer(13, 3) = "12 volts at 12 amps."
Answer(13, 4) = "2.5 millivolts at 20,000 amps."

Question(14) = "The spark is produced at the spark plug when the contact breakers are"
Answer(14, 1) = "just opening."
Answer(14, 2) = "fully closed."
Answer(14, 3) = "just closed."
Answer(14, 4) = "fully open."

Question(15) = "The effect of increasing the contact breaker gap is to"
Answer(15, 1) = "reduce the dwell angle."
Answer(15, 2) = "increase the dwell angle."
Answer(15, 3) = "increase the dwell time."
Answer(15, 4) = "cause arcing across the points."

Question(16) = "The coil produces a high induced voltage in the secondary windings using the principle of"
Answer(16, 1) = "mutual-induction."
Answer(16, 2) = "motion-induction."
Answer(16, 3) = "reverse-induction."
Answer(16, 4) = "self-induction."

Question(17) = "The maximum output voltage produced in the secondary windings of the coil will depend upon"
Answer(17, 1) = "the magnitude of the primary voltage."
Answer(17, 2) = "the spark plug electrode gap."
Answer(17, 3) = "the rotor gap."
Answer(17, 4) = "the resistance of the H.T. leads."

Question(18) = "The high voltage produced in the primary circuit as the contact breaker opens is due to"
Answer(18, 1) = "primary winding self-induction."
Answer(18, 2) = "battery voltage 'surge'."
Answer(18, 3) = "capacitor effect."
Answer(18, 4) = "a continuous E.M.F."

Question(19) = "One of the most apparent effects of an increase in engine speed on the oscilloscope showing the ignition 'secondary' pattern is that"
Answer(19, 1) = "the secondary voltage is reduced."
Answer(19, 2) = "the primary voltage is reduced."
Answer(19, 3) = "the primary voltage is increased."
Answer(19, 4) = "the secondary voltage is increased."

Question(20) = "The ignition timing is advanced mechanically as engine speed increases in order to compensate for"
Answer(20, 1) = "reduced time for combustion."
Answer(20, 2) = "an over-rich fuel/air mixture."
Answer(20, 3) = "an over-weak fuel/air mixture."
Answer(20, 4) = "reduced dwell angle."

Question(21) = "What would be the most likely effect of fitting a spark plug which is graded too hot for the engine"
Answer(21, 1) = "pre-ignition."
Answer(21, 2) = "increased firing voltage."
Answer(21, 3) = "spark plug electrode fouling."
Answer(21, 4) = "boiling of the water in the cooling system."

Question(22) = "The induced secondary voltage will decrease as engine speed is increased due to"
Answer(22, 1) = "reduced dwell time."
Answer(22, 2) = "reduced contact breaker efficiency."
Answer(22, 3) = "reduced dwell angle."
Answer(22, 4) = "coil heat saturation effect."

Question(23) = "Fouling of the spark plug insulation would result in the ceramic material becoming 'conductive'. This would be detected on an oscilloscope pattern by"
Answer(23, 1) = "a low 'firing' voltage."
Answer(23, 2) = "a high 'firing' voltage."
Answer(23, 3) = "reduced secondary oscillations."
Answer(23, 4) = "increased secondary oscillations."

Question(24) = "Increased electrical resistance of the H.T. leads would result in"
Answer(24, 1) = "increased 'firing' voltage."
Answer(24, 2) = "increased secondary current."
Answer(24, 3) = "decreased 'firing' voltage."
Answer(24, 4) = "increased secondary oscillations."

Question(25) = "An over-weak fuel/air mixture would show on an oscilloscope pattern by"
Answer(25, 1) = "increased 'firing' voltage."
Answer(25, 2) = "reduced 'firing' voltage."
Answer(25, 3) = "reduced spark period."
Answer(25, 4) = "increased spark period."

Question(26) = "'Resistor' type spark plugs are fitted in order to"
Answer(26, 1) = "suppress radio interference."
Answer(26, 2) = "reduce spark voltage."
Answer(26, 3) = "improve combustion."
Answer(26, 4) = "reduce exhaust emissions."

Question(27) = "The voltage required to cause a spark across the electrodes of the spark plug would be increased by"
Answer(27, 1) = "increased compression pressure."
Answer(27, 2) = "a reduced electrode gap."
Answer(27, 3) = "increased electrode temperature."
Answer(27, 4) = "advanced ignition timing."

Question(28) = "If the dwell angle of the contact breaker changes considerably as the engine speed increases, the likely cause would be"
Answer(28, 1) = "worn distributor shaft."
Answer(28, 2) = "advance of ignition timing."
Answer(28, 3) = "points bounce."
Answer(28, 4) = "reduced dwell angle."

Question(29) = "The terms 'ionization' and 'firing voltage' refer to"
Answer(29, 1) = "the voltage required to make the plug 'gap' conductive."
Answer(29, 2) = "the voltage of the spark across the electrodes of the spark plug."
Answer(29, 3) = "the voltage across the contact breakers when closed."
Answer(29, 4) = "the 'capacitance' of the capacitor."

Question(30) = "An engine is running rich after checking the float level and tuning the carburettor. The"
Answer(30, 1) = "fuel pump pressure should be checked."
Answer(30, 2) = "ignition timing should be re-checked."
Answer(30, 3) = "inlet manifold should be checked for air leaks."
Answer(30, 4) = "valve clearances should be increased."

Question(31) = "When testing a wax element thermostat, how long should you wait for the heat to fully penetrate the wax"
Answer(31, 1) = "2 to 3 minutes."
Answer(31, 2) = "5 to 10 seconds."
Answer(31, 3) = "10 minutes."
Answer(31, 4) = "20 seconds."

Question(32) = "What is the purpose of the by-pass in the cooling system"
Answer(32, 1) = "to allow circulation when the thermostat is closed."
Answer(32, 2) = "to allow cooling water to reach the underside of the valve seats."
Answer(32, 3) = "to prime the water pump."
Answer(32, 4) = "to by-pass the thermostat in the event of its failure."

Question(33) = "Some carburettors are fitted with pollution control devices.  One such device, for controlling emissions during cold start situations, is called the"
Answer(33, 1) = "fuel temperature compensation device."
Answer(33, 2) = "cold starting enrichment device."
Answer(33, 3) = "piston damper valve."
Answer(33, 4) = "by-pass emulsion system."

Question(34) = "Raising the compression ratio"
Answer(34, 1) = "increases to Carbon Monoxide and Oxides of Nitrogen in the exhaust."
Answer(34, 2) = "increases the Carbon Dioxide and unburnt Hydrocarbons."
Answer(34, 3) = "decreases the amount of unused Oxygen in the exhaust."
Answer(34, 4) = "controls the burning rate which reduces the Oxides of Nitrogen in the exhaust."

Question(35) = "Detonation (pinking) is most likely to occur when"
Answer(35, 1) = "accelerating hard from low speed."
Answer(35, 2) = "the ignition is retarded."
Answer(35, 3) = "throttling back after travelling at high speed."
Answer(35, 4) = "using a fuel with too high an octane rating."

Question(36) = "An air leak on the inlet manifold will"
Answer(36, 1) = "cause the engine to misfire."
Answer(36, 2) = "increase the compression pressure."
Answer(36, 3) = "cause the engine to overheat."
Answer(36, 4) = "increase fuel consumption."

Question(37) = "Running on 'tight' valve clearances"
Answer(37, 1) = "increases the valve opening period."
Answer(37, 2) = "reduces the valve open period."
Answer(37, 3) = "reduces the engine running temperature."
Answer(37, 4) = "makes no difference to valve timing."

Question(38) = "A slack timing chain causes"
Answer(38, 1) = "uneven running."
Answer(38, 2) = "overheating."
Answer(38, 3) = "the distributor to over-advance."
Answer(38, 4) = "a reduction in valve clearance."

Question(39) = "When tuning a carburettor, adjustment should be made"
Answer(39, 1) = "until maximum vacuum is achieved."
Answer(39, 2) = "until maximum tick-over speed is achieved."
Answer(39, 3) = "until 15 in Hg is achieved."
Answer(39, 4) = "by adjusting until maximum Hg is reached and then backing off 1 in Hg."

Question(40) = "What is an acceptable coil primary resistance of a Ballasted ignition coil"
Answer(40, 1) = "1.5 ohms."
Answer(40, 2) = "minimum - 15 ohms."
Answer(40, 3) = "3 ohms."
Answer(40, 4) = "0.4 - 0.8 ohms."

End Sub

Function GetWindowsDir() As String

Dim Temp As String
Dim Ret As Long
Const MAX_LENGTH = 145

Temp = String$(MAX_LENGTH, 0)
Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
Temp = Left$(Temp, Ret)
If Temp <> "" And Right$(Temp, 1) <> "\" Then
    GetWindowsDir = Temp & "\"
Else
    GetWindowsDir = Temp
End If

End Function

Sub CentreMe(P1 As Form)

If TypeOf P1 Is Form Then
    P1.Left = (Screen.Width - P1.Width) / 2
    P1.Top = (Screen.Height - P1.Height) / 2
End If

End Sub

Function IfFileExists(Fname As String) As Boolean

On Local Error Resume Next
 
Dim F As Integer

F = FreeFile()
Open Fname For Input As #F
If Err Then
    IfFileExists = False
Else
    IfFileExists = True
End If
Close #F

End Function

Private Sub optAnswer4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer4.MouseIcon = imgSpanner

End Sub

Private Sub optAnswer4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

optAnswer4.MouseIcon = imgSpanner

End Sub

