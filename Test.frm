VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[No character loaded] - Microsoft Character Animation Previewer"
   ClientHeight    =   5385
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6840
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5595
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Speak"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   5535
      TabIndex        =   17
      Top             =   3000
      Width           =   1170
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   5535
      TabIndex        =   15
      Top             =   255
      Width           =   1170
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   5535
      TabIndex        =   16
      Top             =   735
      Width           =   1170
   End
   Begin VB.Frame fraAnimation 
      Caption         =   "&Animations for"
      Enabled         =   0   'False
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   5355
      Begin VB.CheckBox optOutputStyle 
         Caption         =   "Stop &before next action"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   2925
         TabIndex        =   3
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox optOutputStyle 
         Caption         =   "Play sound &effects"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.ListBox lstAnimation 
         Enabled         =   0   'False
         Height          =   1620
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.Frame fraSpeechOutput 
      Caption         =   "Speech &Output"
      Enabled         =   0   'False
      Height          =   2085
      Left            =   60
      TabIndex        =   4
      Top             =   2445
      Width           =   5295
      Begin VB.TextBox txtSpeakText 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   930
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   255
         Width           =   5025
      End
      Begin VB.CheckBox optBalloonStyle 
         Caption         =   "Display &word balloon"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1290
         Width           =   1935
      End
      Begin VB.CheckBox optBalloonStyle 
         Caption         =   "Si&ze to text"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3375
         TabIndex        =   9
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CheckBox optBalloonStyle 
         Caption         =   "Auto &pace"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2085
         TabIndex        =   8
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CheckBox optBalloonStyle 
         Caption         =   "A&uto hide"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   900
         TabIndex        =   7
         Top             =   1650
         Width           =   1200
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Enabled         =   0   'False
      Height          =   720
      Left            =   60
      TabIndex        =   10
      Top             =   4575
      Width           =   4170
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Move"
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   2700
         TabIndex        =   18
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txtCharPosn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1845
         TabIndex        =   14
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox txtCharPosn 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   375
         TabIndex        =   12
         Top             =   255
         Width           =   570
      End
      Begin VB.Label lblCharPosn 
         Caption         =   "&Y:"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   13
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblCharPosn 
         Caption         =   "&X:"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   300
         Width           =   270
      End
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   6195
      Top             =   1620
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim Character As IAgentCtlCharacterEx

Dim NewBalloonStyleOption As Integer
Dim CharLoaded As Boolean
Dim IgnoreSizeEvent As Boolean
Dim CurrentIndex As Integer
Dim OpenSuccess As Boolean

Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Sub SetBalloonStyleOptions()

'-- This subroutine sets the check boxes for the
'-- the word balloon settings

'-- Check to see if the balloon is on
If Character.Balloon.Style And BalloonOn Then
    optBalloonStyle(0).Value = 1
Else
    optBalloonStyle(0).Value = 0
End If

'-- Check to see if Auto-Hide is on
If Character.Balloon.Style And AutoHide Then
    optBalloonStyle(1).Value = 1
Else
    optBalloonStyle(1).Value = 0
End If

'-- Check to see if Auto-Pace is on
If Character.Balloon.Style And AutoPace Then
    optBalloonStyle(2).Value = 1
Else
    optBalloonStyle(2).Value = 0
End If

'-- Check to see if Size-To-Text is on
If Character.Balloon.Style And SizeToText Then
    optBalloonStyle(3).Value = 1
Else
    optBalloonStyle(3).Value = 0
End If

'-- Set the controls based on Advanced Character Options
If Not Character.Balloon.Enabled Then
    optBalloonStyle(0).Enabled = False
    optBalloonStyle(1).Enabled = False
    optBalloonStyle(2).Enabled = False
    optBalloonStyle(2).Enabled = False
Else
    optBalloonStyle(0).Enabled = True
    optBalloonStyle(1).Enabled = True
    optBalloonStyle(2).Enabled = True
    optBalloonStyle(2).Enabled = True
End If

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

Private Sub lstAnimation_Click()

'-- Enable the Play button
cmdAction(0).Enabled = True

End Sub

Sub EnableControls()
'-- Enable the controls on the page

'-- Enable the Animation List Box
fraAnimation.Enabled = True
lstAnimation.Enabled = True

'-- Enable the Stop and Move buttons
cmdAction(1).Enabled = True
cmdAction(3).Enabled = True

'-- Enable the Play Sound Effects option only
'-- if enabled in the Advanced Character Options
If MyAgent.AudioOutput.Enabled And MyAgent.AudioOutput.SoundEffects Then
    optOutputStyle(0).Enabled = True
End If

'-- Enable the Stop Before Play option
optOutputStyle(1).Enabled = True

'-- Enable the Balloon Style options
optBalloonStyle(0).Enabled = True
optBalloonStyle(1).Enabled = True
optBalloonStyle(2).Enabled = True
optBalloonStyle(3).Enabled = True

'-- Enable the Speech Text Box
fraSpeechOutput.Enabled = True
txtSpeakText.Enabled = True
txtSpeakText.BackColor = vbWindowBackground

'-- Enable the X,Y position fields
fraPosition.Enabled = True
lblCharPosn(0).Enabled = True
txtCharPosn(0).Enabled = True
txtCharPosn(0).BackColor = vbWindowBackground
lblCharPosn(1).Enabled = True
txtCharPosn(1).Enabled = True
txtCharPosn(1).BackColor = vbWindowBackground

End Sub

Private Sub lstAnimation_DblClick()

cmdAction_Click (0)

End Sub

Private Sub lstAnimation_GotFocus()

'-- Make certain that the Move button isn't left
'-- as the default even if there is no selection
'-- in the list
cmdAction(3).Default = False

'-- Make certain that the Play button is the
'-- default button when the list box has focus.
cmdAction(0).Default = True

End Sub

Private Sub optBalloonStyle_GotFocus(Index As Integer)

'-- If a balloon style option gets the focus
'-- and the Speak button is enabled, make the
'-- the Speak button the default button
If cmdAction(2).Enabled Then
    cmdAction(2).Default = True
End If

End Sub

Private Sub txtCharPosn_Change(Index As Integer)

'-- If X or Y is empty then disable
'-- disable the Move button
If txtCharPosn(0).Text = "" Or txtCharPosn(1).Text = "" Then
    cmdAction(3).Enabled = False
    cmdAction(3).Default = False
Else
    cmdAction(3).Enabled = True
    cmdAction(3).Default = True
End If

'-- Check to determine if we get numeric input for
'-- the position, if not goto BadInput
On Error GoTo BadInput
PosnInput = CInt(txtCharPosn(Index).Text)

Exit Sub
BadInput:
Beep
txtCharPosn(Index).SelStart = 0
txtCharPosn(Index).SelLength = Len(txtCharPosn(Index).Text)

End Sub

Private Sub mnuExit_Click()

If CharLoaded = True Then
    Set Character = Nothing
    MyAgent.Characters.Unload "CharacterID"
    Unload Me
Else
    Unload Me
End If

End Sub

Private Sub mnuAbout_Click()

frmAbout.Show 1

End Sub

Private Sub MyAgent_AgentPropertyChange()

'-- Check to see if the user changed settings in
'-- Advanced Character Options

'-- Check to see if the user changed
'-- Play Character Sound Effects
'-- in Advanced Character Options
If Not MyAgent.AudioOutput.SoundEffects Then
    optOutputStyle(0).Enabled = False
Else
    optOutputStyle(0).Enabled = True
End If

'-- Check to see if the user changed
'-- Display Spoken Output In Word Balloon option
'-- in Advanced Character Options
If Not Character.Balloon.Enabled Then
    optBalloonStyle(0).Enabled = False
    optBalloonStyle(1).Enabled = False
    optBalloonStyle(2).Enabled = False
    optBalloonStyle(3).Enabled = False
Else
    optBalloonStyle(0).Enabled = True
    optBalloonStyle(1).Enabled = True
    optBalloonStyle(2).Enabled = True
    optBalloonStyle(3).Enabled = True
End If

End Sub

Private Sub MyAgent_Command(ByVal UserInput As Object)

'-- If the user selects the Advanced Character Options
'-- command in the character's pop-up menu
'-- make the window visible

If UserInput.Name = "AdvCharOptions" Then
    MyAgent.PropertySheet.Visible = True
End If

End Sub

Private Sub MyAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

'-- If the user drags the character
'-- update the character position fields

txtCharPosn(0) = Character.Left
txtCharPosn(1) = Character.Top

End Sub

Private Sub optOutputStyle_Click(Index As Integer)

'-- If the Play Sound Effects option is changed
'-- set the character's output option

If Index = 0 Then
    If optOutputStyle(0).Value Then
        Character.SoundEffectsOn = True
    Else
        Character.SoundEffectsOn = False
    End If
End If

End Sub

Private Sub optBalloonStyle_Click(Index As Integer)

'-- When the user changes a balloon style option
'-- update the character's word balloon settings

Select Case Index
    Case 0 '-- The balloon display option
        If optBalloonStyle(0).Value = 0 Then
            Character.Balloon.Style = Character.Balloon.Style And (Not BalloonOn)
            optBalloonStyle(1).Enabled = False
            optBalloonStyle(2).Enabled = False
            optBalloonStyle(3).Enabled = False
        Else
            Character.Balloon.Style = Character.Balloon.Style Or BalloonOn
            optBalloonStyle(1).Enabled = True
            optBalloonStyle(2).Enabled = True
            optBalloonStyle(3).Enabled = True
        End If
    
    Case 1 '-- The Auto-Hide option
        If optBalloonStyle(1).Value = 0 Then
            Character.Balloon.Style = Character.Balloon.Style And (Not AutoHide)
        Else
            Character.Balloon.Style = Character.Balloon.Style Or AutoHide
        End If
    
    Case 2 '-- The Auto-Pace option
        If optBalloonStyle(2).Value = 0 Then
            Character.Balloon.Style = Character.Balloon.Style And (Not AutoPace)
        Else
            Character.Balloon.Style = Character.Balloon.Style Or AutoPace
        End If
    
    Case 3 '-- The Size-To-Text option
        If optBalloonStyle(3).Value = 0 Then
            Character.Balloon.Style = Character.Balloon.Style And (Not SizeToText)
        Else
            Character.Balloon.Style = Character.Balloon.Style Or SizeToText
        End If
    
End Select

End Sub

Private Sub cmdAction_Click(Index As Integer)

'-- This routine processes the Play, Stop, Speak,
'-- and Move buttons

'-- If Stop Before Play is set, stop the character
'-- before the next request
If optOutputStyle(1).Value Then
    Character.Stop
End If

Select Case Index
    Case 0 '-- The Play button was clicked, play an animation

        '-- Play the animaton selected in the list box
        Character.Play lstAnimation.List(lstAnimation.ListIndex)

    Case 1 '-- The Stop button was chosen, stop the animation
    
        Character.Stop

    Case 2 '-- The Speak button was chosen
    
        '-- Speak the text if there is text
        If Not txtSpeakText.Text = "" Then
            Character.Speak txtSpeakText.Text
        End If
    
        txtSpeakText.SetFocus
        txtSpeakText.SelStart = 0
        txtSpeakText.SelLength = Len(txtSpeakText.Text)
    
    
    Case 3 '-- The Move button was chosen, move to the X,Y position
    
        Character.MoveTo CInt(txtCharPosn(0).Text), CInt(txtCharPosn(1).Text)
        
        txtCharPosn(CurrentIndex).SetFocus
        txtCharPosn(CurrentIndex).SelStart = 0
        txtCharPosn(CurrentIndex).SelLength = Len(txtCharPosn(CurrentIndex).Text)

End Select

End Sub

Private Sub mnuOpen_Click()

Dim AnimationName As Variant

'-- Set a flag to track success
OpenSuccess = False
        
CommonDialog1.CancelError = True

On Error GoTo ErrHandler
    
CommonDialog1.Flags = cdlOFNHideReadOnly

'-- Get the Windows directory name
Dim DirName As String
DirName = GetWindowsDir()

'-- Append the Agent Chars subdirectory
CommonDialog1.InitDir = DirName + "msagent\chars"

'-- Add the filter
CommonDialog1.Filter = "Microsoft Agent Characters (*.acs)|*.acs"
CommonDialog1.FilterIndex = 1

'-- Show the Open dialog
CommonDialog1.ShowOpen
    
'--Unload the previous character
On Error Resume Next
Set Character = Nothing
MyAgent.Characters.Unload "CharacterID"
    
'-- Load the new character
On Error GoTo ErrHandler
MyAgent.Characters.Load "CharacterID", CommonDialog1.filename

OpenSuccess = True

Set Character = MyAgent.Characters("CharacterID")

Me.Caption = Character.Name + " - Microsoft Character Animation Previewer"

'-- Set the character loaded flag
CharLoaded = True

'-- Set the character's language
Character.LanguageID = &H409

'-- Update the caption for the animation list box
fraAnimation.Caption = "&Animations for " + Character.Name

'-- Disable the Play button to avoid trying to play a null animation selection
cmdAction(0).Enabled = False

'-- Load the character's animation into the list box
lstAnimation.Clear
For Each AnimationName In Character.AnimationNames
    lstAnimation.AddItem AnimationName
Next
   
Character.Left = (frmTest.Left + 3450) / Screen.TwipsPerPixelX
Character.Top = (frmTest.Top + 900) / Screen.TwipsPerPixelY
Character.Show

txtCharPosn(0).Text = CStr(Character.Left)
txtCharPosn(1).Text = CStr(Character.Top)

'-- Update the state of the balloon style options
SetBalloonStyleOptions

'-- Initialize the pop-up menu commands
InitPopupMenuCmds

'-- Update the state of the controls to match the
'-- character's settings
EnableControls

lstAnimation.SetFocus

Exit Sub
ErrHandler:
If (Err.Number <> cdlCancel) Then
    If (OpenSuccess = False) Then
        MsgBox "There was an error opening the file " & CommonDialog1.filename
    End If
    Set Character = Nothing
End If

End Sub

Private Sub Form_Load()

'When the form loads, set the IgnoreSizeEvent flag
'(used to differentiate when the Character Animation
'Previewer window is restored), set the CharLoaded flag
'(used to track when a character is loaded),
'and set the initial state of the status bar.
CentreMe Me
IgnoreSizeEvent = True
CharLoaded = False

End Sub

Sub InitPopupMenuCmds()

'-- Add a command to the character to provide access
'-- to the Advanced Character Options
Character.Commands.RemoveAll
Character.Commands.Add "AdvCharOptions", "&Advanced Character Options"

End Sub

Private Sub Form_Resize()

'-- This routines hides or shows the character when the
'-- Character Animation Previewer window is mininized or
'-- restored

If IgnoreSizeEvent Then
    IgnoreSizeEvent = False
    Exit Sub
End If

If CharLoaded Then
    If Me.WindowState = vbMinimized Then
        Character.Hide True
    ElseIf Me.WindowState = vbNormal Then
        Character.Show True
    End If
End If

End Sub

Private Sub txtSpeakText_Change()

'-- The routine makes certain that there is text to
'-- speak before enabling the Speak button

If txtSpeakText.Text = "" Then
    cmdAction(2).Enabled = False
    cmdAction(2).Default = False
Else
    cmdAction(2).Enabled = True
    cmdAction(2).Default = True
End If

End Sub

Private Sub txtCharPosn_GotFocus(Index As Integer)

'-- This routine handles what happens when the
'-- X,Y position fields get the focus

'-- Make the Move button the default button
cmdAction(3).Default = True

'-- Select the text in the field
txtCharPosn(Index).SelStart = 0
txtCharPosn(Index).SelLength = Len(txtCharPosn(Index).Text)

'-- Set a flag to remember the text field we are in
CurrentIndex = Index

End Sub

Private Sub txtSpeakText_GotFocus()

'-- If the user clicks on this text box and
'-- it's enabled make The Speak button the
'-- default button
If cmdAction(2).Enabled Then
    cmdAction(2).Default = True
    txtSpeakText.SelStart = 0
    txtSpeakText.SelLength = Len(txtSpeakText.Text)
End If

End Sub

Sub CentreMe(P1 As Form)

If TypeOf P1 Is Form Then
    P1.Left = (Screen.Width - P1.Width) / 2
    P1.Top = (Screen.Height - P1.Height) / 2
End If

End Sub

