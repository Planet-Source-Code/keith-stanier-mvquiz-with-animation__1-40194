VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmOptions 
   Caption         =   "Character Options"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2115
      Left            =   2850
      TabIndex        =   4
      Top             =   375
      Width           =   1890
      Begin VB.CheckBox chkBalloons 
         Caption         =   "Don't Display Word Balloons"
         Height          =   465
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2625
      TabIndex        =   3
      Top             =   2625
      Width           =   2190
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   3450
      TabIndex        =   2
      Top             =   3675
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   465
      Left            =   3450
      TabIndex        =   1
      Top             =   3150
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   300
      Pattern         =   "*.acs"
      TabIndex        =   0
      Top             =   600
      Width           =   2190
   End
   Begin VB.Label lblInfo 
      Caption         =   "Load Character:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   750
      TabIndex        =   6
      Top             =   150
      Width           =   1440
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   300
      Top             =   3825
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdOK_Click()

Set MyAgent = Nothing
frmMVQuiz.Agent1.Characters.Unload "CharacterID"
frmMVQuiz.Agent1.Characters.Load txtFileName.Text, DATAPATH
Set MyAgent = frmMVQuiz.Agent1.Characters(txtFileName.Text)
MyAgent.LanguageID = &H409
If chkBalloons.Value = 1 Then
    frmMVQuiz.mnuBalloons.Caption = "&Don't Display Word Balloons"
Else
    frmMVQuiz.mnuBalloons.Caption = "&Display Word Balloons"
End If

End Sub

Private Sub File1_Click()

MyAgent.Hide
'Agent1.Characters.Unload MyAgent
frmMVQuiz.Agent1.Characters.Unload "merlin"
Agent1.Characters.Load File1.filename, File1.filename
frmMVQuiz.Agent1.Characters.Load File1.filename, File1.filename
Set MyAgent = Agent1.Characters.Character(File1.filename)
MyAgent.Show

txtFileName.Text = File1.filename

End Sub

Private Sub Form_Load()

File1.Path = "C:\windows\msagent\chars"

End Sub
