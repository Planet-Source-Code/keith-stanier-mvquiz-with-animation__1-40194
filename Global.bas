Attribute VB_Name = "Module1"
Option Explicit

Public Success As Boolean
Public Sound As Boolean
Public RetVal As Long
Public Anim As String
Public Pname As String
Public SpeechID() As String
Public Ap As String
Public MyAgent As IAgentCtlCharacterEx

Public Const BalloonOn As Integer = 1

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

