VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4275
   ClientLeft      =   4275
   ClientTop       =   4965
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrSysInfo 
      Interval        =   1
      Left            =   225
      Top             =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   5100
      TabIndex        =   24
      Top             =   825
      Width           =   1215
   End
   Begin VB.Frame fraInfo 
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   1800
      Width           =   5100
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "This program will give you a 40 question quiz relating to motor vehicle studies."
         Height          =   840
         Index           =   1
         Left            =   255
         TabIndex        =   1
         Top             =   480
         Width           =   4605
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Operating System"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   1
      Left            =   100
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   5100
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   8
         Top             =   850
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2265
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4725
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "General Info"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   4
      Left            =   100
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   5100
      Begin VB.Label lblInfo 
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   4830
      End
      Begin VB.Label lblInfo 
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   4230
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Processor Type and Memory Statistics"
      ClipControls    =   0   'False
      Height          =   1680
      Index           =   2
      Left            =   100
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   5100
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "pagefile"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblR 
         Caption         =   "PageFile"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   840
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   840
         Width           =   1695
      End
      Begin VB.Shape shpFrame 
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Shape shpBar 
         BackStyle       =   1  'Opaque
         DrawMode        =   7  'Invert
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "virtual"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblResInfo 
         Alignment       =   2  'Center
         Caption         =   "physical"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblR 
         Caption         =   "Physical"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblR 
         Caption         =   "Virtual"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   6
         Top             =   480
         Width           =   4900
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Video"
      ClipControls    =   0   'False
      Height          =   1695
      Index           =   3
      Left            =   100
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   5100
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   600
      MouseIcon       =   "About.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "Click to E-mail"
      Top             =   3900
      Width           =   4065
   End
   Begin VB.Label lblEmailMe 
      Alignment       =   2  'Center
      Caption         =   "Any problems or comments E-mail me at:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   34
      Top             =   3600
      Width           =   4065
   End
   Begin VB.Label lblSerialNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   33
      Top             =   1515
      Width           =   5640
   End
   Begin VB.Label lblLicense 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   32
      Top             =   1245
      Width           =   5640
   End
   Begin VB.Label lblClick2 
      Caption         =   "System Info"
      Height          =   240
      Left            =   5325
      TabIndex        =   31
      Top             =   2925
      Width           =   1080
   End
   Begin VB.Label lblClick 
      Caption         =   "Click for"
      Height          =   225
      Left            =   5475
      TabIndex        =   30
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   75
      TabIndex        =   29
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   6300
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   28
      Top             =   75
      UseMnemonic     =   0   'False
      Width           =   6300
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   105
      TabIndex        =   27
      Top             =   450
      Width           =   5595
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   26
      Top             =   765
      Width           =   5640
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   25
      Top             =   990
      Width           =   5625
   End
   Begin VB.Image imgIconMain 
      Height          =   480
      Left            =   450
      Picture         =   "About.frx":030A
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgSystemInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   5550
      Picture         =   "About.frx":0614
      ToolTipText     =   "Click for System Information"
      Top             =   2250
      Width           =   540
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Num As Integer
Dim Status As Long
Dim TotalBytes As Currency
Dim FreeBytes As Currency
Dim BytesAvailableToCaller As Currency
Dim DiskSize As String
Dim DiskSpace As String
Dim Email As String

Dim buffer As String * 100
Dim MainKeyHandle As Long
Dim hKey As Long
Dim rtn As Long
Dim lBuffer As Long
Dim sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim TotalPhysicalMemory As Long
Dim AvailablePhysicalMemory As Long
Dim TotalPageFile As Long
Dim AvailablePageFile As Long
Dim TotalVirtualMemory As Long
Dim AvailableVirtualMemory As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SystemInfo)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4
Const KEY_ALL_ACCESS = &H3F
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const SYNCHRONIZE = &H100000
Const ERROR_SUCCESS = 0&

Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Const PROCESSOR_INTEL_386 = 386
Const PROCESSOR_INTEL_486 = 486
Const PROCESSOR_INTEL_PENTIUM = 586
Const PROCESSOR_MIPS_R4000 = 4000
Const PROCESSOR_ALPHA_21064 = 21064
Const KL_NAMELENGTH = 9

Private Type MYVERSION
    lMajorVersion As Long
    lMinorVersion As Long
    lExtraInfo As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128  ' Maintenance string for PSS usage
End Type

Private Type SystemInfo
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language

Private Sub FillSysInfo()

Dim FreeSpace As Currency
Dim FreeBlock As Currency
Dim YourMem As MEMORYSTATUS
Dim mOS As osinfo
Dim layoutname As String * KL_NAMELENGTH
'Operating System Info. Frame 1
Dim YourSystem As SystemInfo
GetSystemInfo YourSystem
Set mOS = New osinfo

lblInfo(2).Caption = mOS.OSName & " Version: " & Trim(Str$(mOS.OSMajorVersion)) & "." & Trim(Str$(mOS.OSMinorVersion)) & "." & Trim(Str$(mOS.OSBuildNumber)) & "  " & Trim(mOS.PSSInfo)
lblInfo(3).Caption = "User Name: " & GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
lblInfo(4).Caption = "Computer Name: " & GetStringValue("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")

GetSystemInfo YourSystem
Select Case YourSystem.dwProcessorType
    Case PROCESSOR_INTEL_386
            lblInfo(5).Caption = "CPU: " & "Intel 386"
    Case PROCESSOR_INTEL_486
            lblInfo(5).Caption = "CPU: " & "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
            lblInfo(5).Caption = "CPU: " & "Intel Pentium"
    Case PROCESSOR_MIPS_R4000
            lblInfo(5).Caption = "CPU: " & "MIPS R4000"
    Case PROCESSOR_ALPHA_21064
            lblInfo(5).Caption = "CPU: " & "Alpha 21064"
End Select

lblInfo(8).Caption = "Video Driver: " & GetSysIni("boot.description", "display.drv")
lblInfo(9).Caption = "Resolution: " & Screen.Width \ Screen.TwipsPerPixelX & " x " & Screen.Height \ Screen.TwipsPerPixelY

Dim Col As String
Col = DeviceColors((hDC))

Select Case Col
    Case "16"
        lblInfo(10).Caption = "Colours: " & 16
    Case "256"
        lblInfo(10).Caption = "Colours: " & 256
    Case "65536"
        lblInfo(10).Caption = "Colours: " & "True Color (16 bit)"
    Case "1.677722E+07"
        lblInfo(10).Caption = "Colours: " & "True Color (24 bit)"
    Case "4.294967E+09"
        lblInfo(10).Caption = "Colours: " & "True Color (32 bit)"
End Select

'General info. Frame 4
Status = GetDiskFreeSpaceEx("C:", BytesAvailableToCaller, TotalBytes, FreeBytes)
If Status = 0 Then
    DiskSize = "Unknown"
    DiskSpace = "Unknown"
Else
    DiskSize = Format(TotalBytes * 10000, "#,##0") & " Bytes"
    DiskSpace = Format(FreeBytes * 10000, "#,##0") & " Bytes"
End If

rtn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE, buffer, 99)

lblInfo(11).Caption = "Free Disk Space: " & DiskSpace
lblInfo(12).Caption = "Hard Disk Size: " & DiskSize
lblInfo(13).Caption = "Printer: " & GetStringValue("HKEY_CURRENT_CONFIG\System\CurrentControlSet\Control\Print\Printers", "Default")
lblInfo(14).Caption = "Language: " & LPSTRToVBString(buffer)

End Sub

Private Sub cmdOK_Click()

MyAgent.StopAll
Unload Me
frmMVQuiz.optAnswer1.Value = False
frmMVQuiz.optAnswer2.Value = False
frmMVQuiz.optAnswer3.Value = False
frmMVQuiz.optAnswer4.Value = False
frmMVQuiz.cmdContinue.Visible = False

End Sub

Private Sub Form_Load()

MyAgent.StopAll

Num = 1

Email = "kstan1sc@stokecoll.ac.uk"
lblEmail.Caption = Email
Call Setup
Call FillSysInfo

End Sub

Private Sub CPU()

Dim YourMemory As MEMORYSTATUS
Dim intX As Integer
Dim lWidth As Integer

If fraInfo(2).Visible Then
    For intX = 1 To 3
        lblR(intX).Visible = True
        lblResInfo(intX).Visible = True
        shpBar(intX).Visible = True
        shpFrame(intX).Visible = True
    Next
Else
    For intX = 1 To 3
        lblR(intX).Visible = False
        lblResInfo(intX).Visible = False
        shpBar(intX).Visible = False
        shpFrame(intX).Visible = False
    Next
End If
    
YourMemory.dwLength = Len(YourMemory)
GlobalMemoryStatus YourMemory

With YourMemory
lblInfo(6).Caption = "Memory Available: " & Format$(TotalPhysicalMemory, "###,###,###") & " KB. Free Memory: " & Format$(AvailablePhysicalMemory, "###,###,###") & " KB"

'Check width before setting to try and cut down on screen "flashing"
lWidth = shpFrame(1).Width * (.dwAvailPhys / .dwTotalPhys)
If lWidth <> shpBar(1).Width Then
    shpBar(1).Width = lWidth
End If

lWidth = shpFrame(2).Width * (.dwAvailVirtual / .dwTotalVirtual)
If lWidth <> shpBar(2).Width Then
    shpBar(2).Width = lWidth
End If

lWidth = shpFrame(3).Width * (.dwAvailPageFile / .dwTotalPageFile)
If lWidth <> shpBar(3).Width Then
    shpBar(3).Width = lWidth
End If
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblEmail.FontBold = False
lblEmail.ForeColor = vbBlue

End Sub

Private Sub imgSystemInfo_Click()

MyAgent.StopAll

Select Case Num
    Case 0 'Program Info
        fraInfo(0).Visible = True
        fraInfo(1).Visible = False
        fraInfo(2).Visible = False
        fraInfo(3).Visible = False
        fraInfo(4).Visible = False
    Case 1 'OS
        fraInfo(1).Visible = True
        fraInfo(0).Visible = False
        fraInfo(2).Visible = False
        fraInfo(3).Visible = False
        fraInfo(4).Visible = False
        If Sound = True Then
            MyAgent.Speak fraInfo(1).Caption
            MyAgent.Speak lblInfo(2).Caption
            MyAgent.Speak lblInfo(3).Caption
            MyAgent.Speak lblInfo(4).Caption
        End If
        
    Case 2 'Processor
        fraInfo(2).Visible = True
        fraInfo(0).Visible = False
        fraInfo(1).Visible = False
        fraInfo(3).Visible = False
        fraInfo(4).Visible = False
        If Sound = True Then
            MyAgent.Speak fraInfo(2).Caption
            MyAgent.Speak lblInfo(5).Caption
            MyAgent.Speak lblInfo(6).Caption
        End If
        
    Case 3 'Video
        fraInfo(3).Visible = True
        fraInfo(0).Visible = False
        fraInfo(1).Visible = False
        fraInfo(2).Visible = False
        fraInfo(4).Visible = False
        If Sound = True Then
            MyAgent.Speak fraInfo(3).Caption
            MyAgent.Speak lblInfo(8).Caption
            MyAgent.Speak "Resolution: " & Screen.Width \ Screen.TwipsPerPixelX & " by " & Screen.Height \ Screen.TwipsPerPixelY
            MyAgent.Speak lblInfo(10).Caption
        End If
        
    Case 4 'General
        fraInfo(4).Visible = True
        fraInfo(0).Visible = False
        fraInfo(1).Visible = False
        fraInfo(2).Visible = False
        fraInfo(3).Visible = False
        If Sound = True Then
            MyAgent.Speak fraInfo(4).Caption
            MyAgent.Speak lblInfo(12).Caption
            MyAgent.Speak lblInfo(11).Caption
            MyAgent.Speak lblInfo(13).Caption
            MyAgent.Speak lblInfo(14).Caption
        End If
        
End Select

Num = Num + 1
If Num = 5 Then Num = 0

End Sub

Private Sub Setup()

App.Title = "Motor Vehicle Quiz with Animation"
MakeFrm3D Me
lblAppName(0).Caption = App.Title
lblAppName(1).Caption = App.Title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblCopyright(0).Caption = "Copyright © 2002 by Keith Stanier"
lblLicense.Caption = "Registered to: " & GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
lblSerialNum.Caption = "Serial Number: " & GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductID")

fraInfo(0).Caption = App.Title
If Sound = True Then
    MyAgent.Speak App.Title
    MyAgent.Speak lblVersion.Caption
    MyAgent.Speak "Copyright \Map=""2 thousand and 2 ""=""2002""\" & " by \Map=""k'eeth""=""Keith""\" & "\Map=""stannier""=""Stanier""\"
    MyAgent.Speak lblInfo(1).Caption
End If

End Sub

Private Sub MakeFrm3D(TargetForm As Form)

Dim iFormTop As Integer
Dim iFormLeft As Integer
Dim iFormRight As Integer
Dim iFormBottom As Integer
Dim OriginalDrawWidth As Integer
Dim OriginalAutoRedraw As Integer
Dim iOriginalScaleMode As Integer
Dim WHITE
Dim DARK_GRAY
WHITE = &HFFFFFF
DARK_GRAY = &H808080

OriginalDrawWidth = TargetForm.DrawWidth    'Save Original DrawWidth
OriginalAutoRedraw = TargetForm.AutoRedraw  'Save Original AutoRedraw
iOriginalScaleMode = TargetForm.ScaleMode   'Save Original ScaleMode
TargetForm.DrawWidth = 1                    'Lines will be drawn 1 unit thick
TargetForm.AutoRedraw = True                'Let Windows automatically repaint lines
TargetForm.ScaleMode = 1                    'Set ScaleMode to Twips
iFormTop = 0
iFormLeft = 0
iFormRight = TargetForm.ScaleWidth - Screen.TwipsPerPixelY
iFormBottom = TargetForm.ScaleHeight - Screen.TwipsPerPixelX
TargetForm.Line (iFormLeft, iFormTop)-(iFormRight, iFormTop), WHITE
TargetForm.Line (iFormLeft, iFormTop)-(iFormLeft, iFormBottom), WHITE
TargetForm.Line (iFormLeft, iFormBottom)-(iFormRight + Screen.TwipsPerPixelY, iFormBottom), DARK_GRAY
TargetForm.Line (iFormRight, iFormTop)-(iFormRight, iFormBottom + Screen.TwipsPerPixelX), DARK_GRAY
TargetForm.DrawWidth = OriginalDrawWidth    'Reset Original DrawWidth
TargetForm.AutoRedraw = OriginalAutoRedraw  'Reset Original AutoRedraw
TargetForm.ScaleMode = iOriginalScaleMode   'Reset Original ScaleMode

End Sub

Private Sub lblEmail_Click()

On Error GoTo ErrHandler

ShellExecute hwnd, "open", "mailto:" & Email, vbNullString, vbNullString, SW_SHOW

Exit Sub
ErrHandler:
MsgBox "This computer doesn't have E-mail facilities!", 48, "E-mail fault"
Resume Next

End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblEmail.ForeColor = vbBlue
lblEmail.FontBold = True

End Sub

Private Sub tmrSysInfo_Timer()

Dim YourMemory As MEMORYSTATUS
Dim intX As Integer
Dim lWidth As Integer

GetMemoryStats

If fraInfo(2).Visible Then
    For intX = 1 To 3
        lblR(intX).Visible = True
        lblResInfo(intX).Visible = True
        shpBar(intX).Visible = True
        shpFrame(intX).Visible = True
    Next
Else
    For intX = 1 To 3
        lblR(intX).Visible = False
        lblResInfo(intX).Visible = False
        shpBar(intX).Visible = False
        shpFrame(intX).Visible = False
    Next
End If
    
YourMemory.dwLength = Len(YourMemory)
GlobalMemoryStatus YourMemory

With YourMemory
    lblInfo(6).Caption = "Memory Available: " & Format$(TotalPhysicalMemory, "###,###,###") & " KB. Free Memory: " & Format$(AvailablePhysicalMemory, "###,###,###") & " KB"
    
    'Check width before setting to try and cut down on screen "flashing"
    lWidth = shpFrame(1).Width * (.dwAvailPhys / .dwTotalPhys)
    If lWidth <> shpBar(1).Width Then
        shpBar(1).Width = lWidth
    End If

    lWidth = shpFrame(2).Width * (.dwAvailVirtual / .dwTotalVirtual)
    If lWidth <> shpBar(2).Width Then
        shpBar(2).Width = lWidth
    End If

    lWidth = shpFrame(3).Width * (.dwAvailPageFile / .dwTotalPageFile)
    If lWidth <> shpBar(3).Width Then
        shpBar(3).Width = lWidth
    End If
End With

End Sub

Function DeviceColors(hDC As Long) As Single

Const PLANES = 14
Const BITSPIXEL = 12
DeviceColors = 2 ^ (GetDeviceCaps(hDC, PLANES) * GetDeviceCaps(hDC, BITSPIXEL))

End Function

Function GetSysIni(section, key)

Dim RetVal As String
Dim Worked As Long

RetVal = String(255, 0)
Worked = GetPrivateProfileString(section, key, "", RetVal, Len(RetVal), "System.ini")
If Worked = 0 Then
    GetSysIni = "Unknown"
Else
    GetSysIni = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If

End Function

Function LPSTRToVBString(ByVal s$)

Dim Nullpos As Long

Nullpos = InStr(s$, Chr(0))
If Nullpos > 0 Then
    LPSTRToVBString = Left(s$, Nullpos - 1)
Else
    LPSTRToVBString = "Unknown"
End If

End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long

Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

On Error GoTo QueryValueExError

' Determine the size and type of data to be read
lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> 0 Then Error 5

Select Case lType
    ' For strings
    Case REG_SZ:
        sValue = String(cch, 0)

        lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = 0 Then
            vValue = Left(sValue, cch - 1)
        Else
            vValue = "Unknown"
        End If
        ' For DWORDS
    Case REG_DWORD:
        lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = 0 Then vValue = lValue
    Case Else
        'all other data types not supported
        lrc = -1
End Select

QueryValueExExit:
QueryValueEx = lrc
Exit Function

QueryValueExError:
Resume QueryValueExExit

End Function

Function GetStringValue(subkey As String, Entry As String)

Call ParseKey(subkey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, subkey, 0, KEY_READ, hKey) 'open the key
If rtn = ERROR_SUCCESS Then 'if the key could be opened then
    sBuffer = Space(255)     'make a buffer
    lBufferSize = Len(sBuffer)
    rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
    If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
        rtn = RegCloseKey(hKey)  'close the key
        sBuffer = Trim(sBuffer)
        GetStringValue = Left(sBuffer, Len(sBuffer) - 1) 'return the value to the user
    Else                        'otherwise, if the value couldnt be retreived
    'if the user wants errors displayed then
    GetStringValue = "Unknown" 'tell the user what was wrong
    End If
    End If
Else 'otherwise, if the key couldnt be opened
    GetStringValue = "Unknown"       'return Error to the user
    'if the user wants errors displayed then
End If

End Function

Sub ParseKey(Keyname As String, Keyhandle As Long)

rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname

Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1)) 'seperate the Keyname
Keyname = Right(Keyname, Len(Keyname) - rtn)

End Sub

Function GetMainKeyHandle(MainKeyName As String) As Long

Select Case MainKeyName
    Case "HKEY_CLASSES_ROOT"
        GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Private Function GetMemoryStats()

Dim ms As MEMORYSTATUS

GlobalMemoryStatus ms

TotalPhysicalMemory = ms.dwTotalPhys \ 1024
AvailablePhysicalMemory = ms.dwAvailPhys \ 1024
TotalPageFile = ms.dwTotalPageFile \ 1024
AvailablePageFile = ms.dwAvailPageFile \ 1024
TotalVirtualMemory = ms.dwTotalVirtual \ 1024
AvailableVirtualMemory = ms.dwAvailVirtual \ 1024

End Function

