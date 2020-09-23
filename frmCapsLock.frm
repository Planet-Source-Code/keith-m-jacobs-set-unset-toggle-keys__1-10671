VERSION 5.00
Begin VB.Form frmCapsLock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Toggle Lock Keys"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkToggle 
      Caption         =   "SCRL"
      Height          =   375
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   555
   End
   Begin VB.CheckBox chkToggle 
      Caption         =   "CAPS"
      Height          =   375
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   300
      Width           =   555
   End
   Begin VB.CheckBox chkToggle 
      Caption         =   "NUM"
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   555
   End
End
Attribute VB_Name = "frmCapsLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare Type for API call:
      Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128   '  Maintenance string for PSS usage
      End Type

      ' API declarations:

    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
    Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

    ' Constant declarations:
    Private Const KEYEVENTF_EXTENDEDKEY = &H1
    Private Const KEYEVENTF_KEYUP = &H2
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
 
    Private bNoClick As Boolean
Private Sub chkToggle_Click(Index As Integer)
Dim o As OSVERSIONINFO
Dim NumLockState As Boolean
Dim ScrollLockState As Boolean
Dim CapsLockState As Boolean
Dim keys(0 To 255) As Byte

    If bNoClick Then Exit Sub
    
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    GetKeyboardState keys(0)

    Select Case Index
        Case 0  ' NumLock
            ' NumLock handling:
            NumLockState = keys(vbKeyNumlock)
            If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98
    
                keys(vbKeyNumlock) = Abs(Not NumLockState)
                SetKeyboardState keys(0)
            ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   '=== WinNT
                'Simulate Key Press
                keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
                'Simulate Key Release
                keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
            End If

        Case 1  ' CapsLock
            ' CapsLock handling:
            CapsLockState = keys(vbKeyCapital)
            If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98
                keys(vbKeyCapital) = Abs(Not CapsLockState)
                SetKeyboardState keys(0)
            ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   '=== WinNT
                'Simulate Key Press
                keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
                'Simulate Key Release
                keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
            End If

        Case 2  ' ScrollLock
            ' ScrollLock handling:
            ScrollLockState = keys(vbKeyScrollLock)
            If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98
                keys(vbKeyScrollLock) = Abs(Not ScrollLockState)
                SetKeyboardState keys(0)
            ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   '=== WinNT
                'Simulate Key Press
                keybd_event vbKeyScrollLock, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
                'Simulate Key Release
                keybd_event vbKeyScrollLock, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
            End If
    End Select
    
End Sub

Private Sub Form_Load()
Dim lngNumLockState As Long
Dim lngCapsLockState As Long
Dim lngScrollLockState As Long
    
    ' Get starting toggle states
    lngNumLockState = GetKeyState(vbKeyNumlock)
    lngCapsLockState = GetKeyState(vbKeyCapital)
    lngScrollLockState = GetKeyState(vbKeyScrollLock)
    
    bNoClick = True
    If lngNumLockState And 1 Then chkToggle(0).Value = 1
    bNoClick = False
    
    bNoClick = True
    If lngCapsLockState And 1 Then chkToggle(1).Value = 1
    bNoClick = False
    
    bNoClick = True
    If lngScrollLockState And 1 Then chkToggle(2).Value = 1
    bNoClick = False
    
End Sub


