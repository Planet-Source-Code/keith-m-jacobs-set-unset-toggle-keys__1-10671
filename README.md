<div align="center">

## Set/Unset Toggle keys


</div>

### Description

Toggle Num Lock, Caps Lock and Scroll Lock under Windows 95/98/NT/2000
 
### More Info
 
Refer to MSKB Q177674


<span>             |<span>
---                |---
**Submitted On**   |2000-08-14 10:00:18
**By**             |[Keith M Jacobs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/keith-m-jacobs.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD89438142000\.zip](https://github.com/Planet-Source-Code/keith-m-jacobs-set-unset-toggle-keys__1-10671/archive/master.zip)

### API Declarations

```
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128  ' Maintenance string for PSS usage
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
```





