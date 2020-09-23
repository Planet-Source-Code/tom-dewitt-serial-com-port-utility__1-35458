Attribute VB_Name = "API"
Option Explicit
'   File:
'       API.Bas
'   Author:
'       Tom DeWitt
'   Description:
'       This Module Contains The API Functions To Create, Read, and Write To The Windows Registry As Well As The
'   Required Constants.
'-----------------------------------------------------------------------------------------------------------------------
'   Revisions:
'       Original 5/22/2002
'-----------------------------------------------------------------------------------------------------------------------
'   Functions And Subroutines:
'       1.  Sleep
'       2.  SleepEx
'       3.  RegOpenKeyEx                Alias: RegOpenKeyExA
'       4.  RegCloseKey
'       5.  RegQueryValueEx             Alias: RegQueryValueExA
'       6.  RegCreateKeyEx              Alias: RegCreateKeyExA
'       7.  RegDeleteKey                Alias: RegDeleteKeyA
'       8.  RegSetValueEx               Alias: RegSetValueExA
'       9.  RegEnumKey                  Alias: RegEnumKeyA
'       10. RegDeleteValue              Alias: RegDeleteValueA
'       11. RegEnumValue                Alias: RegEnumValueA
'       12. CopyMemory                  Alias: RtlMoveMemory
'       13. FormatMessage               Alias: FormatMessageA
'       14. keybd_event
'       15. GetVersionEx                Alias: GetVersionExA
'-----------------------------------------------------------------------------------------------------------------------
'   Properties:
'-----------------------------------------------------------------------------------------------------------------------
'   Required Functions,Subroutines,Properties,Variables,Etc.:
'
'-----------------------------------------------------------------------------------------------------------------------
'   Variables:
'       Public:
'
'-----------------------------------------------------------------------------------------------------------------------
'   Types:
'       Public
'-----------------------------------------------------------------------------------------------------------------------
Public Type POINTAPI
        x As Long
        y As Long
End Type
'   Constants:
'       Public:
'-----------------------------------------------------------------------------------------------------------------------
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_ALL = &H1F0000
'-----------------------------------------------------------------------------------------------------------------------
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'-----------------------------------------------------------------------------------------------------------------------
Public Const ERROR_SUCCESS = 0&
'-----------------------------------------------------------------------------------------------------------------------
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
'-----------------------------------------------------------------------------------------------------------------------
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
'-----------------------------------------------------------------------------------------------------------------------
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
'-----------------------------------------------------------------------------------------------------------------------
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
'-----------------------------------------------------------------------------------------------------------------------
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As _
    String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal _
    lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
    String) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, ByVal cbName As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal _
    lpValueName As String) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As _
    Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, _
lpcbData As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 _
As Long, ByVal Y2 As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint _
As POINTAPI) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, _
ByVal crColor As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
ByVal crColor As Long, ByVal wFillType As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
ByVal crColor As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth _
As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
As Long
'-----------------------------------------------------------------------------------------------------------------------

