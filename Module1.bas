Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest



Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type Eye
     E_x As Long 'xval
     E_y As Long 'yval
     E_R As Integer 'radius
End Type

Public LEye As Eye
Public REye As Eye

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

'fungsi api untuk baca dan nulis file .ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Function GetX() As Long
    Dim x As POINTAPI
    GetCursorPos x
    GetX = x.x
End Function


Public Function GetY() As Long
    Dim Y As POINTAPI
    GetCursorPos Y
    GetY = Y.Y
End Function

Public Sub FormTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub

Public Sub FormNoTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub

Public Function WriteINI(SectionHeader As String, VariableName As String, Value As String, FileName As String) As String
    On Error Resume Next
    Dim x As Byte
    
    x = WritePrivateProfileString(SectionHeader, VariableName, Value, FileName)
End Function

Public Function ReadINI(SectionHeader As String, VariableName As String, FileName As String) As String
    On Error Resume Next
    Dim Buffer As String
    Dim x As Byte
    
    Buffer = String(255, 0)
    x = GetPrivateProfileString(SectionHeader, VariableName, "Default", Buffer, 255, FileName)
    If x <> 0 Then Buffer = Left$(Buffer, x)
    ReadINI = Buffer
End Function
