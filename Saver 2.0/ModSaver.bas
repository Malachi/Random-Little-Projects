Attribute VB_Name = "ModSaver"
Option Explicit 'Declare variables

'Thanks to Patrick K. Bigley at Planet-Source-Code for providing
'systems tray related code below
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type

    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDOWN = &H201 'Button down
    Public Const WM_LBUTTONUP = &H202 'Button up
    Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
    Public Const WM_RBUTTONDOWN = &H204 'Button down
    Public Const WM_RBUTTONUP = &H205 'Button up
    Public Const WM_RBUTTONDBLCLK = &H206 'Double-click


Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
Dim X As Long

'Thanks to Stewart Macfarlane at Planet-Source-Code for providing
'me with the below function, which made my work easier to accomplish
Public Function get_filename_only(filepath)
    For X = Len(filepath) To 1 Step -1
        If Mid(filepath, X, 1) = "\" Then
            get_filename_only = Right(filepath, Len(filepath) - X)
            Exit Function
        End If
    Next X
    get_filename_only = "Please check filepath it may be incorrect)"
End Function

