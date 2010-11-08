VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Saver 2.0"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdInfo 
      Caption         =   "System Info"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblBugs 
      Caption         =   "This is freeware.  Feel free to distribute this software as you please.  No WARRANTY."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblDescription 
      Caption         =   $"FrmAbout.frx":014A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblTitle 
      Caption         =   "Saver 2.0   By: Alwyn B."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image ImgPic 
      Height          =   1215
      Left            =   120
      Picture         =   "FrmAbout.frx":0211
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare variables

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim rc As Long
Dim SysInfoPath As String
Dim i As Long
Dim hKey As Long
Dim hDepth As Long
Dim KeyValType As Long
Dim tmpVal As String
Dim KeyValSize As Long


Private Sub CmdInfo_Click()
'Shows Info about your computer
Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
'Unload this form and return system resources
Unload Me
End Sub

Private Sub Form_Load()
'Sets where the form loads up at
FrmAbout.Top = FrmSaver.Top
FrmAbout.Left = FrmSaver.Left
End Sub

Public Sub StartSysInfo()
'This is the StartSysInfo sub
'I got help for this from one of Microsofts example form
'I have no idea what this means, only know what it does
    On Error GoTo SysInfoErr
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        Else
            GoTo SysInfoErr
        End If
        Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
    
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'This is a function for the code above
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        KeyVal = Format$("&h" + KeyVal)
    End Select
    GetKeyValue = True
    rc = RegCloseKey(hKey)
    Exit Function
    
GetKeyError:
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)
End Function


