VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSaver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Untitled - Saver 2.0"
   ClientHeight    =   4875
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4350
   Icon            =   "FrmSaver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy Code"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox TxtCode 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Type Code Here"
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox TxtComment 
      Height          =   735
      Left            =   0
      MaxLength       =   165
      MultiLine       =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Type Comments Here"
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Left            =   0
      MaxLength       =   65
      TabIndex        =   0
      ToolTipText     =   "Type Code Name Here"
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label LblCode 
      Caption         =   "Code:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label LblComment 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Code"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Code"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile to Bas"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Code"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Code"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInstruct 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnu_1 
      Caption         =   "mnu_1"
      Visible         =   0   'False
      Begin VB.Menu mnuExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare variables

Dim CodeName As String * 65
Dim CodeComment As String * 165
Dim CodeCode, CodeSave, RetVal As String
Dim Filenum As Integer

Private Sub CmdCopy_Click()
'Copy Code
    Clipboard.SetText (TxtCode.Text)
End Sub

Private Sub Form_Activate()
On Error GoTo Error:

'Sets start up position
FrmSaver.Top = 0
FrmSaver.Left = 0

Error:
    Exit Sub
End Sub

Private Sub Form_Load()
'Makes it so this app cannot be opened more than once at a time
If App.PrevInstance = True Then
    Call MsgBox("Only one instance of Saver 2.0 can be open at once.", vbOKOnly + vbExclamation, "ERROR")
    Unload Me
    End
    Exit Sub
End If

'This will minimize my app when it is first opened
FrmSaver.WindowState = vbMinimized

'This has to do with minimizing to the systems tray
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " Saver 2.0 by: Alwyn B. " & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_LBUTTONDBLCLK
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_RBUTTONUP
        Result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mnu_1
    End Select
End Sub

Private Sub Form_Resize()
'Makes sure minmized window is hidden
'Has to do with minmizeing to the systems tray
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
'This removes the icon from the systems tray
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuAbout_Click()
'Loads the About form
    Load FrmAbout
    FrmAbout.Show
End Sub

Private Sub mnuAbout2_Click()
'Loads the About form
    Load FrmAbout
    FrmAbout.Show
End Sub

Private Sub mnuAll_Click()
'If a textbox has focus, highlight all the text in that textbox(stops errors this way)
If TypeOf FrmSaver.ActiveControl Is TextBox Then
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End If
End Sub

Private Sub mnubas_Click()
'Preload frmComile then show it
    Load FrmCompile
    FrmCompile.Show
End Sub

Private Sub mnuCompile_Click()
'Preload frmComile then show it
    Load FrmCompile
    FrmCompile.Show
End Sub

Private Sub mnuCopy_Click()
'If something is selected in a certain textbox, that's what is copied
    If TxtName.SelLength > 0 Then
        Clipboard.SetText (TxtName.SelText)
    ElseIf TxtComment.SelLength > 0 Then
        Clipboard.SetText (TxtComment.SelText)
    ElseIf TxtCode.SelLength > 0 Then
        Clipboard.SetText (TxtCode.SelText)
    End If
End Sub

Private Sub mnuCut_Click()
'Copy using the mnuCopy_click procedure
    mnuCopy_Click
    
'Now clear that box's selected text
    If TxtName.SelLength > 0 Then
        TxtName.SelText = ""
    ElseIf TxtComment.SelLength > 0 Then
        TxtComment.SelText = ""
    ElseIf TxtCode.SelLength > 0 Then
        TxtCode.SelText = ""
    End If
End Sub

Private Sub mnuDelete_Click()
'Loads the load form using the open_click procedure
'The load form has a delete button on it
    mnuOpen_Click
End Sub

Private Sub mnuExit_Click()
'Unloads program/gives back system resources
    Unload Me
    End
End Sub

Private Sub mnuExit2_Click()
'Unloads form and exits the program
    Unload Me
    End
End Sub

Private Sub mnuInstruct_Click()
'Tells user how to use the software
Call MsgBox("Instructions:" + vbNewLine + vbNewLine + "1.Type a Code Name in the Code Name box" + vbNewLine + "2.Type your comments in the Code Comments box" + vbNewLine + "3.Type(Paste) your code in the Code box" + vbNewLine + "4.Go to File and click save, type a name to save your code under" + vbNewLine + "5.Now go to file and load to load up your code in the future, or new to start a code" + vbNewLine + "6.By clicking on the Convert to Bas option, you just designate the folder you want the bas saved in, and every code you ever made up to that point in time gets compiled into that single bas file", vbOKOnly + vbInformation, "INSTRUCTIONS")
End Sub

Private Sub mnuMinimize_Click()
'Just minimizes the app
'Patricks code will send it to the systems tray
    FrmSaver.WindowState = vbMinimized
End Sub

Private Sub mnuNew_Click()
'Clears form
TxtName = ""
TxtComment = ""
TxtCode.Text = ""

'Lets user know a new code has been made
FrmSaver.Caption = "Untitled - Saver 2.0"
End Sub

Private Sub mnuOpen_Click()
'Loads the load code form
    Load FrmLoad
    FrmLoad.Show
End Sub

Private Sub mnuPaste_Click()
'If textbox has focus, then paste what's on the clipboard(stops error this way)
    If TypeOf FrmSaver.ActiveControl Is TextBox Then
        FrmSaver.ActiveControl.SelText = Clipboard.GetText
    End If
End Sub

Private Sub mnuSave_Click()
On Error GoTo Save_Error:

'Tells if code has been opened/saved already by reading the forms caption
'And if it was opened for editing/saved already it saves automatically without need of the Input Box
'If it is a new code it just keeps going on down the codeing useing the normal procedure to save
If Left$(FrmSaver.Caption, 1) = "C" Then
    CodeName = TxtName.Text
    CodeComment = TxtComment.Text
    CodeCode = TxtCode.Text
    CodeSave = CodeName + CodeComment + CodeCode
    
    Filenum = FreeFile
    Open App.Path + "\" + Mid$(FrmSaver.Caption, 11, InStr(1, FrmSaver.Caption, " ") - 1) + ".cod" For Output As Filenum
        Print #Filenum, RTrim(CodeSave)
    Close Filenum
    
    Call SetAttr(App.Path + "\" + Mid$(FrmSaver.Caption, 11, InStr(1, FrmSaver.Caption, " ") - 1) + ".cod", vbHidden)
    Call MsgBox("Saver 2.0 remembered your last save name for this code and has saved it automatically under that name for you.", vbOKOnly + vbInformation, "SAVED")
Exit Sub
End If

'User picks a save name, and a little safety codeing in case nothing was typed in
RetVal = InputBox("Give a name for the code to be saved under?" + vbNewLine + "No characters such as /\|}{?", "NAME")
If RetVal = "" Then
    GoTo Save_Error:
End If

'Puts all the textboxes on the forms text into a single variable for saveing
CodeName = TxtName.Text
CodeComment = TxtComment.Text
CodeCode = TxtCode.Text
CodeSave = CodeName + CodeComment + CodeCode

'Now saves the code
Filenum = FreeFile
Open App.Path + "\" + RetVal + ".cod" For Output As Filenum
    Print #Filenum, RTrim(CodeSave)
Close Filenum

'Changes caption so user knows code was saved, and knows what the name of the code is
FrmSaver.Caption = "Code Name:" + RetVal + " - Saver 2.0"

'Let user know codes been deleted
Call MsgBox("Code has been saved.", vbOKOnly + vbInformation, "SAVED")

'Make the .cod file hidden, so it doesn't look tacky on the users computer
Call SetAttr(App.Path + "\" + RetVal + ".cod", vbHidden)

Save_Error:
    Exit Sub
End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    Unload Me
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
Dim Result
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub mnuShow_Click()
    FrmSaver.WindowState = vbNormal
End Sub
