VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCompile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile To Bas - Saver 2.0"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "FrmCompile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstCodes1 
      Height          =   2400
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1
   End
   Begin VB.CommandButton CmdCompile 
      Caption         =   "Compile"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   ">"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.ListBox LstCodes2 
      Height          =   2400
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.FileListBox FileCodes 
      Height          =   285
      Left            =   600
      Pattern         =   "*.cod"
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblCompile 
      Caption         =   "Codes To Compile:"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label LblAvailable 
      Caption         =   "Available Codes:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblInstruct 
      Caption         =   "Add the codes that you want compiled into your bas to the list on the right.  Then click the compile button."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare variables
 
Dim Filter, CodeAdd, WhatToAdd As String
Dim Filenum, FileNum2 As Integer
Dim X As Currency
Dim Num As Byte

Private Sub CmdAdd_Click()
On Error Resume Next

'If user added all the codes to lstcodes2's list
'Then this code will add a blank to lstcodes2, this stops that bug
If LstCodes1.ListIndex = -1 Then
    Exit Sub
End If

'If the codes not in the list it'll add it again
Call LstCodes2.AddItem(LstCodes1.Text)
 
'Then it removes that from the list
Call LstCodes1.RemoveItem(LstCodes1.ListIndex)
End Sub

Private Sub CmdAddAll_Click()
On Error Resume Next

'This code will clear lstcodes1 and add all of it's contents to lstcodes2
For X = 0 To LstCodes1.ListCount
    Call LstCodes2.AddItem(LstCodes1.Text)
    LstCodes1.ListIndex = LstCodes1.ListIndex + 1
    If LstCodes2.Text = "" Then
        LstCodes2.RemoveItem (X)
    End If
Next X

LstCodes1.Clear
End Sub

Private Sub CmdCompile_Click()
On Error Resume Next

'Allows user to use save dialog to let the program know where to make the bas file
Filter = "Bas Files(*.bas)|*.bas; |"
CommonDialog1.Filter = Filter
CommonDialog1.ShowSave

'Adds a code so vb recognizes it as a .bas file
'And also adds option explicit at the top
Filenum = FreeFile
Open CommonDialog1.FileName For Output As Filenum
    Print #Filenum, "Attribute VB_Name = " + Chr$(34) + Left$(get_filename_only(CommonDialog1.FileName), InStr(1, get_filename_only(CommonDialog1.FileName), ".") - 1) + Chr$(34)
    Print #Filenum, "Option Explicit" + vbNewLine
Close Filenum

'Sets the selected listindex
LstCodes2.ListIndex = LstCodes2.ListCount - 1

'The for next statement makes sure all the codes are added
For X = 1 To LstCodes2.ListCount

'This will extract the code from the selected .cod file and put it in the CodeAdd variable
FileNum2 = FreeFile
Num = 140
Open App.Path + "\" + LstCodes2.Text + ".cod" For Input As FileNum2
    CodeAdd = Input(230, FileNum2)
    CodeAdd = Input(LOF(FileNum2) - 230, FileNum2)
Close FileNum2

'This will add the Code/CodeAdd variable to the .bas file
Filenum = FreeFile
Open CommonDialog1.FileName For Append As Filenum
    Print #Filenum, CodeAdd
Close Filenum

'This makes the for next continue the same procedure on with the next code
LstCodes2.ListIndex = LstCodes2.ListIndex - 1
Next X

'Hide instead of unload the form, so next time the form shows up quick
FrmCompile.Hide
End Sub

Private Sub CmdRemove_Click()
On Error Resume Next

'If user added all the codes to lstcodes1's list from lstcodes2
'Then this code will add a blank to lstcodes1, this stops that bug
If LstCodes2.ListIndex = -1 Then
    Exit Sub
End If

'Add's that code back to lstcodes1's list
Call LstCodes1.AddItem(LstCodes2.Text)

'Removes the selected item the lstcodes2 list
Call LstCodes2.RemoveItem(LstCodes2.ListIndex)
End Sub

Private Sub CmdRemoveAll_Click()
On Error Resume Next

'This code will clear lstcodes2 and add all of it's contents to lstcodes1
For X = 0 To LstCodes2.ListCount
    Call LstCodes1.AddItem(LstCodes2.Text)
    LstCodes2.ListIndex = LstCodes2.ListIndex + 1
    If LstCodes1.Text = "" Then
    LstCodes1.RemoveItem (X)
    End If
Next X

LstCodes2.Clear
End Sub

Private Sub Form_Load()
On Error Resume Next

'Add all codes to filelistbox
FileCodes.Path = App.Path

'Transfers what was in the FileListBox to the LstCodes1 list box
'Then gets rid of the .cod extention at the end
FileCodes.ListIndex = FileCodes.ListCount - 1
For X = 0 To FileCodes.ListCount
    Call LstCodes1.AddItem(FileCodes.FileName, 0)
    FileCodes.ListIndex = FileCodes.ListIndex - 1
Next X
For X = 0 To FileCodes.ListCount
    LstCodes1.ListIndex = X
    WhatToAdd = Left(LstCodes1.Text, InStr(1, LstCodes1.Text, ".") - 1)
    LstCodes1.RemoveItem (X)
    Call LstCodes1.AddItem(WhatToAdd, X)
Next X
    LstCodes1.RemoveItem (0)
    
'Sets where this form starts up at
FrmCompile.Top = FrmSaver.Top
FrmCompile.Left = FrmSaver.Left
End Sub
