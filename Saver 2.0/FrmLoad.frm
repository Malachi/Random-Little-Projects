VERSION 5.00
Begin VB.Form FrmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load/Delete Code:"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "FrmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox LstCodes 
      Height          =   3180
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.FileListBox Files 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   0
      Pattern         =   "*.cod"
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare variables

Dim X As Long
Dim WhatToAdd As String
Dim Filenum As Long

Private Sub CmdDelete_Click()
On Error GoTo Del_Error:

'Before you can delete the file, the file must be unhidden, this unhides the file
    Call SetAttr(App.Path + "\" + LstCodes.Text + ".cod", vbNormal)

'Delete selected code(file)
    Kill (App.Path + "\" + LstCodes.Text + ".cod")
    
'Removes that selected code from list after it's deleted
    LstCodes.RemoveItem (LstCodes.ListIndex)

'Let user know the code has been deleted
    Call MsgBox("Code has been deleted!", vbOKOnly + vbInformation, "DELETED")

Del_Error:
If Err.Number = 53 Then
    Call MsgBox("Code file cannot be found to delete it!", vbOKOnly + vbExclamation, "ERROR")
Else
    Exit Sub
End If
End Sub

Private Sub CmdOpen_Click()
On Error GoTo Load_Error:

'Load the code user selected
    Filenum = FreeFile
    Open App.Path + "\" + LstCodes.Text + ".cod" For Input As Filenum
        FrmSaver.TxtName.Text = RTrim(Input(65, Filenum))
        FrmSaver.TxtComment.Text = RTrim(Input(165, Filenum))
        FrmSaver.TxtCode.Text = RTrim(Input(LOF(Filenum) - 230, Filenum))
    Close Filenum

'Change caption so user knows what they are working with
    FrmSaver.Caption = "Code Name:" + LstCodes.Text + " - Saver 1.0"

'Hide instead of Unload this form so it shows easy the next time
    FrmLoad.Hide
    
Load_Error:
    Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next

'Sets where this form is loaded
    FrmLoad.Top = FrmSaver.Top
    FrmLoad.Left = FrmSaver.Left

'Adds all .cod files to FileListBox(behind the listbox is the file list box)
    Files.Path = App.Path

'Transfers what was in the FileListBox to the LstCodes list box
    For X = 0 To Files.ListCount
        Call LstCodes.AddItem(Files.List(X), X)
        LstCodes.ListIndex = X
        WhatToAdd = Left(LstCodes.Text, InStr(1, LstCodes.Text, ".") - 1)
        LstCodes.RemoveItem (X)
        Call LstCodes.AddItem(WhatToAdd, X)
    Next X
    LstCodes.RemoveItem (LstCodes.ListCount - 1)
End Sub

