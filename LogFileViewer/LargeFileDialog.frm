VERSION 5.00
Begin VB.Form LargeFileDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Large File Warning"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "LargeFileDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton LoadButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton RenameButton 
      Caption         =   "&Rename File"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "LargeFileDialog.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label 
      Caption         =   "<>"
      Height          =   1575
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "LargeFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum LARGE_FILE_DIALOG_RESULT
    LFDR_LOAD
    LFDR_CANCEL
    LFDR_RENAME
End Enum
Private m_Result As LARGE_FILE_DIALOG_RESULT

Public Function Execute _
 (ByVal FilePath As String, _
  ByVal FileSize As Long, _
  ByVal MaxFileSize As Long) As LARGE_FILE_DIALOG_RESULT

    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbArrow)
    
    Dim Msg As String
    Msg = _
      "Loading file '" & FilePath & _
      "' may be very time consuming since it is large -- " & CLng(FileSize / 1000) & _
      " KB." & vbCrLf & vbCrLf & _
      "OK: increase the maximum recommended size for this session and load the file." & vbCrLf & _
      "Cancel: cancel loading the file." & vbCrLf & _
      "Rename File: rename existing file to start logging to new file and clear the display."

    Label.Caption = Msg
    Call Me.Show(vbModal)
    Execute = m_Result

End Function

Private Sub LoadButton_Click()
    m_Result = LFDR_LOAD
    Call Me.Hide
End Sub

Private Sub CancelButton_Click()
    m_Result = LFDR_CANCEL
    Call Me.Hide
End Sub

Private Sub RenameButton_Click()
    m_Result = LFDR_RENAME
    Call Me.Hide
End Sub
