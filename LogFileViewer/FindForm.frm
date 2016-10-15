VERSION 5.00
Begin VB.Form FindForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   Icon            =   "FindForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox WholeWordEdit 
      Caption         =   "Whole word"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame DirectionFrame 
      Caption         =   "Direction"
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton UpEdit 
         Caption         =   "Up"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton DownEdit 
         Caption         =   "Down"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CheckBox MatchCaseEdit 
      Caption         =   "Match case"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TextEdit 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label FindLabel 
      Caption         =   "Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Public Event Find()
Public Event FindNext()
Private m_NewSearch As Boolean

'CSEH: ShowErr
Private Sub CancelButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call Hide

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FindNextButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    If m_NewSearch Then
        RaiseEvent Find
        m_NewSearch = False
    Else
        RaiseEvent FindNext
    End If
    
    Call ActivateTextBox(TextEdit)
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub Form_Activate()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call ActivateTextBox(TextEdit)
    FindNextButton.Enabled = (Len(TextEdit.Text) > 0)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub TextEdit_Change()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    m_NewSearch = True
    FindNextButton.Enabled = (Len(TextEdit.Text) > 0)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub MatchCaseEdit_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    m_NewSearch = True

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub WholeWordEdit_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    m_NewSearch = True

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub UpEdit_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    m_NewSearch = True

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub DownEdit_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    m_NewSearch = True

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

