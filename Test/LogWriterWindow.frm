VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form LogWriterWindow 
   Caption         =   "Log Writer"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DeleteFileButton 
      Caption         =   "Delete File"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   8640
      Top             =   1080
   End
   Begin VB.CommandButton BrowseButton 
      Caption         =   "Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDlg 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Log (*.log, *.out, *.csv, *.txt)|*.log;*.out;*.csv;*.txt|Any (*.*)|*.*"
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox IntervalEdit 
      Height          =   285
      Left            =   6480
      TabIndex        =   5
      Text            =   "1000"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox LineEdit 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   120
      TabIndex        =   6
      Text            =   "This is the time for all good men to come to the aid of their party."
      Top             =   600
      Width           =   10695
   End
   Begin VB.CommandButton WriteLineButton 
      Caption         =   "Write Line"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LinesWrittenCaption 
      AutoSize        =   -1  'True
      Caption         =   "Lines Written:"
      Height          =   195
      Left            =   9360
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label LinesWrittenCountLabel 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   10440
      TabIndex        =   8
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Interval (ms):"
      Height          =   195
      Left            =   5520
      TabIndex        =   7
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "LogWriterWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private m_FilePath As String
Private FSO As New Scripting.FileSystemObject

Private Sub ActivateEdit(ByVal Edit As Object)
    
    Edit.SelStart = 0
    Edit.SelLength = Len(LineEdit.Text)
    Edit.SetFocus

End Sub

Private Sub SetFilePath(ByVal FilePath As String)
    
    m_FilePath = FilePath
    Me.Caption = "Log Writer - " & FilePath
    
End Sub

Private Sub WriteLine()
    
    Dim Stream As Scripting.TextStream
    Set Stream = FSO.OpenTextFile(m_FilePath, ForAppending, Create:=True)
    Call Stream.WriteLine(LineEdit.Text)
    Call Stream.Close
    
    LinesWrittenCountLabel.Caption = LinesWrittenCountLabel.Caption + 1
    
End Sub

'CSEH: ShowErr
Private Sub BrowseButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call CommonDlg.ShowOpen
    If CommonDlg.Filename <> "" Then
        Call SetFilePath(CommonDlg.Filename)
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub DeleteFileButton_Click()
    
    If FSO.FileExists(m_FilePath) Then
        Dim mbResult As VbMsgBoxResult
        mbResult = MsgBox("Are you sure you want to delete '" & m_FilePath & "'?", vbQuestion + vbYesNo)
        If mbResult = vbYes Then
            Call FSO.DeleteFile(m_FilePath)
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Call SetFilePath(App.Path & "\Test.log")
    IntervalEdit_LostFocus

End Sub

Private Sub IntervalEdit_LostFocus()

    On Error GoTo BAD_INTERVAL
    
    Timer.Interval = CLng(IntervalEdit.Text)
    
    Exit Sub

BAD_INTERVAL:
    
    IntervalEdit.Text = "1000"
    Resume

End Sub

'CSEH: ShowErr
Private Sub StartButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    StopButton.Enabled = True
    StartButton.Enabled = False
    Timer.Enabled = True
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub StopButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Timer.Enabled = False
    StartButton.Enabled = True
    StopButton.Enabled = False

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub Timer_Timer()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call WriteLine

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub WriteLineButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call WriteLine
    Call ActivateEdit(LineEdit)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub
