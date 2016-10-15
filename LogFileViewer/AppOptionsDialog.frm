VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form AppOptionsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "AppOptionsDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MaxFileSizeEdit 
      Height          =   285
      Left            =   2295
      TabIndex        =   12
      Text            =   "???"
      Top             =   4260
      Width           =   735
   End
   Begin VB.CommandButton LoadDefaultsButton 
      Caption         =   "Load &Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame DefaultLogFileFrame 
      Caption         =   "Log file to load when program starts"
      Height          =   3975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6975
      Begin VB.CheckBox ReloadLogFileEdit 
         Caption         =   "Open file when close this dialog"
         Height          =   195
         Left            =   3240
         TabIndex        =   0
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton AbsLogFileNameUseActiveButton 
         Caption         =   "Use Current File"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Browse"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox DefLogFileEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   360
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "<default log file name>"
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton AbsLogFileNameBrowseButton 
         Caption         =   "Browse"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Browse"
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox LogFileNameCallTypeEdit 
         Height          =   315
         ItemData        =   "AppOptionsDialog.frx":000C
         Left            =   5040
         List            =   "AppOptionsDialog.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox AbsLogFileNameEdit 
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   6375
      End
      Begin VB.TextBox LogFileNameProcNameEdit 
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox LogFileNameProgidEdit 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   2535
      End
      Begin VB.OptionButton ReadFromObjOpt 
         Caption         =   "Read from object"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton AbsPathOpt 
         Caption         =   "Absolute path"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton TempDirOpt 
         Caption         =   "In current user's temp directory"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox TempDirLogFileNameEdit 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selected path:"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label LogFileStatusLabel 
         AutoSize        =   -1  'True
         Caption         =   "<file status>"
         Height          =   195
         Left            =   5880
         TabIndex        =   20
         Top             =   600
         Width           =   840
      End
      Begin VB.Label CallTypeLabel 
         AutoSize        =   -1  'True
         Caption         =   "Call type:"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   3240
         Width           =   645
      End
      Begin VB.Label ProcNameLabel 
         AutoSize        =   -1  'True
         Caption         =   "Procedure name:"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Class PROGID:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog LogFileDlg 
      Left            =   6120
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "csv"
      Filter          =   "Log File (*.log, *.csv, *.txt)|*.log;*.csv;*.txt|Any (*.*)|*.*"
      FilterIndex     =   1
   End
   Begin MSComDlg.CommonDialog ExportFileDlg 
      Left            =   6600
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "log"
      Filter          =   "Tab-delimited (*.log)|*.log|Comma Separated (*.csv)|*.csv|XML (*.xml)|*.xml"
      FilterIndex     =   1
   End
   Begin VB.Label MaxFileSizeLabel 
      AutoSize        =   -1  'True
      Caption         =   "Maximum File Size (KB):"
      Height          =   195
      Left            =   480
      TabIndex        =   22
      Top             =   4320
      Width           =   1680
   End
End
Attribute VB_Name = "AppOptionsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private m_Cancel As Boolean
Private Const BASE_FILE_DLG_OPTIONS As Long = _
 MSComDlg.cdlOFNFileMustExist + _
 MSComDlg.cdlOFNHideReadOnly + _
 MSComDlg.cdlOFNPathMustExist + _
 MSComDlg.cdlOFNOverwritePrompt

Private Sub AbsLogFileNameEdit_Change()
    Call DefaultLogFileNameChanged
    AbsLogFileNameEdit.ToolTipText = AbsLogFileNameEdit.Text
    Call EnableOK
End Sub

Private Sub DefLogFileEdit_Change()
    DefLogFileEdit.ToolTipText = DefLogFileEdit.Text
End Sub

Private Sub DefLogFileEdit_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'CSEH: ShowErr
Private Sub AbsLogFileNameBrowseButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Dim FileName As String
    FileName = BrowseToOpenLogFile(AbsLogFileNameEdit.Text)
    If FileName = "" Then Exit Sub
    AbsLogFileNameEdit.Text = FileName
    Call ActivateTextBox(AbsLogFileNameEdit)
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub AbsLogFileNameUseActiveButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    AbsLogFileNameEdit.Text = ViewerWindow.LogViewUpdater.FilePath
    Call ActivateTextBox(AbsLogFileNameEdit)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub CancelButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    m_Cancel = True
    Call Hide

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Updates the display for the default log file name.
'
Private Sub DefaultLogFileNameChanged()
    
    On Error GoTo ERROR_HANDLER
    
    Dim FileName As String
    FileName = GetDefaultLogFilename
    
    DefLogFileEdit.Text = FileName
    
    If FSO.FileExists(FileName) Then
        DefLogFileEdit.BackColor = vbGreen
        LogFileStatusLabel.Caption = "File Exists"
    Else
        DefLogFileEdit.BackColor = vbButtonFace
        LogFileStatusLabel.Caption = "Not Found"
    End If

    Exit Sub
ERROR_HANDLER:
    DefLogFileEdit.BackColor = vbRed
    DefLogFileEdit.Text = Err.Description
    LogFileStatusLabel.Caption = "Error"
    
End Sub

'CSEH: ShowErr
Private Sub LoadDefaultsButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim DefOpt As New AppOptions
    Call LoadFromObject(DefOpt)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

Private Sub LogFileNameCallTypeEdit_Click()
    Call DefaultLogFileNameChanged
End Sub

Private Sub LogFileNameProcNameEdit_Change()
    Call DefaultLogFileNameChanged
    Call EnableOK
End Sub

Private Sub LogFileNameProgidEdit_Change()
    Call DefaultLogFileNameChanged
    Call EnableOK
End Sub

'CSEH: ShowErr
Private Sub OkButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    If Not ValidateInput Then Exit Sub
    Call SaveToObject
    Call Hide
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

Private Sub TempDirLogFileNameEdit_Change()
    Call DefaultLogFileNameChanged
    Call EnableOK
End Sub

Private Sub AllowNumericKeysOnly(KeyAscii As Integer)

On Error GoTo ERROR_HANDLER
    
    If KeyAscii <> Asc(vbBack) And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
        VBA.Beep
    End If
    
Exit Sub
ERROR_HANDLER:
    Debug.Assert False

End Sub

Private Sub MaxFileSizeEdit_KeyPress(KeyAscii As Integer)

    Call AllowNumericKeysOnly(KeyAscii)

End Sub

' Shows the dialog with the state of an options object.  If the user selects
' OK, the object state is updated from the state of the dialog controls.
' If the user selects Cancel, the dialog is closed and the object is not
' modified.
'
Public Function Execute() As Boolean
    
    m_Cancel = False
    Call LoadFromObject(g_AppOptions)
    Call EnableControls
    Call Show(vbModal)
    Execute = Not m_Cancel

End Function

' Prompts the user to select a log file and returns the file's path.  Returns
' an empty string if the user cancels.
'
Public Function BrowseToOpenLogFile _
 (Optional ByVal Dir As String = "", _
  Optional ByVal Name As String = "", _
  Optional ByVal Title As String = "") As String
    
    On Error GoTo ERROR_HANDLER
    
    LogFileDlg.Flags = MSComDlg.cdlOFNHideReadOnly ' OK to specify nonexistant file/folder
    LogFileDlg.InitDir = Dir
    LogFileDlg.FileName = Name
    LogFileDlg.DialogTitle = Title
    LogFileDlg.ShowOpen
    BrowseToOpenLogFile = LogFileDlg.FileName
    
    Exit Function
    
ERROR_HANDLER:
    If Err.Number <> MSComDlg.cdlCancel Then Call RaiseErr

End Function

' Prompts the user to select one or more log files and returns a list of the
' file paths as a one dimensional array of strings in a variant.  Returns
' Empty if the user cancels.
'
Public Function BrowseToOpenLogFileList _
 (Optional ByVal Dir As String = "", _
  Optional ByVal Name As String = "", _
  Optional ByVal Title As String = "") As Variant
    
    On Error GoTo ERROR_HANDLER
    
    LogFileDlg.Flags = BASE_FILE_DLG_OPTIONS + cdlOFNAllowMultiselect + cdlOFNExplorer
    LogFileDlg.InitDir = Dir
    LogFileDlg.FileName = Name
    LogFileDlg.DialogTitle = Title
    LogFileDlg.ShowOpen
    
    Dim Strings() As String
    Strings = Split(LogFileDlg.FileName, Chr(0))
    Debug.Assert LBound(Strings) = 0
    Debug.Assert UBound(Strings) >= LBound(Strings)
    
    Dim Result() As String
    If UBound(Strings) = 0 Then
        ReDim Result(0)
        Result(0) = Strings(0)
    Else
        ReDim Result(UBound(Strings) - 1)
        Dim i As Long
        For i = 0 To UBound(Result)
            Result(i) = Strings(0) & "\" & Strings(i + 1)
        Next
    End If
    
    BrowseToOpenLogFileList = Result
    
    Exit Function
    
ERROR_HANDLER:
    If Err.Number <> MSComDlg.cdlCancel Then Call RaiseErr

End Function

' Prompts the user to select a log file and returns the file's path.  Returns
' an empty string if the user cancels.
'
Public Function BrowseToSaveLogFile _
 (Optional ByVal Dir As String = "", _
  Optional ByVal Name As String = "", _
  Optional ByVal Title As String = "") As String

    On Error GoTo ERROR_HANDLER

    LogFileDlg.Flags = BASE_FILE_DLG_OPTIONS
    LogFileDlg.InitDir = Dir
    LogFileDlg.FileName = Name
    LogFileDlg.DialogTitle = Title
    LogFileDlg.ShowSave
    BrowseToSaveLogFile = LogFileDlg.FileName

    Exit Function

ERROR_HANDLER:
    If Err.Number <> MSComDlg.cdlCancel Then Call RaiseErr

End Function

' Prompts the user to select an export file name and returns the file's path.
' Returns an empty string if the user cancels.
'
Public Function BrowseToSaveExportFile _
 (Optional ByVal Dir As String = "", _
  Optional ByVal Name As String = "", _
  Optional ByVal Title As String = "") As String
    
    On Error GoTo ERROR_HANDLER
    
    ExportFileDlg.Flags = BASE_FILE_DLG_OPTIONS
    ExportFileDlg.InitDir = Dir
    ExportFileDlg.FileName = Name
    ExportFileDlg.DialogTitle = Title
    ExportFileDlg.ShowSave
    BrowseToSaveExportFile = ExportFileDlg.FileName
    
    Exit Function
    
ERROR_HANDLER:
    If Err.Number <> MSComDlg.cdlCancel Then Call RaiseErr
    
End Function

' Loads the UI controls from the state of an options object.
'
Private Sub LoadFromObject(ByVal Options As AppOptions)

    Debug.Assert Not Options Is Nothing
    
    With Options
    
        MaxFileSizeEdit.Text = .MaxFileSizeBytes / 1000
        TempDirLogFileNameEdit.Text = .TempDirLogFileName
        AbsLogFileNameEdit.Text = .AbsoluteLogFileName
        LogFileNameProgidEdit.Text = .LogFileNameProgID
        LogFileNameProcNameEdit.Text = .LogFileNameProcName
        
        If .LogFileNameCallType = VBA.VbGet Then
            LogFileNameCallTypeEdit.ListIndex = 0
        ElseIf .LogFileNameCallType = VBA.VbMethod Then
            LogFileNameCallTypeEdit.ListIndex = 1
        Else
            LogFileNameCallTypeEdit.ListIndex = -1
        End If
        
        If .LogFileNameSource = LFS_TEMP Then
            TempDirOpt.Value = True
        ElseIf .LogFileNameSource = LFS_ABSOLUTE Then
            AbsPathOpt.Value = True
        ElseIf .LogFileNameSource = LFS_OBJECT Then
            ReadFromObjOpt.Value = True
        Else
            TempDirOpt.Value = False
            AbsPathOpt.Value = False
            ReadFromObjOpt.Value = False
        End If
        
    End With
    
End Sub

Private Function GetDefaultLogFilename() As String

    If TempDirOpt Then
        
        GetDefaultLogFilename = FSO.BuildPath(FSO.GetSpecialFolder(Scripting.TemporaryFolder), TempDirLogFileNameEdit.Text)
    
    ElseIf AbsPathOpt Then
        
        GetDefaultLogFilename = AbsLogFileNameEdit.Text
    
    ElseIf ReadFromObjOpt Then
        
        Dim Obj As Object
        Set Obj = CreateObject(LogFileNameProgidEdit.Text)
        Dim CallType As VBA.VbCallType
        CallType = VbGet
        If LogFileNameCallTypeEdit.ListIndex = 1 Then CallType = VbMethod
        GetDefaultLogFilename = VBA.CallByName(Obj, LogFileNameProcNameEdit.Text, CallType)
    
    End If

End Function

' Validates that the value of an edit control is a numeric value in a
' specific range of values.  If not, the text of the control is selected
' and an error is raised.
'
Private Sub ValidateNumEdit(ByVal Edit As TextBox, ByVal First As Long, ByVal Last As Long)

    On Error GoTo BAD_VALUE
    If Not IsNumeric(Edit.Text) Then GoTo BAD_VALUE
    Dim IntValue As Long
    IntValue = Edit.Text
    If IntValue < First Or IntValue > Last Then GoTo BAD_VALUE
    Exit Sub
    
BAD_VALUE:
    Call ActivateTextBox(Edit)
    RaiseMsg Edit.Name & ": Bad value '" & Edit.Text & "' -- must be a number from " & First & " to " & Last & "."
    
End Sub

' Validate input values.
'
Private Function ValidateInput() As Boolean

    Call ValidateNumEdit(MaxFileSizeEdit, 100, &H7FFFFFFF / 1000)
    
    Dim FileName As String
    FileName = GetDefaultLogFilename
    If Not FSO.FileExists(FileName) Then
        Dim Msg As String
        Msg = "Log file not found '" & FileName & "'"
        If MsgBox(Msg, vbExclamation + vbOKCancel) = vbCancel Then Exit Function
    End If

    ValidateInput = True
    
End Function

' Loads the state of an options object from the state of the UI controls.
'
Private Sub SaveToObject()

    Debug.Assert Not g_AppOptions Is Nothing
    
    With g_AppOptions
    
        .MaxFileSizeBytes = MaxFileSizeEdit.Text * 1000
        .TempDirLogFileName = TempDirLogFileNameEdit.Text
        .AbsoluteLogFileName = AbsLogFileNameEdit.Text
        .LogFileNameProgID = LogFileNameProgidEdit.Text
        .LogFileNameProcName = LogFileNameProcNameEdit.Text
        
        If LogFileNameCallTypeEdit.ListIndex = 0 Then
            .LogFileNameCallType = VBA.VbGet
        ElseIf LogFileNameCallTypeEdit.ListIndex = 1 Then
            .LogFileNameCallType = VBA.VbMethod
        Else
            RaiseMsg "INTERNAL ERROR: bad call type"
        End If
        
        If TempDirOpt.Value Then
            .LogFileNameSource = LFS_TEMP
        ElseIf AbsPathOpt.Value Then
            .LogFileNameSource = LFS_ABSOLUTE
        ElseIf ReadFromObjOpt.Value Then
            .LogFileNameSource = LFS_OBJECT
        Else
            RaiseMsg "INTERNAL ERROR: bad log file name source"
        End If
        
    End With

End Sub

Private Sub EnableControls()
    
    TempDirLogFileNameEdit.Enabled = TempDirOpt.Value
    AbsLogFileNameEdit.Enabled = AbsPathOpt.Value
    AbsLogFileNameBrowseButton.Enabled = AbsPathOpt.Value
    AbsLogFileNameUseActiveButton.Enabled = AbsPathOpt.Value
    LogFileNameProgidEdit.Enabled = ReadFromObjOpt.Value
    LogFileNameProcNameEdit.Enabled = ReadFromObjOpt.Value
    LogFileNameCallTypeEdit.Enabled = ReadFromObjOpt.Value

    Call DefaultLogFileNameChanged
    
    Call EnableOK

End Sub

Private Sub EnableOK()
    
    If TempDirOpt.Value Then
        OkButton.Enabled = Len(TempDirLogFileNameEdit.Text) > 0
    ElseIf AbsPathOpt.Value Then
        OkButton.Enabled = Len(AbsLogFileNameEdit.Text) > 0
    Else
        OkButton.Enabled = Len(LogFileNameProgidEdit.Text) > 0 And Len(LogFileNameProcNameEdit.Text) > 0
    End If

End Sub

'CSEH: ShowErr
Private Sub ReadFromObjOpt_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call EnableControls

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub AbsPathOpt_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call EnableControls

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub TempDirOpt_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call EnableControls

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

