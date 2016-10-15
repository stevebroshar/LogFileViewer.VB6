VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form LogFormatPropsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Format Properties"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "LogFormatPropsDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox ColCaptionsEdit 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3625
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"LogFormatPropsDialog.frx":000C
   End
   Begin VB.CommandButton UseCurrentColsButton 
      Caption         =   "Load From Current File"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox ColDelimEdit 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox ColLayoutEdit 
      Height          =   315
      ItemData        =   "LogFormatPropsDialog.frx":008E
      Left            =   120
      List            =   "LogFormatPropsDialog.frx":009E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CheckBox HasHeaderEdit 
      Caption         =   "First Line of File is Header"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox NameEdit 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label ColDelimLabel 
      AutoSize        =   -1  'True
      Caption         =   "Delimiter:"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label ColCaptionsLabel 
      AutoSize        =   -1  'True
      Caption         =   "Column Captions:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label ColLayoutLabel 
      AutoSize        =   -1  'True
      Caption         =   "Column Layout:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label NameLabel 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "LogFormatPropsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private m_Cancel As Boolean
Private m_Fmt As LogFormat
Private Enum COLUMN_LAYOUT_LIST_INDEX
    CLLI_NONE
    CLLI_TAB
    CLLI_STRING
    CLLI_CSV
End Enum

Private Function CLLIfromCL(ByVal CL As COLUMN_LAYOUT) As Long

    If CL = CL_NONE Then
        CLLIfromCL = CLLI_NONE
    ElseIf CL = CL_TAB Then
        CLLIfromCL = CLLI_TAB
    ElseIf CL = CL_STRING Then
        CLLIfromCL = CLLI_STRING
    ElseIf CL = CL_CSV Then
        CLLIfromCL = CLLI_CSV
    Else
        RaiseMsg "INTERNAL ERROR: Unknown column layout."
    End If
    
End Function

Private Function CLfromCLLI(ByVal CLLI As Long) As COLUMN_LAYOUT

    If CLLI = CLLI_NONE Then
        CLfromCLLI = CL_NONE
    ElseIf CLLI = CLLI_TAB Then
        CLfromCLLI = CL_TAB
    ElseIf CLLI = CLLI_STRING Then
        CLfromCLLI = CL_STRING
    ElseIf CLLI = CLLI_CSV Then
        CLfromCLLI = CL_CSV
    Else
        RaiseMsg "INTERNAL ERROR: Unknown column layout list index."
    End If

End Function

' Reads the column captions from the edit control -- ignoring the last item
' if it is blank.
'
Private Function GetRowCaptions() As usStringList
    
    Dim Result As New usStringList
    Result.AsVariant = Split(ColCaptionsEdit.Text, vbCrLf)
    If Result.Count > 0 Then
        If Len(Result(Result.Count)) = 0 Then
            Call Result.Remove(Result.Count)
        End If
    End If
    Set GetRowCaptions = Result

End Function

Private Sub UpdateColCaptionsLabel()
    
    ColCaptionsLabel.Caption = "Column Captions (" & CStr(GetRowCaptions.Count) & "): "

End Sub

' Loads the controls from the object state.
'
Private Sub LoadFromObject()

    NameEdit.Text = m_Fmt.Name
    ColLayoutEdit.ListIndex = CLLIfromCL(m_Fmt.ColumnLayout)
    ColDelimEdit.Text = m_Fmt.ColumnDelimiter
    ColCaptionsEdit.Text = m_Fmt.ColumnCaptions.AsText
    
    If m_Fmt.HasHeaderLine Then
        HasHeaderEdit.Value = vbChecked
    Else
        HasHeaderEdit.Value = vbUnchecked
    End If
    
End Sub

' Validates the settings in the controls.
'
Private Sub ValidateInput()

    Dim SelName As String
    SelName = Trim$(NameEdit.Text)
    
    If Len(SelName) = 0 Then
        Call ActivateTextBox(NameEdit)
        RaiseMsg "Name is blank."
    End If

    ' validate that format name is unique
    Dim Fmt As LogFormat
    For Each Fmt In g_AppOptions.LogFormats
        If Fmt.Name = SelName And Not Fmt Is m_Fmt Then
            Call ActivateTextBox(NameEdit)
            RaiseMsg "A log format named '" & SelName & "' already exists."
        End If
    Next
    
    ' update name edit if not trimmed
    If NameEdit.Text <> SelName Then NameEdit.Text = SelName
    
    ' validate delimiter string
    If ColLayoutEdit.ListIndex = CLLI_STRING And Len(ColDelimEdit.Text) = 0 Then
        Call ColDelimEdit.SetFocus
        RaiseMsg "Column delimiter must be at least one character long."
    End If

    ' ensure at least one column if column layout
    If ColLayoutEdit.ListIndex <> CLLI_NONE Then
        If GetRowCaptions.Count = 0 Then
            Call ColCaptionsEdit.SetFocus
            RaiseMsg "No columns defined."
        End If
    End If
    
End Sub

' Updates the object state from the settings in the controls.
'
Private Sub SaveToObject()
    
    ' validate input
    ' NOTE: do this before modifying any object state
    Call ValidateInput

    Call m_Fmt.Rename(NameEdit.Text)
    m_Fmt.ColumnLayout = CLfromCLLI(ColLayoutEdit.ListIndex)
    m_Fmt.ColumnDelimiter = ColDelimEdit.Text
    m_Fmt.HasHeaderLine = (HasHeaderEdit.Value = vbChecked)
    m_Fmt.ColumnCaptions.AsText = GetRowCaptions.AsText

End Sub

' Shows the dialog to allow the user to edit a log format.  If the user
' selects OK, then the object state is modified and this returns true.
' If the user selects Cancel, then this returns false without modifying
' the object state.
'
Public Function Execute(ByVal LogFormat As LogFormat) As Boolean
 
    m_Cancel = False
    Set m_Fmt = LogFormat
    Call LoadFromObject
    Call Show(vbModal)
    Call Unload(Me)
    Set m_Fmt = Nothing
    Execute = Not m_Cancel

End Function

'CSEH: ShowErr
Private Sub Form_Activate()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call ActivateTextBox(NameEdit)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub OkButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call SaveToObject
    Call Hide
    
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

'CSEH: ShowErr
Private Sub ColCaptionsEdit_Change()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call UpdateColCaptionsLabel

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub ColLayoutEdit_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    ColDelimEdit.Visible = ColLayoutEdit.ListIndex = CLLI_STRING
    ColDelimLabel.Visible = ColLayoutEdit.ListIndex = CLLI_STRING
    ColCaptionsEdit.Visible = ColLayoutEdit.ListIndex <> CLLI_NONE
    ColCaptionsLabel.Visible = ColLayoutEdit.ListIndex <> CLLI_NONE
    UseCurrentColsButton.Visible = ColLayoutEdit.ListIndex <> CLLI_NONE
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub UseCurrentColsButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim WorkFmt As New LogFormat
    WorkFmt.ColumnDelimiter = ColDelimEdit.Text
    WorkFmt.ColumnLayout = CLfromCLLI(ColLayoutEdit.ListIndex)
    Dim LineItems As Variant
    LineItems = WorkFmt.SplitLine(g_AppOptions.HeadLine)
    ColCaptionsEdit.Text = Join(LineItems, vbCrLf)
    Call ActivateTextBox(ColCaptionsEdit)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub
