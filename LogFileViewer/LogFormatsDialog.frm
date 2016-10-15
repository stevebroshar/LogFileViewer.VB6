VERSION 5.00
Begin VB.Form LogFormatsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Formats"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "LogFormatsDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CopyButton 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Creates a new format based on the selected format"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton UseButton 
      Caption         =   "&Use"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Applies the selected format to the view"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton CloseButton 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Deletes the selected formats"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton EditButton 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Opens the format editor for the selected format"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton NewButton 
      Caption         =   "&New"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Creates a new format based on the open file"
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox FormatList 
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "LogFormatsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private Function GetSelectedIndex() As Long
    
    Debug.Assert FormatList.SelCount = 1
    Dim i As Long
    For i = 0 To FormatList.ListCount - 1
        If FormatList.Selected(i) Then
            GetSelectedIndex = i
            Exit Function
        End If
    Next

End Function

Private Function GetSelectedFormat() As LogFormat
    
    Set GetSelectedFormat = g_AppOptions.LogFormats.Item(GetSelectedIndex + 1)

End Function

Private Sub LoadControls()

    Call FormatList.Clear
    Dim Fmt As LogFormat
    For Each Fmt In g_AppOptions.LogFormats
        Call FormatList.AddItem(Fmt.Name)
        If Fmt Is ViewerWindow.LogViewUpdater.LogFormat Then
            FormatList.Selected(FormatList.ListCount - 1) = True
        End If
    Next

End Sub

Private Sub EnableControls()

    UseButton.Enabled = FormatList.SelCount = 1
    CopyButton.Enabled = FormatList.SelCount = 1
    EditButton.Enabled = FormatList.SelCount = 1
    DeleteButton.Enabled = FormatList.SelCount > 0
    
End Sub

'CSEH: ShowErr
Private Sub CloseButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call Unload(Me)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FormatList_Click()
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

'CSEH: Skip
Private Sub FormatList_DblClick()
    Call EditButton_Click
End Sub

'CSEH: ShowErr
Private Sub UseButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    ViewerWindow.LogViewUpdater.LogFormat = GetSelectedFormat()
    Call Me.Hide
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub NewButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim LogFmt As LogFormat
    Set LogFmt = g_AppOptions.LogFormats.AddNew
    If LogFormatPropsDialog.Execute(LogFmt) Then
        Call FormatList.AddItem(LogFmt.Name)
        Call LogFmt.Save
    Else
        Call g_AppOptions.LogFormats.Remove(g_AppOptions.LogFormats.Count)
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub CopyButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim LogFmt As LogFormat
    Set LogFmt = g_AppOptions.LogFormats.AddCopy(GetSelectedFormat)
    If LogFormatPropsDialog.Execute(LogFmt) Then
        Call FormatList.AddItem(LogFmt.Name)
        Call LogFmt.Save
    Else
        Call g_AppOptions.LogFormats.Remove(g_AppOptions.LogFormats.Count)
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub EditButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim SelIndex As Long
    SelIndex = GetSelectedIndex
    Dim LogFmt As LogFormat
    Set LogFmt = g_AppOptions.LogFormats.Item(SelIndex + 1)
    If LogFormatPropsDialog.Execute(LogFmt) Then
        FormatList.List(SelIndex) = LogFmt.Name
        Call LogFmt.Save
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub DeleteButton_Click()
    
    Dim SelNames As New usStringList
    Dim SelIndicies As New usLongList
    Dim i As Long
    For i = 0 To FormatList.ListCount - 1
        If FormatList.Selected(i) Then
            Call SelNames.Add(g_AppOptions.LogFormats.Item(i + 1).Name)
            Call SelIndicies.Add(i)
        End If
    Next
    
    Dim Msg As String
    Msg = "The following formats will be permanently deleted: '" & SelNames.Join("', '") & "'."
    If MsgBox(Msg, vbOKCancel + vbExclamation) <> vbOK Then Exit Sub

    For i = SelNames.Count To 1 Step -1
        Call FormatList.RemoveItem(SelIndicies(i))
        Dim LogFmt As LogFormat
        Set LogFmt = g_AppOptions.LogFormats.Item(SelIndicies(i) + 1)
        Call LogFmt.Delete
        Call g_AppOptions.LogFormats.Remove(SelIndicies(i) + 1)
    Next

End Sub

Public Sub Execute()
    
    Call LoadControls
    Call EnableControls
    Call Show(vbModal)
    Call Unload(Me)

End Sub

