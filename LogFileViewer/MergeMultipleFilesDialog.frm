VERSION 5.00
Begin VB.Form MergeMultipleFilesDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Merge Multiple Files"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "MergeMultipleFilesDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton RemoveButton 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List 
      Height          =   2400
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&Merge"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label ListLabel 
      AutoSize        =   -1  'True
      Caption         =   "Input Files:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "MergeMultipleFilesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private m_MergedFileName As String

' Shows the dialog.  If the user selects to merge files, then this returns
' the name of the generated file.  Returns an empty string if the user
' chooses to cancel.
'
Public Function Execute() As String
    
    m_MergedFileName = ""
    List.Clear
    If Not AddInputFiles Then Exit Function
    Call EnableControls
    Call Show(vbModal)
    Execute = m_MergedFileName

End Function

Private Sub EnableControls()

    RemoveButton.Enabled = List.SelCount > 0
    OkButton.Enabled = List.ListCount > 1
    
End Sub

' Indicates whether a particular string matches an item in the list control
' ignoring case.
'
Private Function ExistsInList(ByVal Text As String) As Boolean

    Text = UCase$(Text)
    Dim i As Long
    For i = 0 To List.ListCount - 1
        If UCase$(List.List(i)) = Text Then
            ExistsInList = True
            Exit Function
        End If
    Next i

End Function

' Browses for files and adds them to the input file list.  Returns false if
' the user selects to cancel.
'
Private Function AddInputFiles() As Boolean

    Dim FileNameList As Variant
    FileNameList = AppOptionsDialog.BrowseToOpenLogFileList(Title:="Select Input Files")
    If IsEmpty(FileNameList) Then Exit Function
    
    Dim SkippedFiles As New usStringList
    
    Dim FileName As Variant
    For Each FileName In FileNameList
        If ExistsInList(FileName) Then
            Call SkippedFiles.Add(FileName)
        Else
            Call List.AddItem(FileName)
        End If
    Next
    
    If SkippedFiles.Count > 0 Then
        Dim Msg As String
        Msg = "Input file list already contains:"
        Dim i As Long
        For i = 1 To SkippedFiles.Count
            Msg = Msg & vbCrLf & SkippedFiles(i)
        Next
        Call MsgBox(Msg, vbExclamation)
    End If
    
    Call EnableControls
    
    AddInputFiles = True

End Function

'CSEH: ShowErr
Private Sub AddButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call AddInputFiles
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub List_Click()
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
Private Sub RemoveButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim i As Long
    While i < List.ListCount
        If List.Selected(i) Then
            Call List.RemoveItem(i)
        Else
            i = i + 1
        End If
    Wend
    
    Call EnableControls

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Merges the selected files into a single new file.
'
Private Sub MergeFiles()
        
    ' open a new file
    m_MergedFileName = AppOptionsDialog.BrowseToSaveLogFile(Title:="Save Merged File As")
    If m_MergedFileName = "" Then RaiseCancel
    If ExistsInList(m_MergedFileName) Then
        Call RaiseMsg("'" & m_MergedFileName & "' is one of the input files.  Please select a different merge (output) file.")
    End If
    Dim MergedFile As Scripting.TextStream
    Set MergedFile = FSO.CreateTextFile(m_MergedFileName, Unicode:=True)
    
    ' copy lines of input files into new file
    Dim BaseFirstLine As String
    Dim FirstLine As String
    Dim i As Long
    For i = 0 To List.ListCount - 1
    
        Dim InFilename As String
        InFilename = List.List(i)
        Dim TestFmt As Scripting.Tristate
        TestFmt = Scripting.TristateUseDefault ' open as unicode or ASCII
        Dim InFile As Scripting.TextStream
        Set InFile = FSO.OpenTextFile(InFilename, Format:=TestFmt)
        If Not InFile.AtEndOfStream Then
        
            ' read first line of input file
            FirstLine = InFile.ReadLine
            
            If i = 0 Then
                
                ' store this as the base first-line
                BaseFirstLine = FirstLine
                Call MergedFile.WriteLine(FirstLine)
            
            Else
                
                ' ask user to continue if this file's first-line does not match base
                If FirstLine <> BaseFirstLine Then
                    Dim Msg As String
                    Msg = _
                      "The first line of '" & InFilename & _
                      "' does not match the first line of the first input file '" & List.List(0) & _
                      "' which indicates that the files may not have compatible formats.  " & _
                      "Do you want to continue merging the files?"
                    If MsgBox(Msg, vbQuestion + vbYesNo) = vbNo Then
                        Call MergedFile.Close
                        Call FSO.DeleteFile(m_MergedFileName)
                        Call RaiseCancel
                    End If
                    Call MergedFile.WriteLine(FirstLine)
                End If
            
            End If
                        
            ' copy rest of input file to output file
            While Not InFile.AtEndOfStream
                Call MergedFile.WriteLine(InFile.ReadLine)
            Wend
            
        End If
        
    Next
    
    Call MergedFile.Close
    
End Sub

'CSEH: ShowErr
Private Sub OkButton_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    On Error GoTo CUSTOM_ERROR_HANDLER
    
    Call MergeFiles
    Call Hide
    
    Exit Sub

CUSTOM_ERROR_HANDLER:
    
    If ErrIsCancel Then Exit Sub
    GoTo ERROR_HANDLER

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

    m_MergedFileName = ""
    Call Hide
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

