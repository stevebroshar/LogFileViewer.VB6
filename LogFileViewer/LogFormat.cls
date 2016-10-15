VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Public Enum COLUMN_LAYOUT
    CL_NONE = 0
    CL_TAB = 1
    CL_STRING = 2
    CL_CSV = 3
End Enum

Private m_Name As String
Private m_FileHasHeaderLine As Boolean
Private m_ColumnLayout As COLUMN_LAYOUT
Private m_ColumnDelimiter As String
Private m_ColumnCaptions As New usStringList
Private m_ColumnWidths As New usLongList
Private m_CsvSplitter As New usStringList
'Private m_IdentHeaderLines As New usObjectCollection 'of usStringList
'Private m_IdentFileNamePatterns As New usStringList

Private Sub Class_Initialize()
    Call Clear
End Sub

Public Sub Clear()
    m_Name = ""
    m_ColumnLayout = CL_NONE
    m_ColumnDelimiter = ","
    m_FileHasHeaderLine = True
    Call m_ColumnWidths.Clear
    'Call m_IdentHeaderLines.Clear
    'Call m_IdentFileNamePatterns.Clear
End Sub

' Display name of the format.
'
Public Property Get Name() As String
    Name = m_Name
End Function

Private Function FindUniqueName(ByVal BaseName As String) As String
    Dim CandidateName As String
    CandidateName = BaseName
    Dim i As Long
    Do
        If Not FSO.FileExists(g_AppOptions.LogFormatsPath(CandidateName)) Then
            FindUniqueName = CandidateName
            Exit Function
        End If
        i = i + 1
        CandidateName = BaseName + Str(i)
    Loop
End Function

Public Sub SetNameUnique(ByVal Name As String)
    m_Name = FindUniqueName(Name)
End Sub

' Indicates whether the first line of the log file is the header.
'
Public Property Get HasHeaderLine() As Boolean
    HasHeaderLine = m_FileHasHeaderLine
End Property
Public Property Let HasHeaderLine(ByVal RHS As Boolean)
    m_FileHasHeaderLine = RHS
End Property

' The column layout type.
'
Public Property Get ColumnLayout() As COLUMN_LAYOUT
    ColumnLayout = m_ColumnLayout
End Property
Public Property Let ColumnLayout(ByVal RHS As COLUMN_LAYOUT)
    m_ColumnLayout = RHS
End Property

' The deliminiter used in the file to separate columns.  Only used if
' layout is CL_STRING.
'
Public Property Get ColumnDelimiter() As String
    ColumnDelimiter = m_ColumnDelimiter
End Property
Public Property Let ColumnDelimiter(ByVal RHS As String)
    m_ColumnDelimiter = RHS
End Property

' The column names to use for both identifing a file's format and for
' showing in the view.
'
Public Property Get ColumnCaptions() As usStringList
    Set ColumnCaptions = m_ColumnCaptions
End Property

' The text of the header line which consists of the column captions
' concatenated together with the column-deliminiter as separator.
'
Public Property Get HeaderText() As String
    HeaderText = JoinColumnsOfLine(m_ColumnCaptions)
End Property

' Returns a string that is the concatenation of the string list items
' separated with the format's separator.
'
Private Function JoinColumnsOfLine(ByVal ColumnStrings As usStringList) As String
    If m_ColumnLayout = CL_TAB Then
        JoinColumnsOfLine = ColumnStrings.Join(vbTab)
    ElseIf m_ColumnLayout = CL_STRING Then
        JoinColumnsOfLine = ColumnStrings.Join(m_ColumnDelimiter)
    ElseIf m_ColumnLayout = CL_CSV Then
        JoinColumnsOfLine = ColumnStrings.AsCSV
    ElseIf m_ColumnLayout = CL_NONE Then
        RaiseMsg "There is no column delimiter for the no-column layout."
    Else
        RaiseMsg "INTERNAL ERROR: Unexpected column layout type."
    End If
End Function

' Loads a 0-based variant array with the column items of a log line.
'
' Text is readonly and passed by-ref for performance.
'
' This is more efficient than SplitLine since it usually results in one less
' copy of the array.
'
Public Sub SplitLineToArray(ByRef Text As String, ByRef Items As Variant)
    If m_ColumnLayout = CL_TAB Then
        Items = Split(Text, vbTab)
    ElseIf m_ColumnLayout = CL_STRING Then
        Items = Split(Text, m_ColumnDelimiter)
    ElseIf m_ColumnLayout = CL_CSV Then
        If Len(Text) > 0 Then
            m_CsvSplitter.AsCSV = Text
            Call m_CsvSplitter.CopyToVariant(Items, 0)
        Else
            Items = Split("", " ")
        End If
    ElseIf m_ColumnLayout = CL_NONE Then
        RaiseMsg "There is no column delimiter for the no-column layout."
    Else
        RaiseMsg "INTERNAL ERROR: Unexpected column layout type."
    End If
End Sub

' Returns a 0-based variant array loaded with the column items of a log line.
'
' Text is readonly and passed by-ref for performance.
'
' This code could be written in a cleaner way if the result was usStringList,
' but that would have significant performance impact for non-CSV formats.
'
Public Function SplitLine(ByRef Text As String) As Variant
    Dim Result() As String
    Call SplitLineToArray(Text, Result)
    SplitLine = Result
End Function

' The column widths of the view stored for this format.
'
Public Property Get ColumnWidths() As usLongList
    Set ColumnWidths = m_ColumnWidths
End Property

'' Headers that identify the format.
''
'Public Property Get IdentifyingHeaders() As usObjectCollection
'    Set IdentifyingHeaders = m_IdentHeaderLines
'End Property
'
'' File patterns that identify the format.
''
'Public Property Get IdentifyingFileNamePatterns() As usStringList
'    Set IdentifyingFileNamePatterns = m_IdentFileNamePatterns
'End Property

Private Function SaveToStringList() As usStringList
    Dim Lines As New usStringList
    Lines.Value("ColumnLayout") = m_ColumnLayout
    Lines.Value("ColumnDelimiter") = m_ColumnDelimiter
    Lines.Value("ColumnWidths") = Join(m_ColumnWidths.AsVariant, ",")
    Lines.Value("ColumnCaptions") = m_ColumnCaptions.AsCSV
    Lines.Value("HasHeaderLine") = CStr(CLng(m_FileHasHeaderLine))
    ''Lines.Value("HeaderLines") = m_IdentHeaderLines.AsCSV
    ''Lines.Value("FileNamePatterns") = m_IdentFileNamePatterns.AsCSV
    Set SaveToStringList = Lines
End Function

Private Sub LoadFromStringList(ByVal Lines As usStringList)
    m_ColumnLayout = Lines.Value("ColumnLayout")
    m_ColumnDelimiter = Lines.Value("ColumnDelimiter")
    m_ColumnWidths.AsVariant = Split(Lines.Value("ColumnWidths"), ",")
    m_ColumnCaptions.AsCSV = Lines.Value("ColumnCaptions")
    m_FileHasHeaderLine = Lines.Value("HasHeaderLine")
    ''m_IdentHeaderLines.AsCSV = Lines.Value("HeaderLines")
    ''m_IdentFileNamePatterns.AsCSV = Lines.Value("FileNamePatterns")
End Sub

' Returns the state of the object as a serialized string.
'
Public Property Get AsText() As String
    AsText = SaveToStringList.AsCSV
End Property

' Sets the state of the object from a serialized string.
'
Public Property Let AsText(RHS As String)
    Dim Lines As New usStringList
    Lines.AsCSV = RHS
    Call LoadFromStringList(Lines)
End Property

' Loads the state of the object from a file under the user's app-data folder.
'
Public Sub Load(ByVal Name As String)
    Dim FilePath As String
    FilePath = g_AppOptions.LogFormatsPath(Name)
    Dim Lines As New usStringList
    Call Lines.LoadFromFile(FilePath)
    Call LoadFromStringList(Lines)
    m_Name = Name
End Sub

' Saves the state of the object to a file under the user's app-data folder.
'
Public Sub Save()
    Dim FilePath As String
    FilePath = g_AppOptions.LogFormatsPath(m_Name)
    Call SaveToStringList.SaveToFile(FilePath)
End Sub

' Deletes this object's persisted state file under the user's app-data folder
' if it exists.  Then, saves the state to the specified file.
'
Public Sub Rename(ByVal Name As String)
    If Name = m_Name Then Exit Sub
    Dim NewFilePath As String
    NewFilePath = g_AppOptions.LogFormatsPath(Name)
    If FSO.FileExists(NewFilePath) Then
        Err.Raise vbObjectError, Description:="A log format named '" & Name & "' already exists."
    End If
    Dim OldFilePath As String
    OldFilePath = g_AppOptions.LogFormatsPath(m_Name)
    If FSO.FileExists(OldFilePath) Then
        Call FSO.DeleteFile(OldFilePath)
    End If
    Call SaveToStringList.SaveToFile(NewFilePath)
    m_Name = Name
End Sub

' Deletes this object's persisted state file under the user's app-data folder
' if it exists.  Then, clears the object state since the caller is expected
' to delete or recycle this instance after calling this.
'
Public Sub Delete()
    Dim FilePath As String
    FilePath = g_AppOptions.LogFormatsPath(m_Name)
    If FSO.FileExists(FilePath) Then
        Call FSO.DeleteFile(FilePath)
    End If
    Call Clear
End Sub

