VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogViewUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private Type LOG_FILE_INFO
    f As Scripting.File
    ModDate As Date
    Size As Long
End Type

Public Enum ERROR_NUMBERS
    E_FILE_NOT_FOUND = vbObjectError + 4100
    E_EMPTY_LOG = vbObjectError + 4101
End Enum

Public Event FilePathChanged()
Public Event SelectedItemChanged()
Public Event ItemCountChanged(ByVal Count As Long)
Public Event FormatChanged()

Private WithEvents m_GridView As MSComctlLib.ListView
Attribute m_GridView.VB_VarHelpID = -1
Private WithEvents m_StreamView As RichTextLib.RichTextBox
Attribute m_StreamView.VB_VarHelpID = -1

Private m_LastSelectedItem As MSComctlLib.ListItem
Private m_FilePath As String
Private m_CanUpdate As Boolean
Private m_LnDelim As String
Private m_LnDelimLen As Long
Private m_LastLogFileModDate As Date
Private m_LastLogFileSize As Long
Private m_NextToLastLogFilePos As Long
Private m_LastLogFilePos As Long
Private m_LastLogEntryLine As String
Private m_ItemReadCount As Long
Private m_LogFormat As New LogFormat
Private Const MAX_LINE_LEN As Long = 4096

' Initializes object instance.
'
Private Sub Class_Initialize()
    m_LnDelim = vbCrLf
    m_LnDelimLen = Len(vbCrLf)
End Sub

' Searches for a line terminator (CR-LF, CR, or LF) in a string.  If found,
' this sets the line-terminator state and returns the position of the found
' terminator.  Returns 0 if not found.
'
' Text is readonly and passed byref for performance.
'
Private Function DetermineLineTerminator(Text As String) As Long

    Dim SearchText As Variant
    For Each SearchText In Array(vbCrLf, vbCr, vbLf)
        Dim Pos As Long
        Pos = InStr(Text, SearchText)
        If Pos > 0 Then
            m_LnDelim = SearchText
            m_LnDelimLen = Len(SearchText)
            DetermineLineTerminator = Pos
            Exit Function
        End If
    Next
    
End Function

' Returns the first log-format that matches by header-line.
'
' HeadLine and FilePath are readonly and passed byref for performance.
'
Private Function FindMatchingLogFormat(HeadLine As String, FileName As String) As LogFormat
    
    Dim Fmt As LogFormat
    For Each Fmt In g_AppOptions.LogFormats
        If Fmt.ColumnLayout <> CL_NONE And Fmt.HasHeaderLine Then
            If HeadLine = Fmt.HeaderText Then
                Set FindMatchingLogFormat = Fmt
                Exit Function
            End If
        End If
    Next

End Function

' Indicates whether the list of lines seem to be from a file with the
' specified line format.  Returns true if each line consists of more than one
' item and the lines each consist of the same number of items when parsed by
' the format.
'
Private Function LooksLikeLineFormat(ByVal Fmt As LogFormat, ByVal Lines As usStringList) As Boolean

    Debug.Assert Lines.Count > 1
    
    Dim LineItems As Variant
    LineItems = Fmt.SplitLine(Lines(1))
    
    Dim ItemCount As Long
    ItemCount = UBound(LineItems) - LBound(LineItems) + 1
    If ItemCount < 2 Then Exit Function
    
    Dim i As Long
    For i = 2 To Lines.Count
        LineItems = Fmt.SplitLine(Lines(i))
        Dim NextItemCount As Long
        NextItemCount = UBound(LineItems) - LBound(LineItems) + 1
        If NextItemCount < 2 Then Exit Function
        If NextItemCount <> ItemCount Then Exit Function
    Next
    LooksLikeLineFormat = True
    
End Function

' Tries to find a format that matches the file's format.  Otherwise, creates
' new format based on the file contents.
'
' FileText and FilePath are readonly and passed byref for performance.
'
' ALGORITHM
' if first line matches first line of known format
'     return matching format
' elseif line count >= 3
'     count CSV items of 1st 3 lines
'         if CSV item count of each line > 1 and equal
'             return CSVFormat
'     for delim in common delimintors (,;:|~.!@#$%^&*-=+\/?)
'         count delimited items of 1st 3 lines
'             if delimited item count of each line > 1 and equal
'                 return delimited format
'     return UnknownFormat
' else
'     return UnknownFormat
'
Private Function BestLogFormat(FileText As String, FilePath As String) As LogFormat

    Set BestLogFormat = FindMatchingLogFormat(g_AppOptions.HeadLine, FilePath)
    If Not BestLogFormat Is Nothing Then
        Exit Function
    End If
    
    ' default to new format
    Set BestLogFormat = New LogFormat
    If InStr(g_AppOptions.HeadLine, vbTab) Then
        ' line contains a tab char -- so it's probably tab-delimited
        BestLogFormat.ColumnLayout = CL_TAB
        Exit Function
    End If
    
    ' use 1st three lines to determine format
    Dim Pos2ndTerm As Long
    Dim Pos3rdTerm As Long
    Pos2ndTerm = InStr(Len(g_AppOptions.HeadLine) + m_LnDelimLen, FileText, m_LnDelim)
    If Pos2ndTerm > 0 Then
        Pos3rdTerm = InStr(Pos2ndTerm + m_LnDelimLen, FileText, m_LnDelim)
        If Pos3rdTerm > 0 Then
        
            Dim FirstLines As New usStringList
            FirstLines.AsVariant = Split(Left$(FileText, Pos3rdTerm), m_LnDelim)
            Debug.Assert FirstLines.Count = 3
            
            BestLogFormat.ColumnLayout = CL_CSV
            If LooksLikeLineFormat(BestLogFormat, FirstLines) Then
                Exit Function
            End If
            
            BestLogFormat.ColumnLayout = CL_STRING
            Dim CommonDelims As Variant
            CommonDelims = Array(",", ";", ":", "|", "~", ".", "!", "@", "#", "$", "%", "^", "&", "*", "-", "=", "+", "\", "/", "?")
            Dim Delim As Variant
            For Each Delim In CommonDelims
                BestLogFormat.ColumnDelimiter = Delim
                If LooksLikeLineFormat(BestLogFormat, FirstLines) Then
                    Exit Function
                End If
            Next
            
        End If
    End If
    
    ' treat as a stream
    BestLogFormat.ColumnLayout = CL_NONE
    BestLogFormat.HasHeaderLine = False

End Function

' Returns if no first line is found (a line terminator in first 4096
' characters).  Otherwise, determines the file format from the first line,
' user preferences and possibly user input.  Loads the first line as a header
' if the file format indicates that the file has a header line.  Then, loads
' the rest of the file lines.
'
' The file size is passed as an optimization.  The caller already has the
' value so it is passed rather than having this procedure query the file
' system for the info again.
'
' The first line must be no more than 4096 characters long.  This procedure
' treats a very long first line the same as not finding a first line.  This
' limitation allows for some ability to avoid reading binary (non-text)
' files.  For example, consider a binary file that is 10MB long and has
' one line terminator in the middle.  The first line would seem to be 5MB
' long.  Loading extremely long lines causes bad behavior throughout this
' application.
'
Private Sub LoadFromFileObject(ByVal LogFile As Scripting.File, ByVal LogFileSize As Long)

    ' test programming conditions
    ' NOTE: there is a small possibility that the file was updated since the
    '       size was cached, but that shouldn't be a problem.
    Debug.Assert LogFile.Size = LogFileSize
    
    ' load file contents into memory
    ' NOTE: can be a time and memory costly operation
    Dim LogStream As Scripting.TextStream
    Set LogStream = OpenLogFile(LogFile)
    If LogStream.AtEndOfStream Then
        ' file is empty
        Call LogStream.Close
        Exit Sub
    End If
    Dim FileText As String
    FileText = LogStream.ReadAll
    Call LogStream.Close
    
    ' determine line terminator
    Dim FirstTermPos As Long
    FirstTermPos = DetermineLineTerminator(FileText)
    If FirstTermPos < 1 Then
        If Len(FileText) < MAX_LINE_LEN Then
            ' file may not contain a full line
            Exit Sub
        End If
        Call RaiseMsg _
         ("Unable to find a line terminator in first " & Len(FileText) & _
          " characters of file.  " & _
          "The file may be binary which is not supported by this application.")
    End If
    If FirstTermPos > MAX_LINE_LEN Then
        Call RaiseMsg _
         ("The character length of the first line " & Len(FileText) & _
          " is greater than maximum supported " & CStr(MAX_LINE_LEN) & _
          ".  The file might be binary which is not supported by this application.")
    End If
    g_AppOptions.HeadLine = Left$(FileText, FirstTermPos - 1)
    
    ' determine file format
    LogFormat = BestLogFormat(FileText, LogFile.Name)
    
    Dim ColCaptions As New usStringList
    If m_LogFormat.HasHeaderLine Then
        
        ' read column captions from first line
        ColCaptions.AsVariant = m_LogFormat.SplitLine(g_AppOptions.HeadLine)
        
        ' remove first line from file-text
        FileText = Mid$(FileText, Len(g_AppOptions.HeadLine) + m_LnDelimLen + 1)
    
        Debug.Assert m_NextToLastLogFilePos = 0
        Debug.Assert m_LastLogFilePos = 0
        m_LastLogFilePos = Len(g_AppOptions.HeadLine) + m_LnDelimLen
        m_LastLogEntryLine = g_AppOptions.HeadLine
        
    Else
        
        ' use default column captions
        Set ColCaptions = m_LogFormat.ColumnCaptions
    
    End If
    
    ' load default column captions
    Call LoadColumnsFromHeaderLine(ColCaptions)
    
    ' process file contents
    Call AddLinesFromText(FileText)
    
End Sub

' Updates the list view with new items in a file object.
'
' The file size is passed as an optimization.  The caller already has the
' value so it is passed rather than having this procedure query the file
' system for the info again.
'
Private Sub UpdateFromFileObject(ByVal LogFile As Scripting.File, ByVal LogFileSize As Long)

    ' test programming conditions
    Debug.Assert LogFile.Size = LogFileSize
    Debug.Assert LogFile.Size >= m_LastLogFileSize
    Debug.Assert m_LastLogFileSize >= m_LastLogFilePos
    Debug.Assert m_LastLogFilePos >= m_NextToLastLogFilePos

    ' open the file for reading at beginning of last line read
    Dim LogStream As Scripting.TextStream
    Set LogStream = OpenLogFile(LogFile)
    Call LogStream.Skip(m_NextToLastLogFilePos)
        
    Dim FileText As String
    
    If Not LogStream.AtEndOfStream Then

        Dim LogLine As String
        LogLine = LogStream.ReadLine
        
        If LogLine <> m_LastLogEntryLine Then
        
            '
            ' last read line does not match file contents
            '
        
            ' close the file
            Call LogStream.Close
            
            ' raise error
            Dim ErrMsg As String
            ErrMsg = "The log file content does not match the log entries that are currently displayed."
            Call RaiseMsg(ErrMsg)
        
        End If

        ' read to end-of-file
        If Not LogStream.AtEndOfStream Then
            FileText = LogStream.Read(LogFileSize)
        End If

    End If
    
    ' close the file
    Call LogStream.Close
    
    ' process new file contents
    If FileText <> "" Then
        Call AddLinesFromText(FileText)
    End If
    
End Sub

' Opens a file for read-access and returns a text-stream interface for
' reading the file's contents.  The FileSystemObject determines whether the
' file is ASCII or UNICODE -- probably based on the first few characters of
' the file.
'
Private Function OpenLogFile(ByVal LogFile As Scripting.File) As Scripting.TextStream
    
    Dim TestFmt As Scripting.Tristate
    TestFmt = Scripting.TristateUseDefault
    Set OpenLogFile = LogFile.OpenAsTextStream(Scripting.ForReading, TestFmt)

End Function

' Returns log file info for a file or raises an error with Number of
' E_FILE_NOT_FOUND.
'
' NOTE: This is fast since it does not open the file.
'
Private Function ReadLogFileInfo(ByVal FilePath As String) As LOG_FILE_INFO

    On Error GoTo ERROR_HANDLER
    
    Dim Info As LOG_FILE_INFO
    Set Info.f = FSO.GetFile(FilePath)
    Info.ModDate = Info.f.DateLastModified
    Info.Size = Info.f.Size
    ReadLogFileInfo = Info
    
    Exit Function

ERROR_HANDLER:
    
    Call RaiseMsg(Err.Description, E_FILE_NOT_FOUND)

End Function

' Adds a view line for each line of the given text.
'
' Text is readonly and passed by-ref for performance.
'
Private Sub AddLinesFromText(ByRef Text As String)
    
    Dim Line As String
    Dim LineLen As Long
    Dim LineFirstPos As Long
    Dim LineLastPos As Long
    LineFirstPos = 1

    Do
    
        ' find next line-terminator
        LineLastPos = InStr(LineFirstPos, Text, m_LnDelim)
        If LineLastPos = 0 Then Exit Do
        
        ' extract next line
        LineLen = LineLastPos + 1 - LineFirstPos - 1
        Line = Mid$(Text, LineFirstPos, LineLen)

        ' process line
        Call AddLine(Line)
        
        ' move text pointer to beginning of next line
        LineFirstPos = LineLastPos + m_LnDelimLen
        
        ' increment sync state
        m_NextToLastLogFilePos = m_LastLogFilePos
        m_LastLogFilePos = m_LastLogFilePos + LineLen + m_LnDelimLen
        m_LastLogEntryLine = Line
        
    Loop

End Sub

' Adds a view line.
'
' Text is readonly and passed by-ref for performance.
'
Private Sub AddLine(ByRef Text As String)
    
    ' increment item-read count
    ' NOTE: do this before any early exits
    m_ItemReadCount = m_ItemReadCount + 1
    
    If m_GridView.Visible Then
        Debug.Assert Not m_StreamView.Visible
        Call AddEntryFromLine(Text)
    Else
        Debug.Assert m_StreamView.Visible
        m_StreamView.SelText = Text & vbCrLf
    End If

End Sub

' Adds a list-view item for a line-string.
'
' Text is readonly and passed by-ref for performance.
'
Private Sub AddEntryFromLine(ByRef Text As String)

    ' split line into items
    Dim ItemValueList() As String
    Call m_LogFormat.SplitLineToArray(Text, ItemValueList)
    Debug.Assert LBound(ItemValueList) = 0
    
    ' format sequence number value
    Dim SeqText As String
    SeqText = m_ItemReadCount

    ' create new list-view item
    ' NOTE: will be new top item unless the list is sorted
    Dim ListItem As MSComctlLib.ListItem
    Set ListItem = m_GridView.ListItems.Add(1, Text:=SeqText)
    
    ' load subsequent columns
    Dim i As Long
    For i = 0 To UBound(ItemValueList)
    
        ' add new column if more items than in log view
        If i >= m_GridView.ColumnHeaders.Count - 1 Then
            Dim HeadText As String
            HeadText = "Column " & CStr(m_GridView.ColumnHeaders.Count)
            Call m_GridView.ColumnHeaders.Add(Text:=HeadText)
        End If
        
        ListItem.SubItems(i + 1) = ItemValueList(i)
        
    Next i
    
End Sub

' Defines the list-view columns based on a list of column names.
'
Private Sub LoadColumnsFromHeaderLine(ByVal ColumnCaptions As usStringList)
    
    ' add sequence number column
    Dim NewColHeader As ColumnHeader
    Set NewColHeader = m_GridView.ColumnHeaders.Add(Text:="#")
    
    If NewColHeader.Index <= m_LogFormat.ColumnWidths.Count Then
        NewColHeader.Width = m_LogFormat.ColumnWidths.Item(NewColHeader.Index)
    End If

    ' load columns from array of items
    Dim i As Long
    For i = 1 To ColumnCaptions.Count
    
        Dim ColCap As String
        ColCap = ColumnCaptions(i)
        
        Set NewColHeader = m_GridView.ColumnHeaders.Add(Text:=ColCap)
        
        If NewColHeader.Index <= m_LogFormat.ColumnWidths.Count Then
            NewColHeader.Width = m_LogFormat.ColumnWidths.Item(NewColHeader.Index)
        End If
    
    Next i

End Sub

' Sorts the list-view in alphabetical order of one of the columns.
'
Private Sub ColumSort(ByVal ColumnIndex As Long)

    If m_GridView.SortKey = ColumnIndex - 1 Then
        m_GridView.SortOrder = 1 - m_GridView.SortOrder
    Else
        m_GridView.SortKey = ColumnIndex - 1
        m_GridView.SortOrder = lvwAscending
    End If

End Sub

' Stores the width of each column into the active log format.
'
Private Sub CopyColumnWidthsToLogFormat()

    ' limit the column count
    Dim ColCount As Long
    ColCount = m_GridView.ColumnHeaders.Count
    If ColCount > 100 Then ColCount = 100

    ' make sure there are at least as many option-values as columns
    If ColCount > m_LogFormat.ColumnWidths.Count Then
        m_LogFormat.ColumnWidths.Count = ColCount
    End If

    ' load column widths into option-values
    Dim i As Long
    For i = 1 To ColCount
        m_LogFormat.ColumnWidths.Item(i) = m_GridView.ColumnHeaders(i).Width
    Next i

End Sub

'CSEH: DebugAssert
Private Sub m_GridView_KeyUp(KeyCode As Integer, Shift As Integer)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call CheckForSelectedItemChange
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

'CSEH: DebugAssert
Private Sub m_GridView_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call CheckForSelectedItemChange
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

' Fires SelectedItemChanged if selected-item of log view has changed since
' this was last called.  This needs to be called when any type of event
' occurs that might have caused a change to which item is selected.
'
Private Sub CheckForSelectedItemChange()

    If Not m_LastSelectedItem Is m_GridView.SelectedItem Then
        
        Set m_LastSelectedItem = m_GridView.SelectedItem
    
        RaiseEvent SelectedItemChanged
        
    End If

End Sub

' Attaches the log view controls to this instance.
'
Public Sub Attach _
 (ByVal GridView As MSComctlLib.ListView, _
  ByVal StreamView As RichTextLib.RichTextBox)
    Set m_GridView = GridView
    Set m_StreamView = StreamView
End Sub

' Returns the active log file path.
'
Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

' Sets the active log file path.
'
Private Sub SetFilePath(ByVal RHS As String)
    If RHS <> m_FilePath Then
        m_FilePath = RHS
        RaiseEvent FilePathChanged
    End If
End Sub

' Indicates whether update is allowed.
'
Public Property Get CanUpdate() As Boolean
    CanUpdate = m_CanUpdate
End Property

' Returns a reference to the log format.
'
Public Property Get LogFormat() As LogFormat
    Set LogFormat = m_LogFormat
End Property

' Sets the log format reference.
'
Public Property Let LogFormat(ByVal RHS As LogFormat)
    If Not RHS Is m_LogFormat Then
        Call SaveFormatPrefs
        Set m_LogFormat = RHS
        RaiseEvent FormatChanged
    End If
End Property

' Updates and saves the active log format if it is a known format.
'
Public Sub SaveFormatPrefs()
    If g_AppOptions.LogFormats.Find(m_LogFormat) > 0 Then
        Call CopyColumnWidthsToLogFormat
        Call m_LogFormat.Save
    End If
End Sub

' Clears the log view and the cached state about the last read position in
' the active log file.
'
Public Sub Clear()
    
    m_CanUpdate = False
    m_LastLogFileModDate = 0
    m_LastLogFileSize = -1
    m_LastLogFilePos = 0
    m_NextToLastLogFilePos = 0
    m_LastLogEntryLine = ""
    m_ItemReadCount = 0
    
    If m_GridView.Visible Then
    
        ' clear any list items
        If m_GridView.ListItems.Count > 0 Then
            Call m_GridView.ListItems.Clear
            Call CheckForSelectedItemChange
            RaiseEvent ItemCountChanged(m_ItemReadCount)
        End If
        
        ' store current column widths
        ' NOTE: do this before clearing the columns
        Call SaveFormatPrefs
        
        ' clear the columns
        Call m_GridView.ColumnHeaders.Clear
    
    End If
        
    ' clear stream view
    m_StreamView.Text = ""

End Sub

' Clears the display and then loads it from the contents of a file.
' Propagates an error if the file does not exit, does not contain at least
' one line or its format cannot be determined.
'
Public Sub Load(ByVal FilePath As String)
    
    ' clear view and last known position in file
    Call Clear

    ' set file path
    Call SetFilePath(FilePath)
    
    ' read file info
    Dim LogInfo As LOG_FILE_INFO
    LogInfo = ReadLogFileInfo(FilePath)
    
    ' set mouse pointer to hourglass for rest of procedure
    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbHourglass)
    
    ' error if file is empty
    ' NOTE: a unicode file is empty if it's length is less than 4 bytes
    ' NOTE: the first 2 bytes of a unicode file mark the file as unicode and not part of the file contents
    If LogInfo.Size < 4 Then
        Call RaiseMsg("File is empty.", E_EMPTY_LOG)
    End If
        
    ' load file
    On Error GoTo LOAD_FAILURE
    Call LoadFromFileObject(LogInfo.f, LogInfo.Size)
    On Error GoTo 0
    
    ' store current file statistics for subsequent update
    ' NOTE: this is done after the contents are processed
    m_LastLogFileModDate = LogInfo.ModDate
    m_LastLogFileSize = LogInfo.Size
    
    ' select top item in list
    If m_GridView.ListItems.Count > 0 Then
        If Not m_GridView.SelectedItem Is Nothing Then
            m_GridView.SelectedItem.Selected = False
        End If
        m_GridView.SelectedItem = m_GridView.ListItems(1)
    End If
    
    ' invoke change events
    Call CheckForSelectedItemChange
    RaiseEvent ItemCountChanged(m_ItemReadCount)
    
    m_CanUpdate = True
    
    Exit Sub
    
LOAD_FAILURE:
    Dim ErrCache As ERR_CACHE
    Call CacheErr(ErrCache)
    Call SetFilePath("")
    Call Clear
    Call RaiseCachedErr(ErrCache)

End Sub

' Updates the log view from the active log file.
'
' Call this periodically to keep the view in sync with the file contents.
'
' IMPLEMENTATION
' This minimizes resource use and execution time in the very likely event that
' the file has not changed since the last sync.  First, the file is tested for
' changes.  If no changes are detected since the last sync, then control is
' returned to the caller without incuring the relatively large overhead of
' opening the file.  If changes are detected, then the file is opened and then
' closed as soon as possible.
'
' Note that while the file is open, another process will not be able to delete
' it.  While it is not open, it can be deleted.  For example, even if
' FSO.FileExists indicates that the file does exist, the file could be deleted
' by another process before a subseqent access to the file -- i.e. FSO.GetFile.
'
Public Sub Update()

    Debug.Assert m_CanUpdate

    ' read file info
    Dim LogInfo As LOG_FILE_INFO
    LogInfo = ReadLogFileInfo(m_FilePath)
    
    ' exit if last mod-date and size are same as last update
    If LogInfo.ModDate = m_LastLogFileModDate And LogInfo.Size = m_LastLogFileSize Then
        Exit Sub
    End If
    
    m_CanUpdate = False
    
    If m_LastLogFileSize > LogInfo.Size Then
        Dim ErrMsg As String
        ErrMsg = "The log file is smaller than it was during the last update."
        Call RaiseMsg(ErrMsg)
    End If

    If Not m_GridView.Visible Then
        Dim RTSS As New RichTextSaveSel
        Call RTSS.Save(m_StreamView)
        m_StreamView.SelStart = Len(m_StreamView.Text)
        m_StreamView.SelHangingIndent = 150
    End If

    ' set mouse pointer to hourglass for rest of procedure
    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbHourglass)
    
    ' update from the file
    Call UpdateFromFileObject(LogInfo.f, LogInfo.Size)
    
    ' store current file statistics for next synch
    ' NOTE: this is done after the contents are processed
    m_LastLogFileModDate = LogInfo.ModDate
    m_LastLogFileSize = LogInfo.Size
    
    ' invoke change events
    Call CheckForSelectedItemChange
    RaiseEvent ItemCountChanged(m_ItemReadCount)
    
    m_CanUpdate = True

End Sub
