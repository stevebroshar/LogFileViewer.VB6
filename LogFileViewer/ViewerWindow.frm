VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ViewerWindow 
   Caption         =   "[no log file loaded]"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ViewerWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NoItemsLabel 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "No log entries"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox StreamLogView 
      Height          =   1335
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1e6
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"ViewerWindow.frx":030A
   End
   Begin RichTextLib.RichTextBox DetailEdit 
      Height          =   5895
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   10398
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"ViewerWindow.frx":038B
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   480
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":041B
            Key             =   "Information"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":086D
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":0CBF
            Key             =   "Question"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":1111
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":1563
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewerWindow.frx":19B5
            Key             =   "SortUp"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   2400
   End
   Begin VB.Frame Splitter 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   4440
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Key             =   "Latency"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "Live"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "Items"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9225
            MinWidth        =   2293
            Key             =   "File"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView GridLogView 
      Height          =   1305
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2302
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Menu FileMenuItem 
      Caption         =   "&File"
      Begin VB.Menu OpenFileMenuItem 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu OpenRecentMenuItem 
         Caption         =   "Open Recen&t"
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu RecentFileMenuItem 
            Caption         =   ""
            Index           =   8
         End
      End
      Begin VB.Menu SaveAsMenuItem 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu RenameFileMenuItem 
         Caption         =   "&Rename File..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu CopyFileMenuItem 
         Caption         =   "&Copy File..."
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu DeleteFileMenuItem 
         Caption         =   "&Delete File"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu SepMenuItem6 
         Caption         =   "-"
      End
      Begin VB.Menu ExportMenuItem 
         Caption         =   "&Export..."
         Shortcut        =   ^E
      End
      Begin VB.Menu SepMenuItem1 
         Caption         =   "-"
      End
      Begin VB.Menu MergeMultipleFilesMenuItem 
         Caption         =   "&Merge..."
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu SepMenuItem10 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMenuItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenuItem 
      Caption         =   "&Edit"
      Begin VB.Menu CutMenuItem 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu CopyMenuItem 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu PasteMenuItem 
         Caption         =   "&Paste"
         Shortcut        =   ^V
         Visible         =   0   'False
      End
      Begin VB.Menu SelectAllMenuItem 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu SepMenuItem7 
         Caption         =   "-"
      End
      Begin VB.Menu FindMenuItem 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNextMenuItem 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu SepMenuItem5 
         Caption         =   "-"
      End
      Begin VB.Menu OptionsMenuItem 
         Caption         =   "&Options..."
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu ViewMenuItem 
      Caption         =   "&View"
      Begin VB.Menu LiveMenuItem 
         Caption         =   "&Live"
         Shortcut        =   {F9}
      End
      Begin VB.Menu GridViewMenuItem 
         Caption         =   "&Grid View"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu DetailMenuItem 
         Caption         =   "&Detail"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu GridLinesMenuItem 
         Caption         =   "Grid &Lines"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu SepMenuItem3 
         Caption         =   "-"
      End
      Begin VB.Menu RefreshMenuItem 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu UpdateMenuItem 
         Caption         =   "&Update"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu FormatMenuItem 
      Caption         =   "&Format"
      Begin VB.Menu FormatsMenuItem 
         Caption         =   "&Manage Formats..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu SaveFormatMenuItem 
         Caption         =   "&Save This Format..."
      End
      Begin VB.Menu EditFormatMenuItem 
         Caption         =   "&Edit This Format..."
      End
   End
   Begin VB.Menu HelpMenuItem 
      Caption         =   "&Help"
      Begin VB.Menu AboutMenuItem 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "ViewerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Host container for Pages"
'CSEH: Skip
Option Explicit

Private Type FIND_OPTIONS
    Text As String
    MatchCase As Boolean
    WholeWord As Boolean
    Down As Boolean
End Type

Private m_FindOptions As FIND_OPTIONS
Private m_DefLogFileName As String
Private m_OutboundDragFileName As String
Private WithEvents m_LogViewUpdater As LogViewUpdater
Attribute m_LogViewUpdater.VB_VarHelpID = -1
Private WithEvents m_FindEvents As FindForm
Attribute m_FindEvents.VB_VarHelpID = -1
Private m_TimerIntervalController As TimerIntervalController
Private m_MaxFileSizeBytes As Long

' Returns a formatted date/time stamp string.
'
Private Function FormattedDateTimeStamp() As String
    FormattedDateTimeStamp = Format$(Now, "YYYY-MM-DD HH-NN-SS")
End Function

' Returns a formatted timestamp string.
'
Private Function FormattedTimeStamp() As String
    FormattedTimeStamp = Format$(Now, "Long Time")
End Function

' Returns a reference to a status bar panel.
'
Private Property Get LatencyStatusPanel() As MSComctlLib.Panel
    Set LatencyStatusPanel = StatusBar.Panels("Latency")
End Property
Private Property Get LiveStatusPanel() As MSComctlLib.Panel
    Set LiveStatusPanel = StatusBar.Panels("Live")
End Property
Private Property Get ItemsStatusPanel() As MSComctlLib.Panel
    Set ItemsStatusPanel = StatusBar.Panels("Items")
End Property
Private Property Get FileStatusPanel() As MSComctlLib.Panel
    Set FileStatusPanel = StatusBar.Panels("File")
End Property

' Loads the window caption based on the active log file path and format.
'
Private Sub LoadWindowCaption()

    Dim CapText As String
    
    If m_LogViewUpdater Is Nothing Then
        CapText = "[no file]"
    Else
        If Len(m_LogViewUpdater.FilePath) > 0 Then
            CapText = FSO.GetFileName(m_LogViewUpdater.FilePath)
        Else
            CapText = "[no file]"
        End If
        Dim FmtName As String
        FmtName = m_LogViewUpdater.LogFormat.Name
        If Len(FmtName) > 0 Then
            CapText = CapText & " - " & FmtName
        End If
    End If
    
    CapText = CapText & " - " & App.Title
    
    Me.Caption = CapText

End Sub

' Based on user selection, freezes or clears the display.
'
Private Sub ClearOrFreeze(ByVal ReasonMsg As String)

    Dim Msg As String
    Msg = _
      ReasonMsg & vbCrLf & vbCrLf & _
      "Do you want to clear the display?" & vbCrLf & vbCrLf & _
      "Select Yes to clear the display and load the new contents of the log file or No to freeze the display."
      
    If MsgBox(Msg, vbQuestion + vbYesNo) = vbNo Then
        LiveUpdate = False
    Else
        Call m_LogViewUpdater.Clear
    End If
    
End Sub

' Updates the view from the current contents of the active file.
'
Private Sub UpdateView()
        
    If m_LogViewUpdater.CanUpdate Then
        
        Dim DurationTimer As usSimpleTimer
        Set DurationTimer = New usSimpleTimer
        
        On Error GoTo UPDATE_FAILURE
        Call m_LogViewUpdater.Update
        On Error GoTo 0
        
        Call m_TimerIntervalController.Adjust(DurationTimer.ElapsedTime)
        
        LatencyStatusPanel = "Updated " & FormattedTimeStamp
    
    Else
        
        Call LoadViewFromFile(m_LogViewUpdater.FilePath)
    
    End If
    
    
    Exit Sub
    
UPDATE_FAILURE:

    If Err.Number = E_FILE_NOT_FOUND Then
        Call ClearOrFreeze("The log file was renamed, moved or deleted.")
    Else
        Call ClearOrFreeze(Err.Description)
    End If
    
End Sub

' Returns the list view item that has a sequence number that is the closest to
' the specified ID without going over its value.
'
Private Function FindItemWithClosestSeqNum(ByVal ID As Long) As MSComctlLib.ListItem

    Dim BestItemID As Long
    Dim BestItem As MSComctlLib.ListItem
    Dim ItemID As Long
    Dim Item As MSComctlLib.ListItem
    For Each Item In GridLogView.ListItems
    
        ItemID = 0
        On Error Resume Next
        ItemID = Item.Text
        On Error GoTo 0
        If ItemID > 0 Then
            If ItemID = ID Then
                Set FindItemWithClosestSeqNum = Item
                Exit Function
            End If
            If ItemID < ID And ItemID > BestItemID Then
                BestItemID = ItemID
                Set BestItem = Item
            End If
        End If
        
    Next Item
    
    Set FindItemWithClosestSeqNum = BestItem

End Function

' Starts monitoring a file.
'
Private Sub AddMruFilePath(ByVal FilePath As String)

    Dim ExistingPos As Long
    ExistingPos = g_AppOptions.MruFilePaths.FindIgnoringCase(FilePath)
    If ExistingPos > 0 Then
        Call g_AppOptions.MruFilePaths.Remove(ExistingPos)
    End If
    Call g_AppOptions.MruFilePaths.Insert(1, FilePath)
    If g_AppOptions.MruFilePaths.Count > 9 Then
        g_AppOptions.MruFilePaths.Count = 9
    End If
    
End Sub

' Clears the view and then loads it from a log file.
'
Private Sub LoadViewFromFile(ByVal FilePath As String)
    
    Dim DurationTimer As usSimpleTimer
    Set DurationTimer = New usSimpleTimer
    
    On Error GoTo LOAD_ERROR
    Call m_LogViewUpdater.Load(FilePath)
    On Error GoTo 0

    Call m_TimerIntervalController.Clear

    LatencyStatusPanel = "Loaded " & FormattedTimeStamp
    
    Exit Sub
    
LOAD_ERROR:

    If Err.Number = E_FILE_NOT_FOUND Then
        
        Call m_TimerIntervalController.Adjust(DurationTimer.ElapsedTime)
        LatencyStatusPanel = "No File " & FormattedTimeStamp
    
    ElseIf Err.Number = E_EMPTY_LOG Then
        
        Call m_TimerIntervalController.Adjust(DurationTimer.ElapsedTime)
        LatencyStatusPanel = "Empty " & FormattedTimeStamp
    
    Else
        
        Call RaiseErr
    
    End If

End Sub

' Returns the size of a file or zero if the file does not exist.
'
Private Function ReadSizeOfFile(ByVal FilePath As String) As Long

    On Error GoTo ERROR_HANDLER
    Dim f As Scripting.File
    Set f = FSO.GetFile(FilePath)
    ReadSizeOfFile = f.Size
    Exit Function
    
ERROR_HANDLER:
    ReadSizeOfFile = 0
    
End Function

' Handles the condition that the log file is larger than the maximum size.
' Returns False if the user cancels.
'
Private Function ConfirmLoadLargeFile(ByVal FilePath As String, ByVal FileSize As Long) As Boolean
    
    Debug.Assert FileSize > m_MaxFileSizeBytes
    
RETRY:
    Dim LFResult As LARGE_FILE_DIALOG_RESULT
    LFResult = LargeFileDialog.Execute(FilePath, FileSize, m_MaxFileSizeBytes)
    If LFResult = LFDR_LOAD Then
        m_MaxFileSizeBytes = FileSize * 1.1
    ElseIf LFResult = LFDR_CANCEL Then
        Exit Function
    ElseIf LFResult = LFDR_RENAME Then
        If Not RenameLogFile(FilePath) Then GoTo RETRY
    Else
        Debug.Assert False
    End If
    ConfirmLoadLargeFile = True

End Function

' Starts monitoring a file and returns True.  Returns False if the user
' cancels.
'
Private Function OpenFile(ByVal FilePath As String) As Boolean

    ' fully expand file path and dereference if shell link
    FilePath = FSO.GetAbsolutePathName(FilePath)
    If LCase$(FSO.GetExtensionName(FilePath)) = "lnk" Then
        Dim Shell As New WshShell
        Dim Shortcut As WshShortcut
        Set Shortcut = Shell.CreateShortcut(FilePath)
        FilePath = Shortcut.TargetPath
    End If
    
    ' check for and handle very large file
    ' TODO: support skipping lines instead of increasing the max size
    If g_AppOptions.MaxFileSizeBytes > m_MaxFileSizeBytes Then
        m_MaxFileSizeBytes = g_AppOptions.MaxFileSizeBytes
    End If
    Dim LogFileSize As Long
    LogFileSize = ReadSizeOfFile(FilePath)
    If LogFileSize > m_MaxFileSizeBytes Then
        If Not ConfirmLoadLargeFile(FilePath, LogFileSize) Then
            Exit Function
        End If
    End If
    
    ' load file
    Call LoadViewFromFile(FilePath)
    
    ' add path to MRU list
    Call AddMruFilePath(FilePath)
    
    OpenFile = True
    
End Function

' Clears the view and then loads it from the active log file.
'
Private Sub RefreshView()

    ' store sequence number of selected item
    Dim SelID As Long
    On Error Resume Next
    SelID = GridLogView.SelectedItem.Text
    On Error GoTo 0

    Call LoadViewFromFile(m_LogViewUpdater.FilePath)

    ' restore item selection
    If SelID > 0 Then
        Dim SelItem As MSComctlLib.ListItem
        Set SelItem = FindItemWithClosestSeqNum(SelID)
        If Not SelItem Is Nothing Then
            If Not GridLogView.SelectedItem Is Nothing Then
                GridLogView.SelectedItem.Selected = False
            End If
            GridLogView.SelectedItem = SelItem
            SelItem.EnsureVisible
        End If
    End If
    
End Sub

' Updates the item-detail dialog with the data of the selected log view item.
'
Private Sub UpdateItemDetailView()
    
    If Not DetailEdit.Visible Then Exit Sub
    
    If GridLogView.SelectedItem Is Nothing Or GridLogView.ColumnHeaders.Count < 1 Then
    
        DetailEdit.Font.Italic = True
        DetailEdit.Text = "No log entry is selected"
    
    Else
        
        DetailEdit.Font.Italic = False
        DetailEdit.Text = ""
        
        DetailEdit.SelBold = True
        DetailEdit.SelText = "Item #" & GridLogView.SelectedItem.Text
        DetailEdit.SelBold = False
        
        Dim i As Long
        For i = 2 To GridLogView.ColumnHeaders.Count
            DetailEdit.SelText = vbCrLf
            DetailEdit.SelBold = True
            DetailEdit.SelText = GridLogView.ColumnHeaders(i).Text & ": "
            DetailEdit.SelBold = False
            DetailEdit.SelText = GridLogView.SelectedItem.SubItems(i - 1)
        Next i
    
    End If

End Sub

' Returns a string that is based on the source string, but without any
' characters that would cause it to be an illegal XML tag name.
'
Private Function MakeValidTagName(ByVal TagName As String) As String
        
    Dim Result As String
    Dim i As Long
    For i = 1 To Len(TagName)
        Dim Ch As String
        Ch = Mid$(TagName, i, 1)
        If Ch Like "[a-zA-Z]" Then
            Result = Result & Ch
        End If
    Next i
    If Len(Result) = 0 Then Result = "NoName"
    MakeValidTagName = Result

End Function

' Adds a new sub-element to an element.
'
Private Function AddXmlElement(ByVal ParentNode As Object, ByVal Name As String, Optional ByVal Value As String) As Object

    Dim ChildElement As Object
    Set ChildElement = ParentNode.ownerDocument.createElement(Name)
    If Value <> vbNullString Then
        ChildElement.Text = Value
    End If
    Call ParentNode.appendChild(ChildElement)
    Set AddXmlElement = ChildElement
    
End Function

' Returns the selected log entries as a string formatted as a sequence of
' tab-separated lines.
'
Private Function FormatSelectedLogItemsAsText() As String
    
    Dim LastColumn As Long
    LastColumn = GridLogView.ColumnHeaders.Count
    If LastColumn < 2 Then Exit Function
    
    Dim TextItems() As String
    ReDim TextItems(2 To LastColumn)
    
    Dim Result As String
    Dim ListItem As MSComctlLib.ListItem
    For Each ListItem In GridLogView.ListItems
        If ListItem.Selected Then
            Dim i As Long
            For i = 2 To LastColumn
                TextItems(i) = ListItem.SubItems(i - 1)
            Next i
            Result = Result & Join(TextItems, vbTab) & vbCrLf
        End If
    Next ListItem
    
    FormatSelectedLogItemsAsText = Result

End Function

' Writes the data of each selected log view item to a tab-delimited file.
'
Private Sub SaveSelectedLogItemsToTabFile(ByVal FileName As String)

    Dim File As Scripting.TextStream
    Set File = FSO.CreateTextFile(FileName, Overwrite:=True, Unicode:=True)
    
    Dim LastColumn As Long
    LastColumn = GridLogView.ColumnHeaders.Count
    
    Dim TextItems() As String
    ReDim TextItems(2 To LastColumn)
    
    Dim i As Long
    For i = 2 To LastColumn
        TextItems(i) = GridLogView.ColumnHeaders(i).Text
    Next i
    Call File.WriteLine(Join(TextItems, vbTab))
    
    Dim L As Long
    For L = GridLogView.ListItems.Count To 1 Step -1
        Dim ListItem As MSComctlLib.ListItem
        Set ListItem = GridLogView.ListItems(L)
        If ListItem.Selected Then
            For i = 2 To LastColumn
                TextItems(i) = ListItem.SubItems(i - 1)
            Next i
            Call File.WriteLine(Join(TextItems, vbTab))
        End If
    Next L
    
    Call File.Close
    
End Sub

' Writes the data of each selected log view item to a CSV file.
'
Private Sub SaveSelectedLogItemsToCsvFile(ByVal FileName As String)
    
    Dim File As Scripting.TextStream
    Set File = FSO.CreateTextFile(FileName, Overwrite:=True, Unicode:=True)
    
    Dim LastColumn As Long
    LastColumn = GridLogView.ColumnHeaders.Count
    
    Dim TextItems() As String
    ReDim TextItems(2 To LastColumn)
    
    Dim i As Long
    For i = 2 To LastColumn
        TextItems(i) = GridLogView.ColumnHeaders(i).Text
    Next i
    Dim LineItems As New usStringList
    LineItems.AsVariant = TextItems
    Call File.WriteLine(LineItems.AsCSV)
    
    Dim L As Long
    For L = GridLogView.ListItems.Count To 1 Step -1
        Dim ListItem As MSComctlLib.ListItem
        Set ListItem = GridLogView.ListItems(L)
        If ListItem.Selected Then
            For i = 2 To LastColumn
                TextItems(i) = ListItem.SubItems(i - 1)
            Next i
            LineItems.AsVariant = TextItems
            Call File.WriteLine(LineItems.AsCSV)
        End If
    Next L
    
    Call File.Close

End Sub

' Writes the data of each selected log view item to an XML file.
'
Private Sub SaveSelectedLogItemsToXmlFile(ByVal FileName As String)
    
    Dim Document As Object 'DOMDocument
    Set Document = CreateObject("MSXML.DOMDocument")
    Dim RootElement As Object 'IXMLDOMElement
    Set RootElement = Document.createElement("Errors")
    Call Document.appendChild(RootElement)

    Dim L As Long
    For L = GridLogView.ListItems.Count To 1 Step -1
        Dim ListItem As MSComctlLib.ListItem
        Set ListItem = GridLogView.ListItems(L)
        If ListItem.Selected Then
            Dim ErrorElement As Object 'IXMLDOMElement
            Set ErrorElement = AddXmlElement(RootElement, "Error")
            'Call AddXmlElement(ErrorElement, MakeValidTagName(GridLogView.ColumnHeaders(1).Text), LstItem.Text)
            Dim i As Long
            For i = 2 To GridLogView.ColumnHeaders.Count
                Call AddXmlElement(ErrorElement, MakeValidTagName(GridLogView.ColumnHeaders(i).Text), ListItem.SubItems(i - 1))
            Next i
        End If
    Next L
    
    Call Document.Save(FileName)

End Sub

' Returns a path to use for a backup copy of a file based on the path of the
' existing file.
'
Private Function FormatBackupFilePath(ByVal FilePath As String) As String
    
    Dim FolderName As String
    FolderName = FSO.GetParentFolderName(FilePath)
    Dim FilePrefix As String
    FilePrefix = FSO.GetBaseName(FilePath)
    Dim FileExt As String
    FileExt = FSO.GetExtensionName(FilePath)
    Dim FileStamp As String
    FileStamp = " (backup at " & FormattedDateTimeStamp & ")"
    FormatBackupFilePath = FSO.BuildPath(FolderName, FilePrefix & FileStamp & "." & FileExt)
    
End Function

' Stores the default log file name.
'
' NOTE: Doing this: 1) makes subsequent accesses to the default log file name
' faster and 2) insures that environmental changes (such as COM object
' registration) do not cuase the target file name to change.
'
'CSEH: DebugAssert
Private Sub CacheDefaultLogFileName()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    ' clear the cached value in case failure occurs below
    m_DefLogFileName = ""
    
    On Error GoTo ERROR_HANDLER
    m_DefLogFileName = g_AppOptions.DefaultLogFilename
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

' Gets/sets whether the grid lines are enabled.
'
Private Property Get GridLinesEnabled() As Boolean
    Debug.Assert GridLinesMenuItem.Checked = GridLogView.GridLines
    GridLinesEnabled = GridLinesMenuItem.Checked
End Property
Private Property Let GridLinesEnabled(ByVal RHS As Boolean)
    If GridLinesEnabled = RHS Then Exit Property
    GridLinesMenuItem.Checked = RHS
    GridLogView.GridLines = RHS
End Property

' Gets/sets whether live-update is enabled.
'
Private Property Get LiveUpdate() As Boolean
    LiveUpdate = Timer.Enabled
End Property
Private Property Let LiveUpdate(ByVal Value As Boolean)
    Timer.Enabled = Value
    LiveMenuItem.Checked = Timer.Enabled
    LiveStatusPanel.Enabled = Timer.Enabled
End Property

' Gets/sets whether the view is grid (vs. stream).
'
Private Property Get GridViewEnabled() As Boolean
    Debug.Assert GridViewMenuItem.Checked = GridLogView.Visible
    GridViewEnabled = GridViewMenuItem.Checked
End Property
Private Property Let GridViewEnabled(ByVal RHS As Boolean)
    If GridViewEnabled = RHS Then Exit Property
    GridViewMenuItem.Checked = RHS
    GridLogView.Visible = RHS
    StreamLogView.Visible = Not RHS
    If RHS Then
        DetailEdit.Visible = DetailMenuItem.Checked
        Splitter.Visible = DetailMenuItem.Checked
    Else
        DetailEdit.Visible = False
        Splitter.Visible = False
    End If
End Property

' Gets/sets whether the detail view is visible.
'
Private Property Get DetailVisible() As Boolean
    Debug.Assert DetailEdit.Visible = Splitter.Visible
    DetailVisible = DetailMenuItem.Checked
End Property
Private Property Let DetailVisible(ByVal RHS As Boolean)
    If RHS = DetailVisible Then Exit Property
    DetailMenuItem.Checked = RHS
    DetailEdit.Visible = RHS
    Splitter.Visible = RHS
    If RHS Then
        Call UpdateItemDetailView
    End If
    Call PositionControls
End Property

' Positions the controls relative to each other.
'
Private Sub PositionControls()

    On Error GoTo ERROR_HANDLER

    Dim CenterHeight As Long
    CenterHeight = Me.ScaleHeight - StatusBar.Height

    If DetailVisible Then
    
        Splitter.Height = CenterHeight
        DetailEdit.Height = CenterHeight
    
        DetailEdit.Left = Me.ScaleWidth - DetailEdit.Width
        Splitter.Left = DetailEdit.Left - Splitter.Width
        
        Call GridLogView.Move(0, 0, Splitter.Left, CenterHeight)
        
    Else
    
        Call GridLogView.Move(0, 0, Me.ScaleWidth, CenterHeight)
    
    End If

    Call StreamLogView.Move(0, 0, Me.ScaleWidth, CenterHeight)
    Call NoItemsLabel.Move((GridLogView.Width - NoItemsLabel.Width) / 2, (CenterHeight - NoItemsLabel.Height) / 2)
    
    Exit Sub
    
ERROR_HANDLER:
    ' ignore these errors at run-time
    Debug.Print "ViewerWindow.PositionControls: " & Err.Description
    Resume Next

End Sub

' Handles hovering over a drop target during an inbound drag-and-drop
' operation (OLEDragOver).
'
' Sets the Effect flag (which affects the mouse cursor and drop behavior) to
' allow a drop if the data is a file and the drag operation did not start in
' this application.
'
Private Sub HandleOLEDragOver(Data As MSComctlLib.DataObject, Effect As Long)
    
    Dim DataIsSingleFile As Boolean
    If Data.GetFormat(MSComctlLib.ccCFFiles) Then
        DataIsSingleFile = (Data.Files.Count = 1)
    End If
    
    If DataIsSingleFile And m_OutboundDragFileName = "" Then
        Effect = OLEDropEffectConstants.vbDropEffectCopy
    Else
        Effect = OLEDropEffectConstants.vbDropEffectNone
    End If

End Sub

' Handles the drop of an inbound drag-and-drop operation (OLEDragDrop) --
' when the mouse is released during a drag-and-drop operation after the mouse
' has been hovering over a drop target that is willing to accept a drop.
'
' The OLEDragOver handler insures that a drop is only allowed if the data is
' a single file, so this just needs to open the file.
'
Private Sub HandleOLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long)

    Debug.Assert Data.GetFormat(MSComctlLib.ccCFFiles)
    Debug.Assert Data.Files.Count = 1
    Call OpenFile(Data.Files(1))

End Sub

' Handles the start of a drag of an outbound drag-and-drop operation.
'
' OLEStartDrag is fired when the user starts to drag an item (presses the
' mouse button and moves the mouse while holding the mouse button).
'
Private Sub HandleOLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

    Debug.Print "HandleOLEStartDrag"

    ' turn off timer for duration of drag/drop operation
    Timer.Enabled = False

    ' format the file path for the temporary file
    Dim BaseName As String
    BaseName = _
      FSO.GetBaseName(m_LogViewUpdater.FilePath) & _
      " entries " & FormattedDateTimeStamp & _
      "." & FSO.GetExtensionName(m_LogViewUpdater.FilePath)
    m_OutboundDragFileName = FSO.BuildPath(FSO.GetSpecialFolder(Scripting.TemporaryFolder), BaseName)

    ' define supported output formats
    Call Data.SetData(vFormat:=MSComctlLib.ccCFText)
    Call Data.SetData(vFormat:=MSComctlLib.ccCFFiles)

    ' define supported effects
    AllowedEffects = OLEDropEffectConstants.vbDropEffectCopy

End Sub

' Handles setting the data for an outbound drag-and-drop operation.
'
Private Sub HandleOLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)

    Debug.Print "HandleOLESetData"

    If DataFormat = MSComctlLib.ClipBoardConstants.ccCFText Then
        Call Data.SetData(FormatSelectedLogItemsAsText, MSComctlLib.ccCFText)
    ElseIf DataFormat = MSComctlLib.ClipBoardConstants.ccCFFiles Then
        Call SaveSelectedLogItemsToTabFile(m_OutboundDragFileName)
        Call Data.Files.Add(m_OutboundDragFileName)
    Else
        Debug.Assert False
    End If

End Sub

' Handles completing an outbound drag-and-drop operation.  This is fired
' after the drop target has handled a drop event (when the user releases the
' mouse button).  This event always follows a StartDrag event with or without
' a SetData event in between.
'
Private Sub HandleOLECompleteDrag(Effect As Long)

    Debug.Print "HandleOLECompleteDrag"

    Effect = OLEDropEffectConstants.vbDropEffectCopy

    On Error Resume Next
    Call FSO.DeleteFile(m_OutboundDragFileName)
    On Error GoTo 0

    m_OutboundDragFileName = ""

    ' restore timer enabled
    Timer.Enabled = LiveMenuItem.Checked

End Sub

' Initializes the form.
'
Private Sub Form_Load()

        On Error GoTo ERROR_HANDLER
    
        ' initialize controls
100     LatencyStatusPanel.Width = 2000
110     LiveStatusPanel.Width = 650
120     ItemsStatusPanel.Width = 1100
130     LiveStatusPanel.Text = "Live"
140     Call LoadWindowCaption
    
        ' create and initialize updater
150     Set m_LogViewUpdater = New LogViewUpdater
160     Call m_LogViewUpdater.Attach(GridLogView, StreamLogView)

        ' create and initialize timer interval contoller
170     Set m_TimerIntervalController = New TimerIntervalController
180     Call m_TimerIntervalController.Attach(Timer)
        
        ' load persistant state
190     Call g_AppOptions.Load
200     GridLinesEnabled = g_AppOptions.GridLines
210     DetailVisible = g_AppOptions.DetailVisible
220     DetailEdit.Width = g_AppOptions.DetailWidth
230     DetailEdit.Left = Me.ScaleWidth - DetailEdit.Width
240     Call PositionControls
250     Call CacheDefaultLogFileName
    
        ' clear the display
260     Call m_LogViewUpdater_ItemCountChanged(0)
270     Call UpdateItemDetailView
    
        ' scan command-line arguments for startup parameters
        Dim CmdLinePrms As New usStringList
280     CmdLinePrms.AsCSV = Interaction.Command$
290     If CmdLinePrms.Count > 1 Then
300         Call MsgBox("Too many command line parameters.  This application can only open one log file.", vbCritical)
310         End
        End If
        
        ' determine name of file to open
        Dim StartupFileName As String
320     If CmdLinePrms.Count = 1 Then
330         StartupFileName = CmdLinePrms.Item(1)
        Else
340         StartupFileName = m_DefLogFileName
        End If
    
        ' load the startup file
350     Call OpenFile(StartupFileName)
    
        ' default to live-updating
360     LiveUpdate = True

        Exit Sub
    
ERROR_HANDLER:
        ' don't continue from this error
370     Call MsgBox("ViewerWindow.Form_Load (line " & Erl & "): " & Err.Description & " (number " & Err.Number & ")", vbCritical)
380     End ' close the application

End Sub

' Resizes the controls of the form to fit the size of the form.
'
'CSEH: ResumeWithDebugPrint
Private Sub Form_Resize()
    '<EhHeader>
    On Error GoTo ERROR_HANDLER
    '</EhHeader>

100     Call PositionControls
    
    '<EhFooter>
    ' GENERATED CODE: do not modify without removing EhFooter marks
    Exit Sub
ERROR_HANDLER:
        Debug.Print "ViewerWindow.Form_Resize(" & Erl & "): " & VBA.Err.Description
        Resume Next
    '</EhFooter>
End Sub

' Handles form cleanup.
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ERROR_HANDLER
    
    ' release find form
    Call Unload(FindForm)
    
    ' disable updating
    Timer.Enabled = False
    
    ' unload forms
    ' NOTE: if you don't do this, sometimes app does not fully close
    Call VB.Unload(AppOptionsDialog)
    
    ' store persistant state
    Call m_LogViewUpdater.SaveFormatPrefs
    g_AppOptions.GridLines = GridLinesEnabled
    g_AppOptions.DetailVisible = DetailVisible
    g_AppOptions.DetailWidth = DetailEdit.Width
    Call g_AppOptions.Save
    
    Exit Sub
    
ERROR_HANDLER:
    ' don't allow unload errors to prevent closing
    Debug.Print "ViewerWindow.Form_Unload: " & Err.Description
    Debug.Assert False
    Resume Next

End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Const MIN_WIDTH As Long = 1000
    Dim MaxSplitterLeft As Long
    MaxSplitterLeft = Me.ScaleWidth - MIN_WIDTH
    
    If Button = vbLeftButton Then
        Dim NewSplitterLeft As Long
        NewSplitterLeft = Splitter.Left + x
        If NewSplitterLeft < MIN_WIDTH Then NewSplitterLeft = MIN_WIDTH
        If NewSplitterLeft > MaxSplitterLeft Then NewSplitterLeft = MaxSplitterLeft
        Splitter.Left = NewSplitterLeft
        GridLogView.Width = Splitter.Left
        DetailEdit.Left = Splitter.Left + Splitter.Width
        DetailEdit.Width = Me.ScaleWidth - (GridLogView.Width + Splitter.Width)
    End If
    
End Sub

' Handles a change in the number of log view items.
'
'CSEH: DebugAssert
Private Sub m_LogViewUpdater_ItemCountChanged(ByVal Count As Long)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    If Count = 1 Then
        ItemsStatusPanel = Count & " item"
    Else
        ItemsStatusPanel = Count & " items"
    End If
    
    NoItemsLabel.Visible = (Count = 0)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

' Handles a change in log file path.
'
'CSEH: DebugAssert
Private Sub m_LogViewUpdater_FilePathChanged()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    FileStatusPanel.Text = m_LogViewUpdater.FilePath
    Call LoadWindowCaption
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

' Handles a change in which log view item is selected.
'
'CSEH: DebugAssert
Private Sub m_LogViewUpdater_SelectedItemChanged()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call UpdateItemDetailView

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub m_LogViewUpdater_FormatChanged()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call LoadWindowCaption
    If m_LogViewUpdater.LogFormat.ColumnLayout = CL_NONE Then
        GridViewEnabled = False
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Handles a timer event.
'
'CSEH: Skip
Private Sub Timer_Timer()
    
RETRY:
    On Error GoTo ERROR_HANDLER
    
    Call UpdateView
    
    Exit Sub
ERROR_HANDLER:
    Dim mbResult As Long
    Dim Msg As String
    Msg = "Unable to update the display.  " & Err.Description
    mbResult = MsgBox(Msg, vbCritical + vbAbortRetryIgnore)
    If mbResult = vbAbort Then End ' close the application
    If mbResult = vbRetry Then Resume RETRY

End Sub

'CSEH: ShowErr
Private Sub OptionsMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    If AppOptionsDialog.Execute() Then
        
        Call CacheDefaultLogFileName
        
        g_AppOptions.Save
        
        If AppOptionsDialog.ReloadLogFileEdit.Value Then
            Call OpenFile(m_DefLogFileName)
        End If
        
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub MergeMultipleFilesMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim FileName As String
    FileName = MergeMultipleFilesDialog.Execute
    If FileName = "" Then Exit Sub
    Call OpenFile(FileName)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub AboutMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    AboutDialog.Show vbModal

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub ExitMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    End
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: DebugAssert
Private Sub FileMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim FileIs As Boolean
    FileIs = FSO.FileExists(m_LogViewUpdater.FilePath)
    CopyFileMenuItem.Enabled = FileIs
    RenameFileMenuItem.Enabled = FileIs
    DeleteFileMenuItem.Enabled = FileIs
    SaveAsMenuItem.Enabled = FileIs
    ExportMenuItem.Enabled = GridViewEnabled And (GridLogView.ListItems.Count > 0)
    
    If g_AppOptions.MruFilePaths.Count < 2 Then
        OpenRecentMenuItem.Enabled = False
    Else
        OpenRecentMenuItem.Enabled = True
        Dim i As Long
        For i = 1 To 8
            If i < g_AppOptions.MruFilePaths.Count Then
                RecentFileMenuItem(i).Caption = g_AppOptions.MruFilePaths(i + 1)
                RecentFileMenuItem(i).Visible = True
            Else
                RecentFileMenuItem(i).Visible = False
            End If
        Next
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    Debug.Assert False
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub RefreshMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Call RefreshView

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub UpdateMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Call UpdateView

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'' Shows a tool-tip with the full text of a cell.  This information is useful
'' to the user when the cell text display is truncated due to column width.
''
''CSEH: DebugAssert
'Private Sub GridLogView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' TODO: this almost works ... but not when the log view is horizontally scrolled
'    Dim ItemUnderCursor As ListItem
'    Set ItemUnderCursor = GridLogView.HitTest(x, y)
'    If Not ItemUnderCursor Is Nothing Then
'
'        Dim Pos As Long
'        'Pos = GridLogView.HorizontalScrollPosition
'        Dim ch As ColumnHeader
'        For Each ch In GridLogView.ColumnHeaders
'            If Pos + ch.Width > x Then Exit For
'            Pos = Pos + ch.Width
'        Next ch
'
'        If ch.Index = 1 Then
'            GridLogView.ToolTipText = ItemUnderCursor.Text
'        Else
'            GridLogView.ToolTipText = ItemUnderCursor.ListSubItems(ch.Index - 1)
'        End If
'
'    End If
'
'End Sub

' Sorts the log view based on the contents of one of the columns.
'
Private Sub SortOnColumn(ByVal ColumnIndex As Long)

    If ColumnIndex = 1 Then
        If GridLogView.SortKey = ColumnIndex - 1 Then Exit Sub
        GridLogView.Sorted = False
        Call RefreshView
        GridLogView.SortKey = ColumnIndex - 1
    Else
        GridLogView.Sorted = True
    End If
    
    If GridLogView.SortKey = ColumnIndex - 1 Then
        GridLogView.SortOrder = 1 - GridLogView.SortOrder ' reverses sort order
    Else
        GridLogView.SortKey = ColumnIndex - 1
        GridLogView.SortOrder = lvwAscending
    End If

End Sub

'CSEH: ShowErr
Private Sub GridLogView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Call SortOnColumn(ColumnHeader.Index)
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_DblClick()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    DetailVisible = True

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub OpenFileMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim FileName As String
    FileName = AppOptionsDialog.BrowseToOpenLogFile(FSO.GetParentFolderName(m_DefLogFileName), FSO.GetFileName(m_DefLogFileName), "Open Log File")
    If FileName = "" Then Exit Sub
    Call OpenFile(FileName)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub RecentFileMenuItem_Click(Index As Integer)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call OpenFile(RecentFileMenuItem(Index).Caption)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub SaveAsMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim FileName As String
    FileName = AppOptionsDialog.BrowseToSaveLogFile(FSO.GetParentFolderName(m_LogViewUpdater.FilePath), Title:="Save Log File As")
    If FileName = "" Then Exit Sub
    
    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbHourglass)
    Call FSO.CopyFile(m_LogViewUpdater.FilePath, FileName)
    Call OpenFile(FileName)
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub ExportMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim FileName As String
    FileName = AppOptionsDialog.BrowseToSaveExportFile(Title:="Save Selected Log Entries To File")
    If FileName = "" Then Exit Sub
    Dim FileExt As String
    FileExt = LCase$(FSO.GetExtensionName(FileName))
    
    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbHourglass)
    
    If FileExt = "log" Then
        Call SaveSelectedLogItemsToTabFile(FileName)
    ElseIf FileExt = "csv" Then
        Call SaveSelectedLogItemsToCsvFile(FileName)
    ElseIf FileExt = "xml" Then
        Call SaveSelectedLogItemsToXmlFile(FileName)
    Else
        RaiseMsg "INTERNAL ERROR: Unknown file extension '" & FileExt & "'"
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Prompts the user to select a file name and then renames the active log
' file to this name.  Returns False if the user cancels or selects the same
' path as the existing log file.
'
Private Function RenameLogFile(ByVal FilePath As String) As Boolean

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim ToPath As String
    ToPath = FormatBackupFilePath(FilePath)
    ToPath = AppOptionsDialog.BrowseToSaveLogFile(FSO.GetParentFolderName(ToPath), FSO.GetFileName(ToPath), "Rename Log File")
    If ToPath = "" Then Exit Function
    If FSO.FileExists(ToPath) Then Call FSO.DeleteFile(ToPath)
    Call FSO.MoveFile(FilePath, ToPath)
    Call m_LogViewUpdater.Clear
    RenameLogFile = True

End Function

'CSEH: ShowErr
Private Sub RenameFileMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call RenameLogFile(m_LogViewUpdater.FilePath)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub CopyFileMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim ToPath As String
    ToPath = FormatBackupFilePath(m_LogViewUpdater.FilePath)
    ToPath = AppOptionsDialog.BrowseToSaveLogFile(FSO.GetParentFolderName(ToPath), FSO.GetFileName(ToPath), "Copy Log File")
    If ToPath = "" Then Exit Sub
    
    Dim MPS As New usMousePtrSetter
    Call MPS.Init(Screen, vbHourglass)
    Call FSO.CopyFile(m_LogViewUpdater.FilePath, ToPath)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub DeleteFileMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Dim FileName As String
    FileName = m_LogViewUpdater.FilePath
    
    Dim Msg As String
    Msg = "The log file '" & FileName & "' will be permanently deleted."
    If MsgBox(Msg, vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
    
    Call FSO.DeleteFile(FileName)
    Call m_LogViewUpdater.Clear

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub EditMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    CopyMenuItem.Enabled = Not GridLogView.SelectedItem Is Nothing
    SelectAllMenuItem.Enabled = GridLogView.ListItems.Count > 0
    FindMenuItem.Enabled = GridLogView.ListItems.Count > 0
    FindNextMenuItem.Enabled = GridLogView.ListItems.Count > 0
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub ViewMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    GridViewMenuItem.Enabled = m_LogViewUpdater.LogFormat.ColumnLayout <> CL_NONE
    GridLinesMenuItem.Enabled = GridViewEnabled
    DetailMenuItem.Enabled = GridViewEnabled

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub CopyMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    On Error Resume Next
    Dim RTB As RichTextBox
    Set RTB = ActiveControl ' DetailEdit or StreamLogView
    On Error GoTo ERROR_HANDLER

    If ActiveControl Is GridLogView Then
        Call Clipboard.Clear
        Call Clipboard.SetText(FormatSelectedLogItemsAsText)
    ElseIf Not RTB Is Nothing Then
        Call Clipboard.Clear
        Call Clipboard.SetText(RTB.SelText, vbCFText)
        Call Clipboard.SetText(RTB.SelRTF, vbCFRTF)
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub SelectAllMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    If GridViewEnabled Then
        Dim Item As MSComctlLib.ListItem
        For Each Item In GridLogView.ListItems
            Item.Selected = True
        Next Item
    Else
        StreamLogView.SelStart = 0
        StreamLogView.SelLength = Len(StreamLogView.Text) - 1
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Indicates whether a text string contains a whole word.
'
Private Function MatchesWord(ByVal Text As String, ByVal Word As String, CompareMethod As VbCompareMethod) As Boolean

    Const DELIMS As String = " .,;:'""`~!@#$%^&*()-_=+[]{}\|<>"
    Dim LeftChar As String
    Dim RightChar As String
    Dim LeftOK As Boolean
    Dim RightOK As Boolean
    Dim Pos As Long
    
    While True
    
        Pos = InStr(Pos + 1, Text, Word, CompareMethod)
        If Pos < 1 Then Exit Function
        
        If Pos - 1 = 0 Then
            LeftOK = True
        Else
            LeftChar = Mid$(Text, Pos - 1, 1)
            LeftOK = (LeftChar = "" Or InStr(DELIMS, LeftChar) > 0)
        End If
        
        If Pos + Len(Word) = Len(Text) Then
            RightOK = True
        Else
            RightChar = Mid$(Text, Pos + Len(Word), 1)
            RightOK = (RightChar = "" Or InStr(DELIMS, RightChar) > 0)
        End If
        
        If LeftOK And RightOK Then
            MatchesWord = True
            Exit Function
        End If
        
    Wend

End Function

' Returns the next item in the list view that matches the specified search
' criteria.
'
Private Function FindListItem(ByVal NewSearch As Boolean) As MSComctlLib.ListItem

    If GridLogView.ListItems.Count < 1 Then Exit Function

    Dim CompMeth As VbCompareMethod
    If m_FindOptions.MatchCase Then
        CompMeth = vbBinaryCompare
    Else
        CompMeth = vbTextCompare
    End If
    
    Dim StartPos As Long
    Dim EndPos As Long
    Dim IncAmount As Long
    Dim Match As Boolean
    
    If m_FindOptions.Down Then
        EndPos = GridLogView.ListItems.Count
        IncAmount = 1
    Else
        EndPos = 1
        IncAmount = -1
    End If
    
    If GridLogView.SelectedItem Is Nothing Then
        StartPos = 1
    Else
        StartPos = GridLogView.SelectedItem.Index
    End If
    
    If Not NewSearch Then
        StartPos = StartPos + IncAmount
    End If
        
    Dim FirstPass As Boolean
    FirstPass = True
    
TryAgain:
    Dim i As Long
    For i = StartPos To EndPos Step IncAmount
    
        Dim Item As MSComctlLib.ListItem
        Set Item = GridLogView.ListItems.Item(i)
        
        If m_FindOptions.WholeWord Then
            Match = MatchesWord(Item.Text, m_FindOptions.Text, CompMeth)
        Else
            Match = (InStr(1, Item.Text, m_FindOptions.Text, CompMeth) > 0)
        End If
        If Match Then
            Set FindListItem = Item
            Exit Function
        End If
        
        Dim S As Long
        For S = 1 To GridLogView.ColumnHeaders.Count - 1
            If m_FindOptions.WholeWord Then
                Match = MatchesWord(Item.SubItems(S), m_FindOptions.Text, CompMeth)
            Else
                Match = (InStr(1, Item.SubItems(S), m_FindOptions.Text, CompMeth) > 0)
            End If
            If Match Then
                Set FindListItem = Item
                Exit Function
            End If
        Next S
        
    Next i
    
    ' search from the beginning/end if first search failed
    If FirstPass Then
    
        FirstPass = False
        
        EndPos = StartPos
        If EndPos > GridLogView.ListItems.Count Then
            Debug.Assert m_FindOptions.Down
            EndPos = GridLogView.ListItems.Count
        ElseIf EndPos < 1 Then
            Debug.Assert Not m_FindOptions.Down
            EndPos = 1
        End If
        
        If m_FindOptions.Down Then
            StartPos = 1
        Else
            StartPos = GridLogView.ListItems.Count
        End If
        
        GoTo TryAgain
        
    End If

End Function

' Updates the display based on a found item.
'
Private Sub HandleFoundItem(ByVal Item As MSComctlLib.ListItem)

    If Item Is Nothing Then
        Call MsgBox("Unable to find '" & m_FindOptions.Text & "'.", vbExclamation)
        Exit Sub
    End If
    
    If Not GridLogView.SelectedItem Is Nothing Then GridLogView.SelectedItem.Selected = False
    GridLogView.SelectedItem = Item
    Call UpdateItemDetailView
    Call Item.EnsureVisible
    
End Sub

Private Function FindStreamText(ByVal NewSearch As Boolean) As Long

    If Not m_FindOptions.Down Then
        RaiseMsg "The stream view does not support searching up."
    End If

    Dim Options As Long
    Dim StartPos As Long
    Dim EndPos As Long
    
    If m_FindOptions.MatchCase Then
        Options = Options Or FindConstants.rtfMatchCase
    End If
    If m_FindOptions.WholeWord Then
        Options = Options Or FindConstants.rtfWholeWord
    End If
    
    ' search from insertion point
    StartPos = StreamLogView.SelStart
    EndPos = Len(StreamLogView.Text)
    If Not NewSearch Then StartPos = StartPos + 1
    FindStreamText = StreamLogView.Find(m_FindOptions.Text, StartPos, EndPos, vOptions:=Options)
    If FindStreamText >= 0 Then Exit Function
        
    ' wrap to beginning
    StartPos = 0
    EndPos = StreamLogView.SelStart
    FindStreamText = StreamLogView.Find(m_FindOptions.Text, StartPos, EndPos, vOptions:=Options)

End Function

Private Sub HandleFoundText(ByVal Position As Long)

    If Position < 0 Then
        Call MsgBox("Unable to find '" & m_FindOptions.Text & "'.", vbExclamation)
        Exit Sub
    End If

End Sub

'CSEH: ShowErr
Private Sub m_FindEvents_Find()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    m_FindOptions.Text = FindForm.TextEdit.Text
    m_FindOptions.MatchCase = FindForm.MatchCaseEdit.Value
    m_FindOptions.WholeWord = FindForm.WholeWordEdit.Value
    m_FindOptions.Down = FindForm.DownEdit.Value
    
    If GridViewEnabled Then
        Call HandleFoundItem(FindListItem(True))
    Else
        Call HandleFoundText(FindStreamText(True))
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub m_FindEvents_FindNext()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    ' exit if last find failed
    If m_FindOptions.Text = "" Then Exit Sub
    
    If GridViewEnabled Then
        Call HandleFoundItem(FindListItem(False))
    Else
        Call HandleFoundText(FindStreamText(False))
    End If
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FindMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    ' setup event handlers for find
    Set m_FindEvents = FindForm
    
    ' show find dialog
    Call FindForm.Show
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FindNextMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call m_FindEvents_FindNext

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call HandleOLEDragOver(Data, Effect)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call HandleOLEDragDrop(Data, Effect)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call HandleOLEStartDrag(Data, AllowedEffects)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call HandleOLESetData(Data, DataFormat)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLogView_OLECompleteDrag(Effect As Long)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call HandleOLECompleteDrag(Effect)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub StreamLogView_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Call HandleOLEDragOver(Data, Effect)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub StreamLogView_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Call HandleOLEDragDrop(Data, Effect)

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub DetailMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    DetailVisible = Not DetailVisible

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridLinesMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    GridLinesEnabled = Not GridLinesEnabled
    
'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub LiveMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    LiveUpdate = Not LiveUpdate

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FormatMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    SaveFormatMenuItem.Enabled = True
    Dim Fmt As LogFormat
    For Each Fmt In g_AppOptions.LogFormats
        If Fmt Is m_LogViewUpdater.LogFormat Then
            SaveFormatMenuItem.Enabled = False
        End If
    Next

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub FormatsMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    Dim TimerEnableRestorer As New usPropertyRestorer
    Call TimerEnableRestorer.Store(Timer, "Enabled", False)
    
    Call LogFormatsDialog.Execute
    Call LoadWindowCaption ' format name may have changed

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub SaveFormatMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    Dim NewFmt As LogFormat
    Set NewFmt = g_AppOptions.LogFormats.AddCopy(m_LogViewUpdater.LogFormat)
    Call NewFmt.SetNameUnique(FSO.GetBaseName(m_LogViewUpdater.FilePath))
    If m_LogViewUpdater.LogFormat.ColumnLayout <> CL_NONE Then
        NewFmt.ColumnCaptions.AsVariant = NewFmt.SplitLine(g_AppOptions.HeadLine)
    End If
    If LogFormatPropsDialog.Execute(NewFmt) Then
        'Call FormatList.AddItem(NewFmt.Name)
    Else
        Call g_AppOptions.LogFormats.Remove(g_AppOptions.LogFormats.Count)
    End If
    m_LogViewUpdater.LogFormat = NewFmt

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

Private Sub EditFormatMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>

    If LogFormatPropsDialog.Execute(m_LogViewUpdater.LogFormat) Then
        Call RefreshView
    End If

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

'CSEH: ShowErr
Private Sub GridViewMenuItem_Click()
'<EhHeader>
On Error GoTo ERROR_HANDLER
'</EhHeader>
    
    GridViewEnabled = Not GridViewEnabled
    Call RefreshView

'<EhFooter>
' GENERATED CODE: do not modify without removing EhFooter marks
Exit Sub
ERROR_HANDLER:
    VBA.MsgBox VBA.Err.Description, VBA.vbCritical
'</EhFooter>
End Sub

' Returns a reference to the log view updater.
'
Public Property Get LogViewUpdater() As LogViewUpdater
    Set LogViewUpdater = m_LogViewUpdater
End Property

