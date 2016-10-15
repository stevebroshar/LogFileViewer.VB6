VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private m_GridLines As Boolean
Private m_DetailVisible As Boolean
Private m_DetailWidth As Long
Private m_MaxFileSizeBytes As Long
Private m_LogFileNameSource As LOG_FILENAME_SOURCE
Private m_TempDirLogFileName As String
Private m_AbsLogFileName As String
Private m_LogFileNameProgID As String
Private m_LogFileNameProcName As String
Private m_LogFileNameCallType As VBA.VbCallType
Private m_LogFormats As New LogFormats
Private m_MruFilePaths As New usStringList
Public HeadLine As String
Private m_AppDataFolderPath As String
Private m_PrefsFilePath As String
Private m_LogFormatsFolderPath As String
Private m_Storage As New usStringList

Public Enum LOG_FILENAME_SOURCE
    LFS_TEMP = 1
    LFS_ABSOLUTE = 2
    LFS_OBJECT = 3
End Enum

Private Sub Class_Initialize()

    m_GridLines = True
    m_DetailVisible = True
    m_DetailWidth = 3000
    m_MaxFileSizeBytes = 1000000
    m_LogFileNameSource = LFS_TEMP
    m_TempDirLogFileName = "PlanningErrorLog.csv"
    m_LogFileNameProgID = "epLogger_80.Logger"
    m_LogFileNameProcName = "FileName"
    m_LogFileNameCallType = VBA.VbGet

    Dim Shell As New usSHObject
    m_AppDataFolderPath = FSO.BuildPath(Shell.SpecialFolderPath(SFID_APPDATA), "Cognos")
    If Dir(m_AppDataFolderPath, vbDirectory) = "" Then Call MkDir(m_AppDataFolderPath)
    m_AppDataFolderPath = FSO.BuildPath(m_AppDataFolderPath, App.Title)
    If Dir(m_AppDataFolderPath, vbDirectory) = "" Then Call MkDir(m_AppDataFolderPath)
    m_PrefsFilePath = FSO.BuildPath(m_AppDataFolderPath, "Preferences")
    m_LogFormatsFolderPath = FSO.BuildPath(m_AppDataFolderPath, "LogFormats")
    If Dir(m_LogFormatsFolderPath, vbDirectory) = "" Then Call MkDir(m_LogFormatsFolderPath)

End Sub

Private Function ReadString(ByVal Name As String, ByVal Default As String) As String
    
    Dim LinePos As Long
    LinePos = m_Storage.FindName(Name)
    If LinePos > 0 Then
        ReadString = m_Storage.ValueAt(LinePos)
    Else
        ReadString = Default
    End If
    
End Function

Private Function ReadLong(ByVal Name As String, ByVal Default As Long) As Long
    
    On Error GoTo ERROR_HANDLER
    
    ReadLong = ReadString(Name, Default)
    
    Exit Function

ERROR_HANDLER:
    
    ReadLong = Default
    
End Function

Private Sub LoadLongList(ByVal Value As usLongList, ByVal Name As String, ByVal Default As String)
    
    On Error GoTo ERROR_HANDLER
    
    Value.AsVariant = Split(ReadString(Name, Default), ",")
    
    Exit Sub

ERROR_HANDLER:
    
    Value.Clear

End Sub

' Loads the application options from the user's application data folder.
'
Public Sub Load()

    On Error GoTo ERROR_HANDLER
    
    Call m_Storage.Clear
    If FSO.FileExists(m_PrefsFilePath) Then
        Call m_Storage.LoadFromFile(m_PrefsFilePath)
    End If
    
    Dim Def As New AppOptions
    
    Call RestoreWindowPlacement(ViewerWindow, ReadString("MainWindow", "")) ' causes a resize
    m_GridLines = ReadLong("GridLines", Def.GridLines)
    m_DetailVisible = ReadLong("DetailVisible", Def.DetailVisible)
    m_DetailWidth = ReadLong("DetailWidth", Def.DetailWidth)
    m_MaxFileSizeBytes = ReadLong("MaxFileSize", Def.MaxFileSizeBytes)
    m_LogFileNameSource = ReadLong("LogFileNameSource", Def.LogFileNameSource)
    m_TempDirLogFileName = ReadString("TempDirLogFileName", Def.TempDirLogFileName)
    m_AbsLogFileName = ReadString("AbsoluteLogFileName", Def.AbsoluteLogFileName)
    m_LogFileNameProgID = ReadString("LogFileNameProgID", Def.LogFileNameProgID)
    m_LogFileNameProcName = ReadString("LogFileNameProcName", Def.LogFileNameProcName)
    m_LogFileNameCallType = ReadLong("LogFileNameCallType", Def.LogFileNameCallType)
    m_MruFilePaths.AsCSV = ReadString("MruFiles", "")
    
    Call m_LogFormats.LoadAll

    Exit Sub
    
ERROR_HANDLER:
    ' prevent errors from propagating
    Debug.Print "AppOptions.Load: " & Err.Description
    Resume Next

End Sub

Private Sub WriteString(ByVal Name As String, ByVal Value As String, Optional ByVal Default)

    If Not IsMissing(Default) Then
        If Value = Default Then
            'On Error Resume Next ' ignore error if name not found
            'Call DeleteSetting(App.Title, SECTION_NAME, Name)
            Exit Sub
        End If
    End If
    m_Storage.Value(Name) = Value
    
End Sub

Private Sub WriteLong(ByVal Name As String, ByVal Value As Long, Optional ByVal Default)

    If IsMissing(Default) Then
        Call WriteString(Name, Value)
    Else
        Call WriteString(Name, Value, CLng(Default))
    End If
    
End Sub

' Saves the application options to the user's application data folder.
'
Public Sub Save()

    On Error GoTo ERROR_HANDLER
    
    Dim Def As New AppOptions
    
    Call m_Storage.Clear
        
    Call WriteString("MainWindow", GetWindowPlacementDefinition(ViewerWindow))
    Call WriteLong("GridLines", m_GridLines, Def.GridLines)
    Call WriteLong("DetailVisible", m_DetailVisible, Def.DetailVisible)
    Call WriteLong("DetailWidth", m_DetailWidth, Def.DetailWidth)
    Call WriteLong("MaxFileSize", m_MaxFileSizeBytes, Def.MaxFileSizeBytes)
    Call WriteLong("LogFileNameSource", m_LogFileNameSource, Def.LogFileNameSource)
    Call WriteString("TempDirLogFileName", m_TempDirLogFileName, Def.TempDirLogFileName)
    Call WriteString("AbsoluteLogFileName", m_AbsLogFileName, Def.AbsoluteLogFileName)
    Call WriteString("LogFileNameProgID", m_LogFileNameProgID, Def.LogFileNameProgID)
    Call WriteString("LogFileNameProcName", m_LogFileNameProcName, Def.LogFileNameProcName)
    Call WriteLong("LogFileNameCallType", m_LogFileNameCallType, Def.LogFileNameCallType)
    Call WriteString("MruFiles", m_MruFilePaths.AsCSV)

    Call m_Storage.SaveToFile(m_PrefsFilePath)
    
    Exit Sub
    
ERROR_HANDLER:
    ' prevent errors from propagating
    Debug.Print "AppOptions.Save: " & Err.Description
    Resume Next

End Sub

'Public Property Get LogFormatsFolderPath() As String
'    LogFormatsFolderPath = m_LogFormatsFolderPath
'End Property
Public Property Get LogFormatsPath(ByVal RelPath As String) As String
    LogFormatsPath = FSO.BuildPath(m_LogFormatsFolderPath, RelPath)
End Property

Public Property Get GridLines() As Boolean
    GridLines = m_GridLines
End Property
Public Property Let GridLines(ByVal RHS As Boolean)
    m_GridLines = RHS
End Property

Public Property Get DetailVisible() As Boolean
    DetailVisible = m_DetailVisible
End Property
Public Property Let DetailVisible(ByVal RHS As Boolean)
    m_DetailVisible = RHS
End Property

Public Property Get DetailWidth() As Long
    DetailWidth = m_DetailWidth
End Property
Public Property Let DetailWidth(ByVal RHS As Long)
    m_DetailWidth = RHS
End Property

Public Property Get MaxFileSizeBytes() As Long
    MaxFileSizeBytes = m_MaxFileSizeBytes
End Property
Public Property Let MaxFileSizeBytes(ByVal RHS As Long)
    m_MaxFileSizeBytes = RHS
End Property

Public Property Get LogFileNameSource() As LOG_FILENAME_SOURCE
    LogFileNameSource = m_LogFileNameSource
End Property
Public Property Let LogFileNameSource(ByVal RHS As LOG_FILENAME_SOURCE)
    m_LogFileNameSource = RHS
End Property

Public Property Get TempDirLogFileName() As String
    TempDirLogFileName = m_TempDirLogFileName
End Property
Public Property Let TempDirLogFileName(ByVal RHS As String)
    m_TempDirLogFileName = RHS
End Property

Public Property Get AbsoluteLogFileName() As String
    AbsoluteLogFileName = m_AbsLogFileName
End Property
Public Property Let AbsoluteLogFileName(ByVal RHS As String)
    m_AbsLogFileName = RHS
End Property

Public Property Get LogFileNameProgID() As String
    LogFileNameProgID = m_LogFileNameProgID
End Property
Public Property Let LogFileNameProgID(ByVal RHS As String)
    m_LogFileNameProgID = RHS
End Property

Public Property Get LogFileNameProcName() As String
    LogFileNameProcName = m_LogFileNameProcName
End Property
Public Property Let LogFileNameProcName(ByVal RHS As String)
    m_LogFileNameProcName = RHS
End Property

Public Property Get LogFileNameCallType() As VBA.VbCallType
    LogFileNameCallType = m_LogFileNameCallType
End Property
Public Property Let LogFileNameCallType(ByVal RHS As VBA.VbCallType)
    m_LogFileNameCallType = RHS
End Property

Public Property Get LogFormats() As LogFormats
    Set LogFormats = m_LogFormats
End Property

Public Property Get MruFilePaths() As usStringList
    Set MruFilePaths = m_MruFilePaths
End Property

' Returns the string from calling the procedure that is setup to read the log
' file name.
'
Public Function ReadLogFilenameViaCall() As String

    Dim Obj As Object
    Set Obj = CreateObject(m_LogFileNameProgID)
    ReadLogFilenameViaCall = VBA.CallByName(Obj, m_LogFileNameProcName, m_LogFileNameCallType)

End Function

' Returns the default log file name as selected by user options.
'
Public Property Get DefaultLogFilename() As String

    If LogFileNameSource = LFS_TEMP Then
        
        DefaultLogFilename = FSO.BuildPath(FSO.GetSpecialFolder(Scripting.TemporaryFolder), m_TempDirLogFileName)
    
    ElseIf LogFileNameSource = LFS_ABSOLUTE Then
        
        DefaultLogFilename = m_AbsLogFileName
    
    ElseIf LogFileNameSource = LFS_OBJECT Then
        
        DefaultLogFilename = ReadLogFilenameViaCall
        
    Else
        
        RaiseMsg "INTERNAL ERROR: bad log file name source"
    
    End If
    
End Property
