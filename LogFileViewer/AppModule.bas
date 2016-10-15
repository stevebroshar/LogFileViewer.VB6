Attribute VB_Name = "AppModule"
Option Explicit

Public g_AppOptions As New AppOptions
Public FSO As New Scripting.FileSystemObject

Private m_ExeModuleVersion As String

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

' Returns the absolute path of an executable (DLL/EXE) module as specified
' by an module-instance-handle.
'
' NOTE
' This procedure could be exposed as public since it could be generally useful.
' But it is not public today since its functionality is not highly related
' to just error handling.
'
' IMPLEMENTATION
' First, the wide-string version of an API is called to obtain the file name.
' If it fails, the narrow-string version is called.  The wide-string version
' should work on NT-based systems and the narrow-string version should work
' on 9X systems.
'
Private Function GetModuleFileNameViaHandle(ByVal ModuleHandle As Long) As String

    Const MAXPATH As Long = 260
    Dim NameBuf As String
    NameBuf = Space$(MAXPATH * 2)

    On Error GoTo WIDE_API_FAILED
    If GetModuleFileNameW(ModuleHandle, NameBuf, Len(NameBuf) / 2) = 0 Then GoTo WIDE_API_FAILED
    On Error GoTo 0

    GetModuleFileNameViaHandle = StrConv(NameBuf, vbFromUnicode)

    Exit Function

WIDE_API_FAILED:

    On Error GoTo 0
    If GetModuleFileNameA(ModuleHandle, NameBuf, Len(NameBuf)) = 0 Then VBA.Err.Raise GetLastError, , "Unable to get file name for module instance"
    GetModuleFileNameViaHandle = NameBuf
    Exit Function

End Function

' Insures that the exe-module-version cache string is loaded.  If it is empty,
' then is is loaded with a string describing the version of the executable
' (DLL/EXE) module that is calling this procedure.
'
' NOTE
' Obtaining the exe-module-version could be exposed as public since it could
' be generally useful.  But it is not public today since its functionality is
' not highly related to just error handling.  Making this functionality
' public is not as simple as making this procedure public.  The public
' functionality would need to expose (for example return) the string -- not
' just store it in a private memory variable.
'
' SIDE-EFFECTS
' This procedure may clear the global error state -- VBA.Err and VBA.Erl.
'
' IMPLEMENTATION
' The VB.App object supports reading the first, second and fourth number of the
' file version -- but not the third.  Often the third number is zero, but not
' always.  Therefore, the file's version information is accessed to read the
' entire version string.  If any error occurs, the version reported by VB.App
' is returned with a question mark in place of the third number.
'
Private Function CacheExeModuleVersion() As String

    ' exit if already cached
    If Len(m_ExeModuleVersion) > 0 Then Exit Function

    ' NOTE: this clears Err and Erl
    On Error GoTo UNABLE_TO_READ_VER_INFO

    ' get absolute path name of module file
    Dim ModuleFileName As String
    ModuleFileName = GetModuleFileNameViaHandle(VB.App.hInstance)

    ' get version string
    Dim FSO As Object 'FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    m_ExeModuleVersion = FSO.GetFileVersion(ModuleFileName)
    Set FSO = Nothing

    Exit Function

UNABLE_TO_READ_VER_INFO:

    Debug.Assert False
    Dim Msg As String
    Msg = "CacheExeModuleVersion: Unable to get version.  " & Err.Description
    Debug.Print Msg
    App.LogEvent Msg

    ' default to the App object version info
    m_ExeModuleVersion = VB.App.Major & "." & VB.App.Minor & ".?." & VB.App.Revision

End Function

' Returns the application version string.
'
Public Function AppVersion() As String
    Call CacheExeModuleVersion
    AppVersion = m_ExeModuleVersion
End Function

' Selects all of the text of a TextBox-like control (i.e. TextBox, RichTextBox).
'
Public Sub SelectAllTextBox(ByVal TextBox As Object)
    TextBox.SelStart = 0
    TextBox.SelLength = Len(TextBox.Text)
End Sub

' Sets the input focus to a TextBox-like control and selects all of its text.
'
Public Sub ActivateTextBox(ByVal TextBox As Object)
    Call TextBox.SetFocus
    Call SelectAllTextBox(TextBox)
End Sub

