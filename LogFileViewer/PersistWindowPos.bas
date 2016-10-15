Attribute VB_Name = "PersistWindowPos"
Option Explicit

Private Type POINTAPI
    x   As Long
    y   As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type WINDOWPLACEMENT
    Length              As Long
    Flags               As Long
    showCmd             As Long
    ptMinPosition       As POINTAPI
    ptMaxPosition       As POINTAPI
    rcNormalPosition    As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

' Returns a string that defines a form's position and size state.
'
Public Function GetWindowPlacementDefinition(ByVal Form As Form) As String
    
    On Error GoTo ERROR_HANDLER
    
    ' read window placement info
    Dim WinPlace As WINDOWPLACEMENT
    WinPlace.Length = LenB(WinPlace)
    If GetWindowPlacement(Form.hwnd, WinPlace) = 0 Then
        Debug.Print "SaveWindowPlacement: Unable to read window placement."
        Exit Function
    End If

    Dim Props As New usStringList
    
    ' store settings
    Props.Value("Left") = WinPlace.rcNormalPosition.Left
    Props.Value("Top") = WinPlace.rcNormalPosition.Top
    Props.Value("Right") = WinPlace.rcNormalPosition.Right
    Props.Value("Bottom") = WinPlace.rcNormalPosition.Bottom
    Props.Value("Flags") = WinPlace.Flags
    Props.Value("Length") = WinPlace.Length
    Props.Value("MaxX") = WinPlace.ptMaxPosition.x
    Props.Value("MaxY") = WinPlace.ptMaxPosition.y
    Props.Value("MinX") = WinPlace.ptMinPosition.x
    Props.Value("MinY") = WinPlace.ptMinPosition.y
    
    ' save window-state -- not allowing minimized
    Dim WinState As Integer
    WinState = Form.WindowState
    If WinState = vbMinimized Then WinState = vbNormal
    Props.Value("WindowState") = WinState
    
    GetWindowPlacementDefinition = Props.AsCSV
    
    Exit Function
    
ERROR_HANDLER:
    
    ' prevent errors from propagating
    Debug.Print "SaveWindowPlacement: " & Err.Description
    Resume Next

End Function

Private Function ReadProp(ByVal Props As usStringList, ByVal Name As String, ByVal Default As String)
    
    Dim Text As String
    Text = Props.Value(Name)
    If Len(Text) = 0 Then Text = Default
    ReadProp = Text

End Function

' Sets a form's position and size state from a previously saved definition.
'
' SIDE EFFECTS
' This routine causes the form's resize function to occur.  Note that this
' will cause the Form_Resize function to run if there is one for the form.
'
Public Sub RestoreWindowPlacement(ByVal Form As Form, ByVal Definition As String)

    On Error GoTo ERROR_HANDLER
        
    ' load current window placement info
    Dim WinPlace As WINDOWPLACEMENT
    WinPlace.Length = LenB(WinPlace)
    If GetWindowPlacement(Form.hwnd, WinPlace) = 0 Then
        Debug.Print "RestoreWindowPlacement: Unable to read window placement."
        Exit Sub
    End If
    
    Dim Props As New usStringList
    Props.AsCSV = Definition
    
    ' load settings
    WinPlace.Flags = ReadProp(Props, "Flags", WinPlace.Flags)
    WinPlace.rcNormalPosition.Left = ReadProp(Props, "Left", WinPlace.rcNormalPosition.Left)
    WinPlace.rcNormalPosition.Top = ReadProp(Props, "Top", WinPlace.rcNormalPosition.Top)
    WinPlace.rcNormalPosition.Right = ReadProp(Props, "Right", WinPlace.rcNormalPosition.Right)
    WinPlace.rcNormalPosition.Bottom = ReadProp(Props, "Bottom", WinPlace.rcNormalPosition.Bottom)
    WinPlace.ptMaxPosition.x = ReadProp(Props, "MaxX", WinPlace.ptMaxPosition.x)
    WinPlace.ptMaxPosition.y = ReadProp(Props, "MaxY", WinPlace.ptMaxPosition.y)
    WinPlace.ptMinPosition.x = ReadProp(Props, "MinX", WinPlace.ptMinPosition.x)
    WinPlace.ptMinPosition.y = ReadProp(Props, "MinY", WinPlace.ptMinPosition.y)

    ' apply settings to form
    ' NOTE: do this even if nothing read since this causes resize
    If SetWindowPlacement(Form.hwnd, WinPlace) = 0 Then
        Debug.Print "Unable to set window placement."
    End If

    ' apply window-state to form
    ' NOTE: do this after SetWindowPlacement
    Form.WindowState = ReadProp(Props, "WindowState", Form.WindowState)

    Exit Sub
    
ERROR_HANDLER:
    ' prevent errors from propagating
    Debug.Print "RestoreWindowPlacement: " & Err.Description
    Resume Next

End Sub

