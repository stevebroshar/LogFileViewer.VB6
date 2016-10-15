Attribute VB_Name = "ErrUtil"
'
' This file contains error handling utilities.
'
' CODE REUSABILITY
' This file is designed to work in any VB project.
'
Option Explicit

Private Const E_FAIL As Long = &H80004005
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_CANCEL As Long = vbObjectError + 9765
Private Const E_USERERR As Long = vbObjectError + 9766

Public Type ERR_CACHE
    Number As Long
    Source As String
    Description As String
    HelpFile As String
    HelpContext As Long
End Type

Public Type PROCEDURE_ERROR_CACHE
    Err As ERR_CACHE
    Erl As Long
    ModuleName As String
    ProcedureName As String
End Type

' Stores the state of the global VBA.Err object into a cache variable.
'
' ADVICE
' Use this procedure to temporarily store the state of VBA.Err.  Then, code
' can be run that may modify the state of VBA.Err.  Then, the cached error
' information can be used even though the global error state has changed.
'
' NOTE
' Common operations that clear the state of the global VBA.Err object
' include: VBA.Err.Clear, COM object calls and the On Error statement.
'
Public Sub CacheErr(ByRef Cache As ERR_CACHE)

    Cache.Number = VBA.Err.Number
    Cache.Description = VBA.Err.Description
    Cache.Source = VBA.Err.Source
    Cache.HelpFile = VBA.Err.HelpFile
    Cache.HelpContext = VBA.Err.HelpContext
    
End Sub

' Loads the state of the global VBA.Err object from a cache variable.
'
Public Sub LoadErrFromCache(ByRef Cache As ERR_CACHE)

    VBA.Err.Number = Cache.Number
    VBA.Err.Description = Cache.Description
    VBA.Err.Source = Cache.Source
    VBA.Err.HelpFile = Cache.HelpFile
    VBA.Err.HelpContext = Cache.HelpContext
    
End Sub

' Stores the information associated with an error that has occurred into a
' cache variable.  This information includes the global VB error state
' (VBA.Err and VBA.Erl) and the names of the caller's source-code module and
' procedure.
'
' ADVICE
' Use this procedure to temporarily store the state associated with an error
' that has occurred.  Code can be run that may modify the global error state,
' and then, the cached error information can be used even though the global
' error state has changed.  See LogCachedProcedureError and
' RaiseCachedProcedureErrorWithLogging.
'
' NOTE
' An On Error statement clears all of the global error state -- both VBA.Err
' and VBA.Erl.  Common operations that clear the state of the global VBA.Err
' object include VBA.Err.Clear and COM object calls.
'
Public Sub CacheProcedureError _
 (ByRef Cache As PROCEDURE_ERROR_CACHE, _
  ByVal ModuleName As String, _
  ByVal ProcedureName As String)
  
    Call CacheErr(Cache.Err)
    Cache.Erl = VBA.Erl
    Cache.ModuleName = ModuleName
    Cache.ProcedureName = ProcedureName
    
End Sub

' Raises an error based on the state of the global VBA.Err object.
'
Public Sub RaiseErr()
    Debug.Assert VBA.Err.Number <> 0
    VBA.Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
End Sub

' Raises an error based on the state of cached error information.
'
Public Sub RaiseCachedErr(ByRef Cache As ERR_CACHE)
    Debug.Assert Cache.Number <> 0
    VBA.Err.Raise Cache.Number, Cache.Source, Cache.Description, Cache.HelpFile, Cache.HelpContext
End Sub

' Raises an error with a custom description and optionally a custom number.
' E_FAIL is used for the number if not specified by the caller.
'
Public Sub RaiseMsg(ByVal Description As String, Optional ByVal Number As Long = E_FAIL)
    Debug.Assert Number <> 0
    VBA.Err.Raise Number, Description:=Description
End Sub

' Raises a cancel error.
'
Public Sub RaiseCancel(Optional ByVal Description As String = "Operation canceled.")
    VBA.Err.Raise E_CANCEL, Description:=Description
End Sub

' Indicates whether the global VBA.Err object indicates cancel.
'
Public Function ErrIsCancel() As Boolean
    ErrIsCancel = (VBA.Err.Number = E_CANCEL)
End Function

' Indicates whether a cached error indicates cancel.
'
Public Function CachedErrIsCancel(ByRef Cache As ERR_CACHE) As Boolean
    CachedErrIsCancel = (Cache.Number = E_CANCEL)
End Function

' Raises a user error.
'
Public Sub RaiseUserError(ByVal Description As String)
    VBA.Err.Raise E_USERERR, Description:=Description
End Sub

' Indicates whether the global VBA.Err object indicates user-error.
'
Public Function ErrIsUserError() As Boolean
    ErrIsUserError = (VBA.Err.Number = E_USERERR)
End Function

' Indicates whether a cached error indicates user-error.
'
Public Function CachedErrIsUserError(ByRef Cache As ERR_CACHE) As Boolean
    CachedErrIsUserError = (Cache.Number = E_USERERR)
End Function

' Raises a not implemented error.
'
Public Sub RaiseNotImplemented(Optional ByVal Description As String = "Operation not implemented.")
    VBA.Err.Raise E_NOTIMPL, Description:=Description
End Sub
