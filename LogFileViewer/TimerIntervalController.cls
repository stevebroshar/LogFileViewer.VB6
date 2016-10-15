VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerIntervalController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COUNT As Long = 8
Private Const MIN_INTERVAL As Long = 1000
Private m_Count As Long
Private m_NextPos As Long
Private m_Durations(MAX_COUNT - 1) As Long
Private m_Timer As VB.Timer

' Attaches a timer object to this instance.
'
Public Sub Attach(ByVal Timer As VB.Timer)
    
    Set m_Timer = Timer

End Sub

' Clears the history and sets the timer interval to the minimum.
'
Public Sub Clear()
    
    m_Count = 0
    m_NextPos = 0
    m_Timer.Interval = MIN_INTERVAL

End Sub

' Adjusts the timer interval based on a duration and several of the
' preceeding durations.
'
Public Sub Adjust(ByVal Duration As Long)
    
    If m_Count < MAX_COUNT Then
        m_Count = m_Count + 1
    End If
    m_Durations(m_NextPos) = Duration
    m_NextPos = m_NextPos + 1
    If m_NextPos >= MAX_COUNT Then
        m_NextPos = 0
    End If
    Dim DurationSum As Long
    Dim i As Long
    For i = 0 To m_Count - 1
        DurationSum = DurationSum + m_Durations(i)
    Next
    Dim AverageDuration As Long
    AverageDuration = DurationSum / m_Count
    Dim NewInterval As Long
    NewInterval = AverageDuration * 10
    If NewInterval < MIN_INTERVAL Then NewInterval = MIN_INTERVAL
    m_Timer.Interval = NewInterval

End Sub

