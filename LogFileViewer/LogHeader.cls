VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Width As Long
Private m_HeaderText As String

Public Property Get Width() As Long

End Property

Public Property Let Width(ByVal RHS As Long)

End Property

Public Property Get HeaderText() As Long

End Property

Public Property Let HeaderText(ByVal RHS As Long)

End Property

Public Property Get Count() As Long

End Property

ColumnWidths() As Long
ColumnHeaders() As String
FileNameExtensions() as string or maybe FileNamePatterns()


