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
Private m_Caption As String

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(ByVal RHS As Long)
    m_Width = RHS
End Property

Public Property Get Caption() As Long
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal RHS As Long)
    m_Caption = RHS
End Property

Public Property Get AsText() As String
    Dim Text As New usStringList
    Text.Value("Width") = m_Width
    Text.Value("Caption") = m_Caption
    AsText = Text.AsCSV
End Property

Public Property Let AsText(RHS As String)
    Dim Text As New usStringList
    Text.AsCSV = RHS
    m_Width = Text.Value("Width")
    m_Caption = Text.Value("Caption")
End Property


