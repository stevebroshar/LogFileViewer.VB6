VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RichTextSaveSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RichTextBox As RichTextLib.RichTextBox
Private m_OrigActive As Boolean
Private m_OrigEnabled As Boolean
Private m_OrigSelStart As Long
Private m_OrigSelLen As Long

Private Sub Class_Initialize()
    
    Call Clear

End Sub

Private Sub Class_Terminate()
    
    Call Restore

End Sub

Public Sub Clear()
    
    Set m_RichTextBox = Nothing
    m_OrigSelStart = -1
    m_OrigSelLen = -1

End Sub

Public Sub Save(ByVal RichTextBox As RichTextLib.RichTextBox)
    
    Debug.Assert Not RichTextBox Is Nothing
    
    m_OrigActive = ViewerWindow.ActiveControl Is RichTextBox
    
    ' disable control to prevent user from changing selection until Restore
    m_OrigEnabled = RichTextBox.Enabled
    RichTextBox.Enabled = False

    ' cache current selection
    If RichTextBox.SelStart < Len(RichTextBox.Text) Then
        m_OrigSelStart = RichTextBox.SelStart
        m_OrigSelLen = RichTextBox.SelLength
    End If

    Set m_RichTextBox = RichTextBox

End Sub

Public Sub Restore()
    
    If Not m_RichTextBox Is Nothing Then
        If m_OrigSelStart >= 0 Then
            m_RichTextBox.SelStart = m_OrigSelStart
            m_RichTextBox.SelLength = m_OrigSelLen
        End If
        m_RichTextBox.Enabled = m_OrigEnabled
        If m_OrigActive Then
            m_RichTextBox.SetFocus
        End If
    End If

End Sub

