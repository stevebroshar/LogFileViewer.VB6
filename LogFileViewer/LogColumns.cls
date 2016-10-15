VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class provides a collection of LogColumn objects.
'
Option Explicit

Private m_Items As New usObjectCollection

' Adds a new item to the collection.
'
Public Function Add() As LogColumn
Attribute Add.VB_Description = "Adds an item to the collection"
    Dim NewItem As New LogColumn
    Call m_Items.Add(NewItem)
    Set Add = NewItem
End Function

' Returns the number of items in the collection.
'
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
    Count = m_Items.Count
End Function

' Indicates whether an item exists in the collection at the specified by index.
'
Public Function Exists(ByVal Index As Long) As Boolean
Attribute Exists.VB_Description = "Indicates whether an item exists in the collection as specified by index"
    Exists = m_Items.Exists(Index)
End Function

' Returns the collection item specified by index.
'
Public Function Item(ByVal Index As Long) As LogColumn
Attribute Item.VB_Description = "Returns an item from the collection either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_Items(Index)
End Function

' Removes the collection item specified by index.
'
Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes an item from the collection"
    Call m_Items.Remove(Index)
End Sub

' Removes all items from the collection.
'
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items from the collection"
    m_Items.Clear
End Sub

' Implements support for enumeration (For Each).
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Items.NewEnum
End Function

Public Property Get AsText() As String
    Dim Text As New usStringList
    Dim Item As LogColumn
    For Each Item In m_Items
        Call Text.Add(Item.AsText)
    Next
    AsText = Text.AsCSV
End Property

Public Property Let AsText(RHS As String)
    Call Clear
    Dim NewItem As New LogColumn
    Dim Text As New usStringList
    Text.AsCSV = RHS
    Dim i As Long
    For i = 1 To Text.Count
        NewItem.AsText = Text(i)
        Call m_Items.Add(NewItem)
    Next
End Property


