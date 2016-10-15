VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class provides a collection of LogFormat objects.
'
'CSEH: Skip
Option Explicit

Private m_Items As New usObjectCollection

' Returns the number of items in the collection.
'
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
    Count = m_Items.Count
End Function

' Returns the position of the item that is the same instance or 0 if not
' found.
'
Public Function Find(ByVal Item As LogFormat) As Long
    Dim Pos As Long
    For Pos = 1 To m_Items.Count
        If m_Items(Pos) Is Item Then
            Find = Pos
            Exit Function
        End If
    Next
End Function

' Returns a reference to the first log format found with the specified name
' or Nothing if no match is found.
'
Public Function FindByName(ByVal Name As String) As LogFormat
    Dim Item As LogFormat
    For Each Item In m_Items
        If Item.Name = Name Then
            Set FindByName = Item
            Exit Function
        End If
    Next
End Function

' Returns the collection item specified by index.
'
Public Function Item(Index As Variant) As LogFormat
    Set Item = m_Items(Index)
End Function

' Adds a new item to the collection.
'
Public Function AddNew(Optional ByVal Name As String = "Format") As LogFormat
    Dim NewItem As New LogFormat
    Call NewItem.SetNameUnique(Name)
    Call m_Items.Add(NewItem)
    Set AddNew = NewItem
End Function

' Adds a new item to the collection based on an existing log format as a
' template.
'
Public Function AddCopy(ByVal Template As LogFormat) As LogFormat
    Dim NewItem As New LogFormat
    NewItem.AsText = Template.AsText
    Call NewItem.SetNameUnique(Template.Name)
    Call m_Items.Add(NewItem)
    Set AddCopy = NewItem
End Function

' Removes the collection item specified by index.
'
Public Sub Remove(Index As Variant)
    Call m_Items.Remove(Index)
End Sub

' Removes all items from the collection.
'
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items from the collection"
    Call m_Items.Clear
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
    Dim Item As LogFormat
    For Each Item In m_Items
        Call Text.Add(Item.AsText)
    Next
    AsText = Text.AsCSV
End Property

Public Property Let AsText(RHS As String)
    Call Clear
    Dim i As Long
    Dim Item As LogFormat
    Dim Text As New usStringList
    Text.AsCSV = RHS
    For i = 1 To Text.Count
        Set Item = New LogFormat
        Item.AsText = Text(i)
        Call m_Items.Add(Item)
    Next
End Property

Private Function FindMatchingFileNames(ByVal Pattern As String) As usStringList
    Dim Result As New usStringList
    Dim FoundName As String
    FoundName = Dir(Pattern)
    While Len(FoundName) > 0
        Call Result.Add(FoundName)
        FoundName = Dir()
    Wend
    Set FindMatchingFileNames = Result
End Function

Public Sub LoadAll()
    Dim LogFmtFileNames As usStringList
    Set LogFmtFileNames = FindMatchingFileNames(g_AppOptions.LogFormatsPath("*"))
    Dim i As Long
    For i = 1 To LogFmtFileNames.Count
        Dim NewItem As LogFormat
        Set NewItem = New LogFormat
        Call NewItem.Load(LogFmtFileNames(i))
        Call m_Items.Add(NewItem)
    Next
End Sub

Public Sub SaveAll()
    Dim Item As LogFormat
    For Each Item In m_Items
        Call Item.Save
    Next
End Sub


