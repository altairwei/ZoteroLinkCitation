' CustomDictionary - A lightweight Dictionary replacement for macOS compatibility.
' On macOS, Scripting.Dictionary is not available. This class provides the same
' interface used by ZoteroLinkCitation: Add, Item, Keys, Count, Exists.
'
' INSTALLATION (macOS):
'   1. In the VBA Editor, go to Insert > Class Module
'   2. In the Properties window (F4 or View > Properties), rename the class to "CustomDictionary"
'   3. Paste ALL the code below into the class module

Option Explicit

Private mKeys() As Variant
Private mValues() As Variant
Private mCount As Long
Private mCapacity As Long

Private Sub Class_Initialize()
    mCount = 0
    mCapacity = 16
    ReDim mKeys(1 To mCapacity)
    ReDim mValues(1 To mCapacity)
End Sub

Private Sub EnsureCapacity()
    If mCount >= mCapacity Then
        mCapacity = mCapacity * 2
        ReDim Preserve mKeys(1 To mCapacity)
        ReDim Preserve mValues(1 To mCapacity)
    End If
End Sub

Private Function FindKey(ByVal key As Variant) As Long
    Dim i As Long
    For i = 1 To mCount
        If mKeys(i) = key Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = 0
End Function

' Add a new key-value pair (matches Scripting.Dictionary.Add behavior)
Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    EnsureCapacity
    mCount = mCount + 1
    mKeys(mCount) = key
    mValues(mCount) = value
End Sub

' Read a value by key: value = dict.Item(key)
Public Property Get Item(ByVal key As Variant) As Variant
    Dim idx As Long
    idx = FindKey(key)
    If idx > 0 Then
        Item = mValues(idx)
    Else
        Item = Empty
    End If
End Property

' Write a value by key: dict.Item(key) = value
Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
    Dim idx As Long
    idx = FindKey(key)
    If idx > 0 Then
        mValues(idx) = value
    Else
        Add key, value
    End If
End Property

' Return the number of items
Public Property Get Count() As Long
    Count = mCount
End Property

' Return all keys as a 0-based Variant array (matches Scripting.Dictionary.Keys)
Public Function Keys() As Variant
    If mCount = 0 Then
        Keys = Array()
        Exit Function
    End If
    Dim result() As Variant
    ReDim result(0 To mCount - 1)
    Dim i As Long
    For i = 1 To mCount
        result(i - 1) = mKeys(i)
    Next i
    Keys = result
End Function

' Check if a key exists
Public Function Exists(ByVal key As Variant) As Boolean
    Exists = (FindKey(key) > 0)
End Function
