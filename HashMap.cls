VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private allCount As Long
Private arraySize As Long
Private keyArray As Variant
Private valArray As Variant

Public Sub Class_Initialize()
    Call init(1009)
End Sub

'reset and init
Public Sub init(size As Long)
    allCount = 0
    arraySize = size
    ReDim keyArray(arraySize)
    ReDim valArray(arraySize)
        
    Dim i As Long
    For i = 0 To arraySize
        Set keyArray(i) = New ArrayList
        Set valArray(i) = New ArrayList
    Next
End Sub

'put item
Public Function putItem(key As Variant, val As Variant)
    Dim index As Long
    index = getIndex(key)
        
    Dim i As Long
    If keyArray(index).size() >= 1 Then
        For i = 0 To keyArray(index).size() - 1
            If keyArray(index).getItem(i) = key Then
                Call valArray(index).setItem(i, val)
                Exit Function
            End If
        Next
    End If
        
    Call keyArray(index).addItem(key)
    Call valArray(index).addItem(val)
    allCount = allCount + 1
End Function

'get not object item
Public Function getItem(key As Variant)
    Dim item As Variant
    Dim index As Long
    index = getIndex(key)
        
    Dim i As Long
    For i = 0 To keyArray(index).size() - 1
        If keyArray(index).getItem(i) = key Then
            item = valArray(index).getItem(i)
        End If
    Next
    getItem = item
End Function

'search contains then get indexes
Public Function containsKey(key As Variant, _
  Optional ByRef arrayIndex As Long, _
  Optional ByRef listIndex As Long)
    Dim result As Boolean
    Dim index As Long
    result = False
    index = getIndex(key)
        
    Dim i As Long
    For i = 0 To keyArray(index).size() - 1
        If keyArray(index).getItem(i) = key Then
            containsKey = True
            arrayIndex = index
            listIndex = i
        End If
    Next
    containsKey = result
End Function

'return allNodes from keyArray
Public Function getKeyArray()
    Dim result(allCount) As Variant
        
    Dim count As Long
    Dim i As Long
    Dim j As Long
    count = 0
    For i = 0 To arraySize
        For j = 0 To keyArray(i).size() - 1
            result(count) = keyArray(i).getItem(j)
            count = count + 1
        Next
    Next
End Function

'remove item
Public Function removeItem(key As String)
    Dim index As Long
    index = getIndex(key)
        
    Dim i As Long
    For i = 0 To keyArray(index).size() - 1
        If keyArray(index).getItem(i) = key Then
            Call keyArray(index).removeItem(i)
            Call valArray(index).removeItem(i)
            allCount = allCount - 1
        End If
    Next
End Function

'HashMap - size()
Public Function size()
    size = allCount
End Function

'get hash of variant then mod arraySize
Private Function getIndex(str As Variant)
    Dim hash As Long
    Dim tmp() As Byte
        
    hash = 0
    tmp() = StrConv(str, vbFromUnicode)

    Dim i As Long
    For i = 0 To UBound(tmp())
        If hash > 60000000 Then
            hash = hash Mod 60000000
        End If
        hash = hash * CLng(31) + CLng(tmp(i))
    Next

    getIndex = hash Mod arraySize
End Function


