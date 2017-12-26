VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private thisSize As Long
Private elements() As Variant

Public Sub Class_Initialize()
    thisSize = 0
End Sub

Public Function addItem(ByVal item As Variant)
    ReDim Preserve elements(thisSize)
    If IsObject(item) Then
        Set elements(thisSize) = item
    Else
        elements(thisSize) = item
    End If
    thisSize = thisSize + 1
End Function

Public Function getItem(index As Long)
    getItem = elements(index)
End Function

Public Function getObject(index As Long)
    Set getObject = elements(index)
End Function

Public Function setItem(index As Long, ByVal item As Variant)
        If IsObject(item) Then
                Set elements(index) = item
        Else
                elements(index) = item
        End If
End Function

Public Function removeItem(index As Long)
    Dim result As Boolean
    ReDim tmpElements(thisSize - 1) As Variant
    result = False
    
    Dim i As Long
    Dim j As Long
    For i = 0 To thisSize - 1
        If index = i Then
            result = True
        Else
            tmpElements(i) = elements(i)
            j = j + 1
        End If
    Next
    
    If result = True Then
        thisSize = thisSize - 1
    End If
    removeItem = result
End Function

Public Function containsValue(item As Variant)
    Dim result As Boolean
    result = False
    Dim i As Long
    For i = 0 To thisSize - 1
        If elements(i) = item Then
            result = True
        End If
    Next
    containsValue = result
End Function

Public Function size()
    size = thisSize
End Function
