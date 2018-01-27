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

'初期化
Public Sub Class_Initialize()
    thisSize = 0
End Sub

'アイテム追加
Public Function addItem(ByVal item As Variant) As Boolean
    Dim result As Boolean
    result = False

    On Error GoTo Finally
    ReDim Preserve elements(thisSize)
    If IsObject(item) Then
        Set elements(thisSize) = item
    Else
        elements(thisSize) = item
    End If
    thisSize = thisSize + 1
    result = True
    
Finally:
    addItem = result
End Function

'アイテム取得
Public Function getItem(index As Long) As Variant
    getItem = elements(index)
End Function

'アイテム取得(オブジェクト)
Public Function getObject(index As Long) As Object
    Set getObject = elements(index)
End Function

'アイテムを指定したインデックスに上書き
Public Function setItem(index As Long, ByVal item As Variant) As Boolean
        Dim result As Boolean
        result = False
        
        On Error GoTo Finally
        If IsObject(item) Then
                Set elements(index) = item
                result = True
        Else
                elements(index) = item
                result = True
        End If
        
Finally:
        setItem = result
End Function

'アイテム削除
Public Function removeItem(index As Long) As Boolean
    Dim result As Boolean
    ReDim tmpElements(thisSize - 2) As Variant
    result = False
    
    Dim i As Long
    Dim j As Long
    For i = 0 To thisSize - 1
        If index = i Then
            result = True
        Else
            tmpElements(j) = elements(i)
            j = j + 1
        End If
    Next
    
    If result = True Then
        ReDim Preserve elements(thisSize - 2)
        For i = 0 To UBound(tmpElements)
            elements(i) = tmpElements(i)
        Next
        thisSize = thisSize - 1
    End If
    removeItem = result
End Function

'同じアイテムがあるかどうかの確認
Public Function containsValue(item As Variant) As Boolean
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

'サイズ取得
Public Function size() As Long
    size = thisSize
End Function
