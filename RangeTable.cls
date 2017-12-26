VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rng_ref As Range
Private rng_cells As Variant
Private col_map As New HashMap
Private initialized As Boolean

Private Sub Class_Initialize()
    initialized = False
End Sub

'init
Public Function init(rng As Range, colNameList As ArrayList)

    'error check
    If rng.Columns.Count <> colNameList.size() Then
        Err.Raise 1000, "", "Error:rng.Columns.Count <> colNameList.size()"
    Else
        
        'initialize
        Set rng_ref = rng
        rng_cells = rng.value
        
        'columns name put to map
        Dim i As Long
        For i = 1 To colNameList.size()
            Dim tmp As String
            tmp = colNameList.getItem(i - 1)
            
            'error check
            If col_map.containsKey(tmp) Then
                Err.Raise 1000, "", "Error:col_map.containsKey(colName)"
            ElseIf rng.Cells(1, i).Text <> tmp Then
                Err.Raise 1000, "", "rng.Cells(1, i).Text <> tmp"
            Else
                Call col_map.putItem(tmp, i)
            End If
        Next
    End If

    initialized = True
End Function

'select
Public Function selectValue(target As String, refColName As String, resultColName As String, Optional errorAlert As Boolean)
        Dim resultList As ArrayList
        Set resultList = selectValues(target, refColName, resultColName, True)
        
        'error check
        If resultList.size() = 0 Then
                If errorAlert = True Then
                        Err.Raise 1000, "", "resultList.size() = 0"
                Else
                        Call resultList.addItem("")
                End If
        End If
        
        selectValue = resultList.getItem(0)
End Function

'selectAll
Public Function selectValues(target As String, refColName As String, resultColName As String, Optional unique As Boolean)
    If initialized = False Then
        Err.Raise 1000, "", "Error:initialized = False"
    End If
    
    'make tmp
    Dim resultList As New ArrayList
    Dim ref_col As Long
    Dim result_col As String
    ref_col = col_map.getItem(refColName)
    result_col = col_map.getItem(resultColName)
    
    'search then add to list
    Dim i As Long
    For i = 1 To rng_ref.Rows.Count
        If rng_cells(i, ref_col) = target Then
            Call resultList.addItem(rng_cells(i, result_col))
            If unique = True Then Exit For
        End If
    Next
    
    Set selectValues = resultList
End Function






