Attribute VB_Name = "Utility"
'/**
' * 機能概要：ハッシュコードを取得する
' * 引数１：文字列
' * 戻り値：long型のハッシュ
' */
Public Function hashCode(str As String)
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

    hashCode = hash
End Function

'/**
' * 機能概要：シート名にエスケープする
' * 引数１：文字列
' * 引数２：最大長（任意）
' * 引数３：エスケープ後の文字
' * 戻り値：エスケープされた文字列
' */
Public Function escapeSheetName(ByVal str As String, Optional ByVal maxLen As Long = 30, Optional ByVal replaced As String = "〓")
    str = replace(str, "\", replaced)
    str = replace(str, "/", replaced)
    str = replace(str, ":", replaced)
    str = replace(str, "*", replaced)
    str = replace(str, "?", replaced)
    str = replace(str, """", replaced)
    str = replace(str, "<", replaced)
    str = replace(str, ">", replaced)
    str = replace(str, "|", replaced)
    str = replace(str, "[", replaced)
    str = replace(str, "]", replaced)
    If Len(str) > maxLen And maxLen > 0 Then
        str = Left(str, maxLen - 1) & "…"
    End If
    escapeSheetName = str
End Function

'/**
' * 機能概要：指定した範囲内の文字列を連結する
' * 引数１：範囲
' * 引数２：区切り文字
' * 引数３：囲い文字
' * 戻り値：連結された文字列
' */
Public Function concat(rng As range, Optional separator As String = "", Optional quot As String = "")
    Dim result As String
    Dim maxRow As Long
    Dim maxCol As Long
    result = ""
    maxRow = rng.Cells.Rows.count
    maxCol = rng.Cells.Columns.count

    For i = 1 To maxRow
        For j = 1 To maxCol
            result = result + quot + rng.Cells(i, j).Text + quot
            If i <> maxRow Or j <> maxCol Then
                result = result + separator
            End If
        Next
    Next
    
    concat = result
End Function
