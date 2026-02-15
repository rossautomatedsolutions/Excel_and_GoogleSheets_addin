Attribute VB_Name = "mod_data_utilities"
Option Explicit

' Removes rows that have no values at all across the current UsedRange.
Public Sub RemoveCompletelyEmptyRows()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim hasValue As Boolean
    Dim blockEnd As Long
    Dim removedRows As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub
    If ur.Rows.Count = 1 And ur.Columns.Count = 1 Then
        If LenB(CStr(ur.Value2)) = 0 Then
            Debug.Print "RemoveCompletelyEmptyRows affected rows: 0"
            Exit Sub
        End If
    End If

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Scan bottom-up so row indexes remain valid while deleting contiguous blocks.
    r = rowCount
    Do While r >= 1
        hasValue = False
        For c = 1 To colCount
            If Not IsEmpty(data(r, c)) Then
                If LenB(CStr(data(r, c))) > 0 Then
                    hasValue = True
                    Exit For
                End If
            End If
        Next c

        If Not hasValue Then
            blockEnd = r
            Do While r >= 1
                hasValue = False
                For c = 1 To colCount
                    If Not IsEmpty(data(r, c)) Then
                        If LenB(CStr(data(r, c))) > 0 Then
                            hasValue = True
                            Exit For
                        End If
                    End If
                Next c

                If hasValue Then Exit Do
                r = r - 1
            Loop

            ' Delete one contiguous empty-row block in a single operation.
            ws.Range( _
                ws.Cells(ur.Row + r, ur.Column), _
                ws.Cells(ur.Row + blockEnd - 1, ur.Column + colCount - 1) _
            ).Delete Shift:=xlShiftUp

            removedRows = removedRows + (blockEnd - r)
        Else
            r = r - 1
        End If
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Debug.Print "RemoveCompletelyEmptyRows affected rows: " & removedRows
End Sub

' Removes columns that have no values at all across the current UsedRange.
Public Sub RemoveCompletelyEmptyColumns()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim hasValue As Boolean
    Dim blockEnd As Long
    Dim removedCols As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Scan right-to-left so column indexes remain valid while deleting contiguous blocks.
    c = colCount
    Do While c >= 1
        hasValue = False
        For r = 1 To rowCount
            If Not IsEmpty(data(r, c)) Then
                If LenB(CStr(data(r, c))) > 0 Then
                    hasValue = True
                    Exit For
                End If
            End If
        Next r

        If Not hasValue Then
            blockEnd = c
            Do While c >= 1
                hasValue = False
                For r = 1 To rowCount
                    If Not IsEmpty(data(r, c)) Then
                        If LenB(CStr(data(r, c))) > 0 Then
                            hasValue = True
                            Exit For
                        End If
                    End If
                Next r

                If hasValue Then Exit Do
                c = c - 1
            Loop

            ' Delete one contiguous empty-column block in a single operation.
            ws.Range( _
                ws.Cells(ur.Row, ur.Column + c), _
                ws.Cells(ur.Row + rowCount - 1, ur.Column + blockEnd - 1) _
            ).Delete Shift:=xlToLeft

            removedCols = removedCols + (blockEnd - c)
        Else
            c = c - 1
        End If
    Loop

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Debug.Print "RemoveCompletelyEmptyColumns affected rows: " & removedCols
End Sub

' Trims leading/trailing whitespace in all text cells inside UsedRange.
Public Sub TrimWhitespaceEntireSheet()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim beforeText As String, afterText As String
    Dim changedCount As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    ' Modify in-memory array first, then write back once for performance.
    For r = 1 To rowCount
        For c = 1 To colCount
            If VarType(data(r, c)) = vbString Then
                beforeText = data(r, c)
                afterText = Trim$(beforeText)
                If StrComp(beforeText, afterText, vbBinaryCompare) <> 0 Then
                    data(r, c) = afterText
                    changedCount = changedCount + 1
                End If
            End If
        Next c
    Next r

    If changedCount > 0 Then ur.Value2 = data

    Debug.Print "TrimWhitespaceEntireSheet affected rows: " & changedCount
End Sub

' Converts text values that are numeric-looking into numeric values inside UsedRange.
Public Sub ConvertTextNumbersToNumeric()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim rawText As String, trimmedText As String
    Dim changedCount As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)

    ' Convert in memory to avoid expensive per-cell worksheet writes.
    For r = 1 To rowCount
        For c = 1 To colCount
            If VarType(data(r, c)) = vbString Then
                rawText = data(r, c)
                trimmedText = Trim$(rawText)

                If LenB(trimmedText) > 0 And IsNumeric(trimmedText) Then
                    data(r, c) = CDbl(trimmedText)
                    changedCount = changedCount + 1
                End If
            End If
        Next c
    Next r

    If changedCount > 0 Then ur.Value2 = data

    Debug.Print "ConvertTextNumbersToNumeric affected rows: " & changedCount
End Sub

' Highlights every row in UsedRange whose full row content appears more than once.
Public Sub HighlightDuplicateRows()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long
    Dim rowKey As String
    Dim counts As Object
    Dim duplicateRows As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    Set counts = CreateObject("Scripting.Dictionary")

    ' First pass: count each distinct row signature.
    For r = 1 To rowCount
        rowKey = BuildRowKey(data, r, colCount)
        If counts.Exists(rowKey) Then
            counts(rowKey) = counts(rowKey) + 1
        Else
            counts.Add rowKey, 1
        End If
    Next r

    ' Clear prior fill inside UsedRange then apply highlight to duplicate rows.
    ur.Interior.Pattern = xlNone

    For r = 1 To rowCount
        rowKey = BuildRowKey(data, r, colCount)
        If counts(rowKey) > 1 Then
            ws.Range( _
                ws.Cells(ur.Row + r - 1, ur.Column), _
                ws.Cells(ur.Row + r - 1, ur.Column + colCount - 1) _
            ).Interior.Color = RGB(255, 242, 204)
            duplicateRows = duplicateRows + 1
        End If
    Next r

    Debug.Print "HighlightDuplicateRows affected rows: " & duplicateRows
End Sub

' Removes duplicate rows (keeps first occurrence) inside UsedRange without row-by-row deletion.
Public Sub RemoveDuplicateRows()
    Dim ws As Worksheet
    Dim ur As Range
    Dim data As Variant
    Dim outputData() As Variant
    Dim rowCount As Long, colCount As Long
    Dim r As Long, c As Long
    Dim writeRow As Long
    Dim key As String
    Dim seen As Object
    Dim removedRows As Long

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur Is Nothing Then Exit Sub

    data = GetRangeValues2D(ur)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    Set seen = CreateObject("Scripting.Dictionary")

    ' First pass: compute how many rows we keep.
    For r = 1 To rowCount
        key = BuildRowKey(data, r, colCount)
        If Not seen.Exists(key) Then seen.Add key, True
    Next r

    ReDim outputData(1 To seen.Count, 1 To colCount)

    ' Second pass: copy only first occurrences into compact output array.
    seen.RemoveAll
    writeRow = 0
    For r = 1 To rowCount
        key = BuildRowKey(data, r, colCount)
        If Not seen.Exists(key) Then
            writeRow = writeRow + 1
            seen.Add key, True
            For c = 1 To colCount
                outputData(writeRow, c) = data(r, c)
            Next c
        Else
            removedRows = removedRows + 1
        End If
    Next r

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Rewrite compacted result once, then clear the trailing old rows in one block.
    ws.Range( _
        ws.Cells(ur.Row, ur.Column), _
        ws.Cells(ur.Row + writeRow - 1, ur.Column + colCount - 1) _
    ).Value2 = outputData

    If writeRow < rowCount Then
        ws.Range( _
            ws.Cells(ur.Row + writeRow, ur.Column), _
            ws.Cells(ur.Row + rowCount - 1, ur.Column + colCount - 1) _
        ).ClearContents
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Debug.Print "RemoveDuplicateRows affected rows: " & removedRows
End Sub



' Always returns a 1-based 2D array for a range, including 1-cell UsedRange cases.
Private Function GetRangeValues2D(ByVal target As Range) As Variant
    Dim raw As Variant
    Dim temp(1 To 1, 1 To 1) As Variant

    raw = target.Value2

    If target.Rows.Count = 1 And target.Columns.Count = 1 Then
        temp(1, 1) = raw
        GetRangeValues2D = temp
    Else
        GetRangeValues2D = raw
    End If
End Function
' Serializes one row of the 2D array to a stable dictionary key.
Private Function BuildRowKey(ByRef data As Variant, ByVal rowIndex As Long, ByVal colCount As Long) As String
    Dim parts() As String
    Dim c As Long

    ReDim parts(1 To colCount)
    For c = 1 To colCount
        parts(c) = Replace$(CStr(data(rowIndex, c)), "|", "||")
    Next c

    BuildRowKey = Join(parts, "|#|")
End Function
