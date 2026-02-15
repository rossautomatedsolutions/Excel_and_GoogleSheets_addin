Attribute VB_Name = "mod_navigation"
Option Explicit

' ============================================================================
' Module: mod_navigation
' Purpose: Navigation and view helpers for the active worksheet.
' Notes  :
'   - Uses direct range/window operations for efficiency.
'   - Avoids activating other sheets/workbooks.
'   - Includes guards for non-worksheet context and empty sheets.
' ============================================================================

' --- Public API ---------------------------------------------------------------

Public Sub GoToA1ActiveSheet()
    ' Navigate to cell A1 of the active worksheet.
    Dim ws As Worksheet
    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Application.Goto Reference:=ws.Range("A1"), Scroll:=True
End Sub

Public Sub GoToLastUsedCell()
    ' Navigate to the last used cell (bottom-right used boundary).
    ' If the sheet is empty, safely falls back to A1.
    Dim ws As Worksheet
    Dim lastCell As Range

    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Set lastCell = FindLastUsedCell(ws)

    If lastCell Is Nothing Then
        Application.Goto Reference:=ws.Range("A1"), Scroll:=True
    Else
        Application.Goto Reference:=lastCell, Scroll:=True
    End If
End Sub

Public Sub SelectUsedRange()
    ' Select the true used range boundaries based on cell content/formulas.
    ' If the sheet has no used cells, safely selects A1.
    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Set targetRange = FindTrueUsedRange(ws)

    If targetRange Is Nothing Then
        ws.Range("A1").Select
    Else
        targetRange.Select
    End If
End Sub

Public Sub ClearAllFiltersActiveSheet()
    ' Clear all active filters on worksheet-level ranges and tables.
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    ' 1) Clear table filters (ListObjects) only where a filter is active.
    For Each lo In ws.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then
                lo.AutoFilter.ShowAllData
            End If
        End If
    Next lo

    ' 2) Clear worksheet AutoFilter (range filter) when active.
    If ws.FilterMode Then
        ws.ShowAllData
    End If
End Sub

Public Sub FreezeFirstColumn()
    ' Freeze the first column for the active window.
    ' Uses window split settings directly (no cursor movement required).
    Dim ws As Worksheet

    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub
    If ActiveWindow Is Nothing Then Exit Sub

    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 0
        .FreezePanes = True
    End With
End Sub

' --- Internal helpers ---------------------------------------------------------

Private Function GetActiveWorksheet() As Worksheet
    ' Returns ActiveSheet only when it is a worksheet; otherwise Nothing.
    If TypeOf ActiveSheet Is Worksheet Then
        Set GetActiveWorksheet = ActiveSheet
    End If
End Function

Private Function FindLastUsedCell(ByVal ws As Worksheet) As Range
    ' Finds the last used cell by searching formulas/values.
    Set FindLastUsedCell = ws.Cells.Find(What:="*", _
                                         After:=ws.Cells(1, 1), _
                                         LookIn:=xlFormulas, _
                                         LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, _
                                         SearchDirection:=xlPrevious, _
                                         MatchCase:=False)
End Function

Private Function FindTrueUsedRange(ByVal ws As Worksheet) As Range
    ' Computes the used-range rectangle from first/last used row/column.
    Dim firstRowCell As Range
    Dim firstColCell As Range
    Dim lastRowCell As Range
    Dim lastColCell As Range

    Set lastRowCell = ws.Cells.Find(What:="*", _
                                    After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    If lastRowCell Is Nothing Then Exit Function

    Set firstRowCell = ws.Cells.Find(What:="*", _
                                     After:=ws.Cells(lastRowCell.Row, lastRowCell.Column), _
                                     LookIn:=xlFormulas, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False)

    Set lastColCell = ws.Cells.Find(What:="*", _
                                    After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    Set firstColCell = ws.Cells.Find(What:="*", _
                                     After:=ws.Cells(lastColCell.Row, lastColCell.Column), _
                                     LookIn:=xlFormulas, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByColumns, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False)

    Set FindTrueUsedRange = ws.Range(ws.Cells(firstRowCell.Row, firstColCell.Column), _
                                     ws.Cells(lastRowCell.Row, lastColCell.Column))
End Function
