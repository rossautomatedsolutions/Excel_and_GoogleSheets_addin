Attribute VB_Name = "mod_formatting"
Option Explicit

'====================================================================================
' Module: mod_formatting
' Purpose:
'   Utility routines for common formatting operations in Excel.
'   All routines avoid Select/Activate patterns and operate directly on objects.
'====================================================================================

' Clears all formatting from the currently active sheet's used range.
Public Sub ClearFormattingActiveSheet()
    Dim targetSheet As Worksheet
    Dim usedData As Range
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveSheet Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    Set targetSheet = ActiveSheet

    ' Exit safely when the worksheet has no used range.
    If IsWorksheetEmpty(targetSheet) Then Exit Sub

    Set usedData = targetSheet.UsedRange

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit
    usedData.ClearFormats

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Clears all formatting from every worksheet in the active workbook.
Public Sub ClearFormattingWorkbook()
    Dim book As Workbook
    Dim ws As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveWorkbook Is Nothing Then Exit Sub
    Set book = ActiveWorkbook

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit

    For Each ws In book.Worksheets
        If Not IsWorksheetEmpty(ws) Then
            ws.UsedRange.ClearFormats
        End If
    Next ws

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Formats the selected range as number with thousands separators and no decimals.
Public Sub FormatSelectionAsNumberNoDecimalsWithCommas()
    ApplyNumberFormatToSelection "#,##0"
End Sub

' Formats the selected range as currency with no decimals.
Public Sub FormatSelectionAsCurrencyNoDecimals()
    ApplyNumberFormatToSelection "$#,##0"
End Sub

' Formats the selected range as percent with no decimals.
Public Sub FormatSelectionAsPercentNoDecimals()
    ApplyNumberFormatToSelection "0%"
End Sub

' Applies bold formatting to the first row of the active worksheet's used range.
Public Sub BoldFirstRow()
    Dim targetSheet As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveSheet Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    Set targetSheet = ActiveSheet

    ' Exit safely if there is no content to format.
    If IsWorksheetEmpty(targetSheet) Then Exit Sub

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit
    targetSheet.UsedRange.Rows(1).Font.Bold = True

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Freezes the first row in the active window while preserving the current selection.
Public Sub FreezeFirstRow()
    Dim wnd As Window
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveWindow Is Nothing Then Exit Sub
    Set wnd = ActiveWindow

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit

    ' Configure splits directly without using Select.
    wnd.SplitColumn = 0
    wnd.SplitRow = 1
    wnd.FreezePanes = True

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Auto-fits every column in the active worksheet's used range.
Public Sub AutoFitAllColumns()
    Dim targetSheet As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveSheet Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    Set targetSheet = ActiveSheet
    If IsWorksheetEmpty(targetSheet) Then Exit Sub

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit
    targetSheet.UsedRange.Columns.AutoFit

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Removes all conditional formatting rules from the active worksheet.
Public Sub RemoveAllConditionalFormatting()
    Dim targetSheet As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    If ActiveSheet Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    Set targetSheet = ActiveSheet

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit
    targetSheet.Cells.FormatConditions.Delete

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

'----------------------------------
' Internal helper routines
'----------------------------------

' Applies a number format string to the current selection when it is a range.
Private Sub ApplyNumberFormatToSelection(ByVal formatMask As String)
    Dim selectedCells As Range
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    Set selectedCells = GetSelectedRange()
    If selectedCells Is Nothing Then Exit Sub

    StartPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
    On Error GoTo CleanExit
    selectedCells.NumberFormat = formatMask

CleanExit:
    EndPerformanceMode prevScreenUpdating, prevEnableEvents, prevCalculation
End Sub

' Returns the current selection as a range; Nothing when selection is absent or non-range.
Private Function GetSelectedRange() As Range
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Set GetSelectedRange = Selection
    End If
    On Error GoTo 0
End Function

' Returns True when a worksheet has no meaningful used cells.
Private Function IsWorksheetEmpty(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then
        IsWorksheetEmpty = True
    Else
        IsWorksheetEmpty = (Application.WorksheetFunction.CountA(ws.Cells) = 0)
    End If
End Function

' Captures current application settings and applies performance-oriented values.
Private Sub StartPerformanceMode(ByRef prevScreenUpdating As Boolean, _
                                 ByRef prevEnableEvents As Boolean, _
                                 ByRef prevCalculation As XlCalculation)
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' Restores application settings after a formatting operation.
Private Sub EndPerformanceMode(ByVal prevScreenUpdating As Boolean, _
                               ByVal prevEnableEvents As Boolean, _
                               ByVal prevCalculation As XlCalculation)
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
End Sub
