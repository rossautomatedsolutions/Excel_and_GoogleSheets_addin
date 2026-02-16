Attribute VB_Name = "mod_sheet_management"
Option Explicit

'====================================================================================
' Module: mod_sheet_management
' Purpose: High-performance worksheet/tab management routines for large workbooks.
'====================================================================================

Private Type TApplicationState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Public Sub UnhideAllSheets()
    ' Unhides every worksheet in the active workbook (including very hidden sheets).
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim appState As TApplicationState

    On Error GoTo CleanFail

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 1000, "UnhideAllSheets", "No active workbook is available."

    ' Workbook structure protection prevents sheet visibility changes.
    If wb.ProtectStructure Then
        MsgBox "Cannot unhide sheets because workbook structure is protected.", vbInformation, "Unhide All Sheets"
        Exit Sub
    End If

    BeginBatch appState

    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Next ws

CleanExit:
    EndBatch appState
    Exit Sub

CleanFail:
    MsgBox "UnhideAllSheets failed: " & Err.Description, vbExclamation, "Sheet Management"
    Resume CleanExit
End Sub

Public Sub HideAllExceptCurrent()
    ' Hides all worksheets except the currently active worksheet.
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim visibleCount As Long
    Dim appState As TApplicationState

    On Error GoTo CleanFail

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Err.Raise vbObjectError + 1001, "HideAllExceptCurrent", "Active sheet is not a worksheet."
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 1002, "HideAllExceptCurrent", "No active workbook is available."

    If wb.ProtectStructure Then
        MsgBox "Cannot hide sheets because workbook structure is protected.", vbInformation, "Hide All Except Current"
        Exit Sub
    End If

    Set currentSheet = ActiveSheet

    ' Ensure we never attempt to hide the only visible worksheet.
    visibleCount = CountVisibleWorksheets(wb)
    If visibleCount <= 1 Then Exit Sub

    BeginBatch appState

    For Each ws In wb.Worksheets
        If ws.Name <> currentSheet.Name And ws.Visible = xlSheetVisible Then
            ws.Visible = xlSheetHidden
        End If
    Next ws

CleanExit:
    EndBatch appState
    Exit Sub

CleanFail:
    MsgBox "HideAllExceptCurrent failed: " & Err.Description, vbExclamation, "Sheet Management"
    Resume CleanExit
End Sub

Public Sub GoToA1AllSheetsThenReturn()
    ' Visits each visible worksheet, positions the active cell/scroll at A1,
    ' then restores the user's original sheet/cell context.
    ' Note: Activate/GoTo is required here because Excel stores active cell/scroll per sheet.
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim originSheet As Worksheet
    Dim originCellAddress As String
    Dim appState As TApplicationState

    On Error GoTo CleanFail

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 1003, "GoToA1AllSheetsThenReturn", "No active workbook is available."

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Err.Raise vbObjectError + 1004, "GoToA1AllSheetsThenReturn", "Active sheet is not a worksheet."
    End If

    Set originSheet = ActiveSheet
    originCellAddress = ActiveCell.Address(False, False)

    BeginBatch appState

    For Each ws In wb.Worksheets
        ' Hidden sheets cannot be activated; skip them gracefully.
        If ws.Visible = xlSheetVisible Then
            Application.Goto ws.Range("A1"), True
        End If
    Next ws

    ' Restore original user context.
    Application.Goto originSheet.Range(originCellAddress), True

CleanExit:
    EndBatch appState
    Exit Sub

CleanFail:
    MsgBox "GoToA1AllSheetsThenReturn failed: " & Err.Description, vbExclamation, "Sheet Management"
    Resume CleanExit
End Sub

Public Sub AlphabetizeSheetTabs()
    ' Alphabetizes all sheet tabs (worksheets and chart sheets) by name.
    Dim wb As Workbook
    Dim appState As TApplicationState
    Dim sheetNames() As String
    Dim i As Long
    Dim countSheets As Long

    On Error GoTo CleanFail

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 1005, "AlphabetizeSheetTabs", "No active workbook is available."

    If wb.ProtectStructure Then
        MsgBox "Cannot reorder sheet tabs because workbook structure is protected.", vbInformation, "Alphabetize Sheet Tabs"
        Exit Sub
    End If

    countSheets = wb.Sheets.Count
    If countSheets <= 1 Then Exit Sub

    ReDim sheetNames(1 To countSheets)

    ' Read all names into memory first for better performance on large workbooks.
    For i = 1 To countSheets
        sheetNames(i) = wb.Sheets(i).Name
    Next i

    QuickSortStrings sheetNames, LBound(sheetNames), UBound(sheetNames)

    BeginBatch appState

    ' Move in reverse before index 1 to preserve ascending final order.
    For i = UBound(sheetNames) To LBound(sheetNames) Step -1
        wb.Sheets(sheetNames(i)).Move Before:=wb.Sheets(1)
    Next i

CleanExit:
    EndBatch appState
    Exit Sub

CleanFail:
    MsgBox "AlphabetizeSheetTabs failed: " & Err.Description, vbExclamation, "Sheet Management"
    Resume CleanExit
End Sub

Public Sub DuplicateSheetWithTimestamp()
    ' Duplicates the active worksheet and appends a timestamped suffix to the new sheet name.
    Dim wb As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim newName As String
    Dim appState As TApplicationState

    On Error GoTo CleanFail

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Err.Raise vbObjectError + 1006, "DuplicateSheetWithTimestamp", "Active sheet is not a worksheet."
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 1007, "DuplicateSheetWithTimestamp", "No active workbook is available."

    If wb.ProtectStructure Then
        MsgBox "Cannot duplicate sheet because workbook structure is protected.", vbInformation, "Duplicate Sheet"
        Exit Sub
    End If

    Set sourceSheet = ActiveSheet

    BeginBatch appState

    ' Copy directly after source sheet to preserve workbook flow.
    sourceSheet.Copy After:=sourceSheet
    Set newSheet = sourceSheet.Next

    newName = BuildTimestampedSheetName(wb, sourceSheet.Name)
    newSheet.Name = newName

CleanExit:
    EndBatch appState
    Exit Sub

CleanFail:
    MsgBox "DuplicateSheetWithTimestamp failed: " & Err.Description, vbExclamation, "Sheet Management"
    Resume CleanExit
End Sub

Private Function CountVisibleWorksheets(ByVal wb As Workbook) As Long
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            CountVisibleWorksheets = CountVisibleWorksheets + 1
        End If
    Next ws
End Function

Private Function BuildTimestampedSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    ' Builds a valid, unique worksheet name up to Excel's 31-character limit.
    Dim stamp As String
    Dim maxBaseLen As Long
    Dim candidate As String
    Dim counter As Long
    Dim suffix As String
    Dim dynamicBaseLen As Long

    stamp = "_" & Format$(Now, "yyyymmdd_hhnnss")

    maxBaseLen = 31 - Len(stamp)
    If maxBaseLen < 1 Then maxBaseLen = 1

    candidate = Left$(baseName, maxBaseLen) & stamp

    counter = 1
    Do While SheetNameExists(wb, candidate)
        counter = counter + 1
        suffix = "_" & CStr(counter)
        dynamicBaseLen = 31 - Len(stamp) - Len(suffix)
        If dynamicBaseLen < 1 Then dynamicBaseLen = 1

        candidate = Left$(baseName, dynamicBaseLen) & stamp & suffix
    Loop

    BuildTimestampedSheetName = candidate
End Function

Private Function SheetNameExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim obj As Object

    On Error Resume Next
    Set obj = wb.Sheets(sheetName)
    SheetNameExists = Not obj Is Nothing
    Set obj = Nothing
    On Error GoTo 0
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    ' In-place quicksort for efficient ordering of large sheet-name arrays.
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop

        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

Private Sub BeginBatch(ByRef state As TApplicationState)
    ' Captures application state and applies performance-focused settings.
    With Application
        state.ScreenUpdating = .ScreenUpdating
        state.EnableEvents = .EnableEvents
        state.DisplayAlerts = .DisplayAlerts
        state.Calculation = .Calculation
        state.StatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Processing sheet operation..."
    End With
End Sub

Private Sub EndBatch(ByRef state As TApplicationState)
    ' Restores application state, even when an error occurs.
    On Error Resume Next
    With Application
        .ScreenUpdating = state.ScreenUpdating
        .EnableEvents = state.EnableEvents
        .DisplayAlerts = state.DisplayAlerts
        .Calculation = state.Calculation
        .StatusBar = state.StatusBar
    End With
    On Error GoTo 0
End Sub
