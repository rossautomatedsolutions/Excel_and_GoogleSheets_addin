Attribute VB_Name = "mod_test_template_generator"
Option Explicit

' =============================================================================
' Module: mod_test_template_generator
' Purpose: Build intentionally messy test workbooks for Ross Spreadsheet Utilities.
' =============================================================================

Public Sub GenerateRossUtilitiesTestWorkbook(Optional largeMode As Boolean = False)
    Const PROC_NAME As String = "GenerateRossUtilitiesTestWorkbook"

    Dim wb As Workbook
    Dim calcMode As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim rowCount As Long

    On Error GoTo ErrHandler

    rowCount = IIf(largeMode, 50000, 10000)

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts
    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Generating Ross Utilities test workbook..."

    Randomize Timer

    Set wb = Workbooks.Add(xlWBATWorksheet)

    With wb.Worksheets(1)
        .Name = "Messy_Data"
        BuildMessyDataSheet .Parent.Worksheets("Messy_Data"), rowCount
    End With

    BuildFormattingChaosSheet AddNamedSheet(wb, "Formatting_Chaos")
    BuildNavigationTestSheet AddNamedSheet(wb, "Navigation_Test")
    BuildDuplicateTestSheet AddNamedSheet(wb, "Duplicate_Test")
    BuildEmptyStructuralSheet AddNamedSheet(wb, "Empty_Structural_Test")

    SaveWorkbookToDesktopIfRequested wb

    Application.StatusBar = False
    MsgBox "Ross Utilities test workbook created successfully." & vbCrLf & _
           "Rows in Messy_Data: " & Format$(rowCount, "#,##0"), vbInformation, "Workbook Created"

CleanExit:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.Calculation = calcMode
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "Error in " & PROC_NAME & ": " & Err.Description, vbExclamation, "Test Workbook Generator"
    Resume CleanExit
End Sub

Private Function AddNamedSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Set AddNamedSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    AddNamedSheet.Name = sheetName
End Function

Private Sub BuildMessyDataSheet(ByVal ws As Worksheet, ByVal dataRows As Long)
    Dim arr() As Variant
    Dim headers As Variant
    Dim r As Long
    Dim c As Long
    Dim duplicateSourceRow As Long

    headers = Array("  Customer ID  ", "amount", "Amount", " %Complete ", "DATE entered", _
                    "DATE entered", "Notes!@#", "city ", "  CITY", "Code-#")

    ReDim arr(1 To dataRows + 1, 1 To 10)

    For c = 1 To 10
        arr(1, c) = headers(c - 1)
    Next c

    For r = 2 To dataRows + 1
        If r Mod 223 = 0 Then
            ' Intentional random blank row.
            For c = 1 To 10
                arr(r, c) = vbNullString
            Next c
        ElseIf r Mod 251 = 0 Then
            ' Intentional duplicate row.
            duplicateSourceRow = r - 1
            For c = 1 To 10
                arr(r, c) = arr(duplicateSourceRow, c)
            Next c
        Else
            arr(r, 1) = "'" & CStr(Int((999999 - 100000 + 1) * Rnd + 100000))                      ' Numeric-like text
            arr(r, 2) = "$" & Format$(Int(500000 * Rnd) + (Rnd * 100), "#,##0.00")               ' Currency as text
            arr(r, 3) = CStr(Int(10000 * Rnd) + 1)                                                   ' Numeric stored as text-like mix
            arr(r, 4) = CStr(Int(100 * Rnd)) & "%"                                                  ' Percent string
            arr(r, 5) = Format$(DateSerial(2019 + Int(8 * Rnd), 1 + Int(12 * Rnd), 1 + Int(28 * Rnd)), "mm/dd/yyyy")
            arr(r, 6) = IIf(Rnd < 0.2, vbNullString, Format$(Date - Int(2000 * Rnd), "yyyy-mm-dd")) ' Date text with blanks
            arr(r, 7) = RandomPersonName() & String$(Int(4 * Rnd), " ")                            ' Trailing spaces
            arr(r, 8) = RandomCityName() & IIf(Rnd < 0.35, "  ", vbNullString)                     ' Trailing spaces
            arr(r, 9) = IIf(Rnd < 0.15, vbNullString, LCase$(RandomCityName()))                      ' Mixed case/blank
            arr(r, 10) = IIf(Rnd < 0.5, "X-" & CStr(Int(999 * Rnd)), Int(99 * Rnd))                ' Mixed types

            If Rnd < 0.06 Then arr(r, 7) = vbNullString
            If Rnd < 0.04 Then arr(r, 2) = vbNullString
        End If
    Next r

    With ws
        .Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        .Rows(1).Font.Bold = True
        .Columns("A:J").HorizontalAlignment = xlLeft
        .Columns("A:J").AutoFit
        AddOrReplaceComment .Range("A1"), _
            "Intentionally messy headers, duplicate columns, mixed data types, blanks, text-formatted numerics, and duplicate rows for cleanup testing."
    End With
End Sub

Private Sub BuildFormattingChaosSheet(ByVal ws As Worksheet)
    Dim arr() As Variant
    Dim r As Long
    Dim c As Long
    Dim target As Range

    ReDim arr(1 To 220, 1 To 15)

    For r = 1 To 220
        For c = 1 To 15
            arr(r, c) = "Cell_" & r & "_" & c
        Next c
    Next r

    With ws
        .Range("A1").Resize(220, 15).Value = arr

        For c = 1 To 15
            .Columns(c).ColumnWidth = 6 + Int(25 * Rnd)  ' Random column widths
        Next c

        For r = 1 To 220
            If Rnd < 0.08 Then .Rows(r).Hidden = True     ' Hidden rows
        Next r

        For c = 1 To 15
            If Rnd < 0.2 Then .Columns(c).Hidden = True   ' Hidden columns
        Next c

        For Each target In .Range("A1:O220")
            If Rnd < 0.18 Then
                target.Font.Name = Choose(Int(5 * Rnd) + 1, "Calibri", "Arial", "Tahoma", "Verdana", "Times New Roman")
                target.Font.Size = 8 + Int(12 * Rnd)
                target.Interior.Color = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
            End If
        Next target

        ' Random conditional formatting rules.
        .Range("A1:O220").FormatConditions.Delete
        .Range("A1:O220").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=\"Cell_150_10\""
        .Range("A1:O220").FormatConditions(1).Interior.Color = RGB(255, 200, 200)

        .Range("C1:C220").FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()=COLUMN()*3"
        .Range("C1:C220").FormatConditions(.Range("C1:C220").FormatConditions.Count).Font.Bold = True

        .Range("E1:E220").FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),7)=0"
        .Range("E1:E220").FormatConditions(.Range("E1:E220").FormatConditions.Count).Interior.Color = RGB(200, 255, 200)

        ' Merged cells to create structural chaos.
        .Range("B4:D4").Merge
        .Range("F8:H9").Merge
        .Range("J12:K13").Merge

        AddOrReplaceComment .Range("A1"), _
            "Intentionally chaotic formatting: random fonts/sizes/colors, hidden rows/columns, conditional formats, random widths, and merged cells."
    End With
End Sub

Private Sub BuildNavigationTestSheet(ByVal ws As Worksheet)
    Dim arr() As Variant
    Dim r As Long

    ReDim arr(1 To 5001, 1 To 6)

    arr(1, 1) = "RecordID"
    arr(1, 2) = "Region"
    arr(1, 3) = "Category"
    arr(1, 4) = "Status"
    arr(1, 5) = "Amount"
    arr(1, 6) = "DateText"

    For r = 2 To 5001
        arr(r, 1) = r - 1
        arr(r, 2) = Choose((r Mod 4) + 1, "North", "South", "East", "West")
        arr(r, 3) = Choose((r Mod 5) + 1, "A", "B", "C", "D", "E")
        arr(r, 4) = Choose((r Mod 3) + 1, "Open", "Closed", "Pending")
        arr(r, 5) = Int(5000 * Rnd)
        arr(r, 6) = Format$(Date - Int(365 * Rnd), "dd-mmm-yyyy")
    Next r

    With ws
        .Range("A1").Resize(5001, 6).Value = arr
        .Rows(1).Font.Bold = True
        .Range("A1:F5001").AutoFilter

        ' Multiple filter states to make navigation/filter reset testing harder.
        .Range("A1:F5001").AutoFilter Field:=2, Criteria1:="<>West"
        .Range("A1:F5001").AutoFilter Field:=4, Criteria1:=Array("Open", "Pending"), Operator:=xlFilterValues

        AddOrReplaceComment .Range("A1"), _
            "Navigation test: large filtered dataset, awkward freeze pane location, and active cell moved far below visible top."
    End With

    ' Freeze panes intentionally in an awkward location.
    ws.Activate
    ws.Parent.Windows(1).FreezePanes = False
    ws.Range("C5").Select
    ws.Parent.Windows(1).FreezePanes = True

    ' Leave active cell far down sheet.
    ws.Range("F4200").Select
End Sub

Private Sub BuildDuplicateTestSheet(ByVal ws As Worksheet)
    Dim arr() As Variant
    Dim r As Long

    ReDim arr(1 To 2001, 1 To 8)

    arr(1, 1) = "OrderID"
    arr(1, 2) = "Customer"
    arr(1, 3) = "SKU"
    arr(1, 4) = "Qty"
    arr(1, 5) = ""
    arr(1, 6) = "Amount"
    arr(1, 7) = "Status"
    arr(1, 8) = "MixedType"

    For r = 2 To 2001
        arr(r, 1) = 100000 + r
        arr(r, 2) = "Customer_" & ((r Mod 110) + 1)
        arr(r, 3) = "SKU-" & ((r Mod 45) + 1)
        arr(r, 4) = (r Mod 12) + 1
        arr(r, 5) = vbNullString                                ' Blank middle column
        arr(r, 6) = Round((r Mod 17) * 19.95, 2)
        arr(r, 7) = Choose((r Mod 3) + 1, "New", "Shipped", "Held")
        arr(r, 8) = IIf(r Mod 2 = 0, CStr((r Mod 500)), (r Mod 500))

        If r Mod 120 = 0 Then
            ' Entire duplicated row.
            arr(r, 1) = arr(r - 1, 1)
            arr(r, 2) = arr(r - 1, 2)
            arr(r, 3) = arr(r - 1, 3)
            arr(r, 4) = arr(r - 1, 4)
            arr(r, 6) = arr(r - 1, 6)
            arr(r, 7) = arr(r - 1, 7)
            arr(r, 8) = arr(r - 1, 8)
        ElseIf r Mod 55 = 0 Then
            ' Partial duplicates (same keys, changed payload).
            arr(r, 1) = arr(r - 2, 1)
            arr(r, 2) = arr(r - 2, 2)
            arr(r, 3) = arr(r - 2, 3)
        End If
    Next r

    With ws
        .Range("A1").Resize(2001, 8).Value = arr
        .Rows(1).Font.Bold = True
        .Columns("A:H").AutoFit
        AddOrReplaceComment .Range("A1"), _
            "Duplicate stress test: full-row duplicates, partial duplicates, mixed types, and an intentionally blank middle column."
    End With
End Sub

Private Sub BuildEmptyStructuralSheet(ByVal ws As Worksheet)
    Dim arr() As Variant
    Dim r As Long

    ReDim arr(1 To 400, 1 To 18)

    arr(1, 1) = "Section"
    arr(1, 4) = "SparseValue"
    arr(1, 10) = "Notes"
    arr(1, 16) = "Flag"

    For r = 2 To 400
        If r Mod 40 = 0 Or r Mod 41 = 0 Then
            ' Intentionally leave completely empty rows in middle blocks.
        Else
            arr(r, 1) = "Block_" & ((r - 2) \ 25 + 1)
            arr(r, 4) = Int(1000 * Rnd)
            arr(r, 10) = IIf(Rnd < 0.7, vbNullString, "Sparse note " & r)
            arr(r, 16) = IIf(Rnd < 0.5, "Y", vbNullString)
        End If
    Next r

    With ws
        .Range("A1").Resize(400, 18).Value = arr
        .Columns("B:C").EntireColumn.ClearContents    ' Completely empty columns
        .Columns("F:H").EntireColumn.ClearContents    ' More empty columns
        .Columns("K:N").EntireColumn.ClearContents
        .Columns("A:R").AutoFit

        AddOrReplaceComment .Range("A1"), _
            "Structural emptiness test: sparse islands of data, fully empty columns, and intentionally empty row bands."
    End With
End Sub

Private Sub SaveWorkbookToDesktopIfRequested(ByVal wb As Workbook)
    Dim saveChoice As VbMsgBoxResult
    Dim desktopPath As String
    Dim fullPath As String

    saveChoice = MsgBox("Save workbook to Desktop as 'Ross_Utilities_Test_Workbook.xlsx'?", _
                        vbYesNo + vbQuestion, "Save Test Workbook")

    If saveChoice = vbYes Then
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        fullPath = desktopPath & Application.PathSeparator & "Ross_Utilities_Test_Workbook.xlsx"

        wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    End If
End Sub

Private Sub AddOrReplaceComment(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    targetCell.ClearComments
    On Error GoTo 0
    targetCell.AddComment commentText
End Sub

Private Function RandomPersonName() As String
    RandomPersonName = Choose(Int(10 * Rnd) + 1, _
                              "Alex", "Jordan", "Casey", "Taylor", "Morgan", _
                              "Riley", "Avery", "Parker", "Quinn", "Drew")
End Function

Private Function RandomCityName() As String
    RandomCityName = Choose(Int(10 * Rnd) + 1, _
                            "Seattle", "Austin", "Boston", "Denver", "Phoenix", _
                            "Miami", "Chicago", "Dallas", "Atlanta", "Portland")
End Function
