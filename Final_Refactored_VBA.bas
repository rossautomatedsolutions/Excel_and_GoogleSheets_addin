Attribute VB_Name = "Final_Refactored_VBA"
Option Explicit

Private Type TAppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Private Sub PushAppState(ByRef state As TAppState)
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
        .StatusBar = "Working..."
    End With
End Sub

Private Sub PopAppState(ByRef state As TAppState)
    With Application
        .ScreenUpdating = state.ScreenUpdating
        .EnableEvents = state.EnableEvents
        .DisplayAlerts = state.DisplayAlerts
        .Calculation = state.Calculation
        .StatusBar = state.StatusBar
    End With
End Sub

Public Sub TransformData_Final()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant
    Dim r As Long
    Dim appState As TAppState

    On Error GoTo ErrHandler

    PushAppState appState

    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 1 Then GoTo SafeExit

    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    For r = 2 To UBound(data, 1)
        If LenB(data(r, 1)) <> 0 Then
            data(r, 1) = UCase$(CStr(data(r, 1)))
        End If

        If IsNumeric(data(r, 2)) Then
            data(r, 3) = CDbl(data(r, 2)) * 1.2
        End If
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2 = data

SafeExit:
    PopAppState appState
    Erase data
    Set ws = Nothing
    Exit Sub

ErrHandler:
    PopAppState appState
    MsgBox "TransformData_Final failed: " & Err.Number & " - " & Err.Description, vbCritical
    Erase data
    Set ws = Nothing
End Sub

Public Sub DeleteRowsByStatus_Final()
    Dim ws As Worksheet
    Dim srcRng As Range
    Dim statusCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim appState As TAppState

    On Error GoTo ErrHandler

    PushAppState appState

    Set ws = ThisWorkbook.Worksheets("Data")
    statusCol = 4

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then GoTo SafeExit

    Set srcRng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    If ws.FilterMode Then ws.ShowAllData

    srcRng.AutoFilter Field:=statusCol, Criteria1:="DELETE"

    On Error Resume Next
    srcRng.Offset(1, 0).Resize(srcRng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo ErrHandler

    If ws.FilterMode Then ws.ShowAllData

SafeExit:
    PopAppState appState
    Set srcRng = Nothing
    Set ws = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not ws Is Nothing Then
        If ws.FilterMode Then ws.ShowAllData
    End If
    On Error GoTo 0

    PopAppState appState
    MsgBox "DeleteRowsByStatus_Final failed: " & Err.Number & " - " & Err.Description, vbCritical
    Set srcRng = Nothing
    Set ws = Nothing
End Sub

Public Sub AggregateByKey_Final()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim srcData As Variant, outData() As Variant
    Dim map As Object
    Dim lastRow As Long
    Dim i As Long, outRow As Long
    Dim k As Variant
    Dim appState As TAppState

    On Error GoTo ErrHandler

    PushAppState appState

    Set wsSrc = ThisWorkbook.Worksheets("Data")
    Set wsOut = ThisWorkbook.Worksheets("Summary")

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo SafeExit

    srcData = wsSrc.Range("A1:B" & lastRow).Value2
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1

    For i = 2 To UBound(srcData, 1)
        If LenB(srcData(i, 1)) <> 0 And IsNumeric(srcData(i, 2)) Then
            If map.Exists(CStr(srcData(i, 1))) Then
                map(CStr(srcData(i, 1))) = CDbl(map(CStr(srcData(i, 1)))) + CDbl(srcData(i, 2))
            Else
                map.Add CStr(srcData(i, 1)), CDbl(srcData(i, 2))
            End If
        End If
    Next i

    wsOut.Cells.ClearContents
    wsOut.Range("A1:B1").Value = Array("Key", "Total")

    If map.Count > 0 Then
        ReDim outData(1 To map.Count, 1 To 2)
        outRow = 1

        For Each k In map.Keys
            outData(outRow, 1) = k
            outData(outRow, 2) = map(k)
            outRow = outRow + 1
        Next k

        wsOut.Range("A2").Resize(map.Count, 2).Value2 = outData
    End If

SafeExit:
    PopAppState appState
    Erase srcData
    Erase outData
    Set map = Nothing
    Set wsSrc = Nothing
    Set wsOut = Nothing
    Exit Sub

ErrHandler:
    PopAppState appState
    MsgBox "AggregateByKey_Final failed: " & Err.Number & " - " & Err.Description, vbCritical
    Erase srcData
    Erase outData
    Set map = Nothing
    Set wsSrc = Nothing
    Set wsOut = Nothing
End Sub
