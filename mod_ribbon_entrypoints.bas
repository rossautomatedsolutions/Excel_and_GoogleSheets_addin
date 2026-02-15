Attribute VB_Name = "mod_ribbon_entrypoints"
Option Explicit

Private Const APP_TITLE As String = "Excel Add-in"

' Generic Ribbon callback for regular buttons.
' Recommended Ribbon XML usage:
'   onAction="RibbonButton_OnAction"
'   tag="proc=mod_example.DoWork;ok=Completed"
Public Sub RibbonButton_OnAction(ByVal control As IRibbonControl)
    HandleRibbonAction control
End Sub

' Callback for toggle buttons. The pressed state is accepted to satisfy Ribbon signature
' and intentionally ignored for procedure dispatch.
Public Sub RibbonToggle_OnAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    HandleRibbonAction control
End Sub

Private Sub HandleRibbonAction(ByVal control As IRibbonControl)
    Dim procedureName As String
    Dim successMessage As String
    Dim errorMessage As String

    procedureName = ResolveProcedureName(control)
    successMessage = ResolveSuccessMessage(control)

    If Len(procedureName) = 0 Then
        MsgBox "This Ribbon command is not configured.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    If Not ExecuteRibbonProcedure(procedureName, errorMessage) Then
        MsgBox "Unable to run command '" & control.Id & "'." & vbCrLf & errorMessage, vbExclamation, APP_TITLE
        Exit Sub
    End If

    If Len(successMessage) > 0 Then
        MsgBox successMessage, vbInformation, APP_TITLE
    End If
End Sub

Private Function ExecuteRibbonProcedure(ByVal procedureName As String, ByRef errorMessage As String) As Boolean
    On Error GoTo RunWithoutBoost
    Application.Run "RunWithPerformanceBoost", procedureName
    ExecuteRibbonProcedure = True
    Exit Function

RunWithoutBoost:
    Err.Clear
    On Error GoTo ExecutionFailed
    Application.Run procedureName
    ExecuteRibbonProcedure = True
    Exit Function

ExecutionFailed:
    errorMessage = Err.Description
    ExecuteRibbonProcedure = False
End Function

Private Function ResolveProcedureName(ByVal control As IRibbonControl) As String
    Dim tagProcedure As String

    tagProcedure = GetTagValue(control, "proc")
    If Len(tagProcedure) > 0 Then
        ResolveProcedureName = tagProcedure
        Exit Function
    End If

    ' Convention fallback:
    ' 1) control.Id equals full VBA procedure name (e.g. mod_sync.RunNow)
    ' 2) control.Id maps to Ribbon_<Id> in this project
    If InStr(1, control.Id, ".", vbTextCompare) > 0 Then
        ResolveProcedureName = control.Id
    Else
        ResolveProcedureName = "Ribbon_" & control.Id
    End If
End Function

Private Function ResolveSuccessMessage(ByVal control As IRibbonControl) As String
    ResolveSuccessMessage = GetTagValue(control, "ok")
End Function

Private Function GetTagValue(ByVal control As IRibbonControl, ByVal keyName As String) As String
    Dim tagText As String
    Dim pairs() As String
    Dim i As Long
    Dim kvp() As String

    tagText = Trim$(control.Tag)
    If Len(tagText) = 0 Then Exit Function

    pairs = Split(tagText, ";")
    For i = LBound(pairs) To UBound(pairs)
        kvp = Split(pairs(i), "=", 2)
        If UBound(kvp) = 1 Then
            If StrComp(Trim$(kvp(0)), keyName, vbTextCompare) = 0 Then
                GetTagValue = Trim$(kvp(1))
                Exit Function
            End If
        End If
    Next i
End Function
