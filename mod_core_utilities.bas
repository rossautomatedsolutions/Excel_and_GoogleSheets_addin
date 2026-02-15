Attribute VB_Name = "mod_core_utilities"
Option Explicit

Public Sub RunWithPerformanceBoost(ByVal taskName As String, ByVal taskProcedure As String)
    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation
    Dim startTime As Double
    Dim elapsedSeconds As Double

    On Error GoTo HandleError

    If LenB(Trim$(taskProcedure)) = 0 Then
        Err.Raise vbObjectError + 2001, "RunWithPerformanceBoost", "taskProcedure cannot be blank."
    End If

    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    startTime = Timer

    Application.Run taskProcedure

CleanExit:
    On Error Resume Next
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    On Error GoTo 0

    If startTime > 0 Then
        elapsedSeconds = Timer - startTime
        If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#

        Debug.Print "[Performance] " & IIf(LenB(Trim$(taskName)) > 0, taskName, taskProcedure) & _
                    " completed in " & Format$(elapsedSeconds, "0.000") & " seconds."
    End If

    Exit Sub

HandleError:
    Debug.Print "[Performance][ERROR] " & IIf(LenB(Trim$(taskName)) > 0, taskName, taskProcedure) & _
                " failed. " & Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub
