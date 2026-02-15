Attribute VB_Name = "mod_ribbon_entrypoints"
Option Explicit

' Ribbon object cache (optional, used when invalidation is needed)
Private pRibbon As IRibbonUI

' ========= Ribbon lifecycle =========
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set pRibbon = ribbon
End Sub

' ========= Sheet Management =========
Public Sub BtnRefreshSheetIndex(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_sheet_management.RefreshSheetIndex
End Sub

Public Sub BtnAddStandardSheet(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_sheet_management.AddStandardSheet
End Sub

Public Sub BtnHideConfigSheets(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_sheet_management.HideConfigSheets
End Sub

' ========= Formatting =========
Public Sub BtnApplyReportTheme(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_formatting.ApplyReportTheme
End Sub

Public Sub BtnFormatHeaders(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_formatting.FormatHeaders
End Sub

Public Sub BtnAutoFitUsedRange(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_formatting.AutoFitUsedRange
End Sub

' ========= Navigation =========
Public Sub BtnGoToDashboard(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_navigation.GoToDashboard
End Sub

Public Sub BtnNextDataSheet(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_navigation.NextDataSheet
End Sub

Public Sub BtnPreviousDataSheet(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_navigation.PreviousDataSheet
End Sub

' ========= Data Tools =========
Public Sub BtnRunDataValidation(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_data_tools.RunDataValidation
End Sub

Public Sub BtnDeduplicateSelection(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_data_tools.DeduplicateSelection
End Sub

Public Sub BtnExportCurrentRange(control As IRibbonControl)
    ' TODO: call your implementation, e.g. mod_data_tools.ExportCurrentRange
End Sub
