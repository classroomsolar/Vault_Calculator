Attribute VB_Name = "OneClickReport"
Option Explicit

' Underground Enclosure / Vault Calculator
' v4.0 Batch 9 — Batches 1–9 integrated; Batch 10 final package/QA remains
' Desktop Excel only. Import this module into the workbook and save as .xlsm.

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub RequireCoreSheet(ByVal sheetName As String)
    If Not SheetExists(sheetName) Then
        Err.Raise vbObjectError + 101, "OneClickReport", _
            "Required report sheet is missing: " & sheetName
    End If
End Sub

Private Sub AddCoreSheet(ByRef c As Collection, ByVal sheetName As String)
    If SheetExists(sheetName) Then c.Add sheetName
End Sub

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(1 To c.Count)
    For i = 1 To c.Count
        arr(i) = c(i)
    Next i
    CollectionToArray = arr
End Function

Private Sub ValidateCoreSheets()
    RequireCoreSheet "00_HOME"
    RequireCoreSheet "00_USER_INPUT"
    RequireCoreSheet "02_RESULTS"
    RequireCoreSheet "03_LAYOUT"
    RequireCoreSheet "04_REPORT"
    RequireCoreSheet "07_PULL_CARD"
    RequireCoreSheet "14_PRINT_CENTER"
    RequireCoreSheet "16_DECISION_DASHBOARD"
    RequireCoreSheet "17_ACTION_PLAN"
    RequireCoreSheet "21_REQUIREMENTS_CLOSURE"
    RequireCoreSheet "22_ELECTRICAL_SCREEN"
    RequireCoreSheet "23_HARD_TEST_REPORT"
    RequireCoreSheet "29_FINAL_REGRESSION"
    RequireCoreSheet "31_FRESH_EXECUTION_200"
    RequireCoreSheet "32_AUDIT_FIX_REGRESSION"
    RequireCoreSheet "33_MASTER_REQUIREMENTS"
    RequireCoreSheet "34_10_BATCH_ROADMAP"
    RequireCoreSheet "35_BATCH_GATE"
    RequireCoreSheet "39_BATCH_3_REGRESSION"
    RequireCoreSheet "43_BATCH_5_REGRESSION"
    RequireCoreSheet "44_SPLICE_TERM_WORKSPACE"
    RequireCoreSheet "45_BATCH_6_REGRESSION"
    RequireCoreSheet "46_SITE_STRUCT_ENV"
    RequireCoreSheet "47_BATCH_7_REGRESSION"
    RequireCoreSheet "48_STUDY_HANDOFF_RESERVE"
    RequireCoreSheet "49_BATCH_8_REGRESSION"
    RequireCoreSheet "50_VAULT_SCHEDULE_BOM"
    RequireCoreSheet "51_AUTO_PLAN_SECTION"
    RequireCoreSheet "52_BATCH_9_REGRESSION"
    RequireCoreSheet "42_SUPPORT_STACK_ANCHOR"
    RequireCoreSheet "41_BATCH_4_REGRESSION"
    RequireCoreSheet "40_ROUTE_PULL_INTEGRATION"
    RequireCoreSheet "38_BATCH_2_REGRESSION"
    RequireCoreSheet "37_PULL_SECTION_PLANNER"
    RequireCoreSheet "36_DUCT_BANK_GEOMETRY"
End Sub

Public Sub RefreshAndValidate()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.CalculateFullRebuild
    ValidateCoreSheets
    ThisWorkbook.Worksheets("14_PRINT_CENTER").Activate
    Application.ScreenUpdating = True
    MsgBox "Refresh and validation completed successfully.", vbInformation, "Vault Calculator"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Refresh / validation stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub

Private Function FullReportSheets() As Variant
    Dim c As New Collection
    AddCoreSheet c, "02_RESULTS"
    AddCoreSheet c, "03_LAYOUT"
    AddCoreSheet c, "04_REPORT"
    AddCoreSheet c, "05_PROJECT_QUAL"
    AddCoreSheet c, "06_CODE_RECOVERY"
    AddCoreSheet c, "07_PULL_CARD"
    AddCoreSheet c, "09_PULL_SENSITIVITY"
    AddCoreSheet c, "10_REDESIGN_ADVISOR"
    AddCoreSheet c, "12_ROUTE_NETWORK"
    AddCoreSheet c, "13_THERMAL_SCREEN"
    AddCoreSheet c, "14_PRINT_CENTER"
    AddCoreSheet c, "15_MASTER_AUDIT"
    AddCoreSheet c, "16_DECISION_DASHBOARD"
    AddCoreSheet c, "17_ACTION_PLAN"
    AddCoreSheet c, "20_TOLERANCE_SENSITIVITY"
    AddCoreSheet c, "21_REQUIREMENTS_CLOSURE"
    AddCoreSheet c, "22_ELECTRICAL_SCREEN"
    AddCoreSheet c, "23_HARD_TEST_REPORT"
    AddCoreSheet c, "25_TOOL_BENCHMARK_100"
    AddCoreSheet c, "26_RESEARCH_REGISTER"
    AddCoreSheet c, "27_CHALLENGE_200"
    AddCoreSheet c, "29_FINAL_REGRESSION"
    AddCoreSheet c, "31_FRESH_EXECUTION_200"
    AddCoreSheet c, "32_AUDIT_FIX_REGRESSION"
    AddCoreSheet c, "33_MASTER_REQUIREMENTS"
    AddCoreSheet c, "34_10_BATCH_ROADMAP"
    AddCoreSheet c, "35_BATCH_GATE"
    AddCoreSheet c, "36_DUCT_BANK_GEOMETRY"
    AddCoreSheet c, "37_PULL_SECTION_PLANNER"
    AddCoreSheet c, "38_BATCH_2_REGRESSION"
    AddCoreSheet c, "39_BATCH_3_REGRESSION"
    AddCoreSheet c, "40_ROUTE_PULL_INTEGRATION"
    AddCoreSheet c, "41_BATCH_4_REGRESSION"
    AddCoreSheet c, "42_SUPPORT_STACK_ANCHOR"
    AddCoreSheet c, "43_BATCH_5_REGRESSION"
    AddCoreSheet c, "44_SPLICE_TERM_WORKSPACE"
    AddCoreSheet c, "45_BATCH_6_REGRESSION"
    AddCoreSheet c, "46_SITE_STRUCT_ENV"
    AddCoreSheet c, "47_BATCH_7_REGRESSION"
    AddCoreSheet c, "48_STUDY_HANDOFF_RESERVE"
    AddCoreSheet c, "49_BATCH_8_REGRESSION"
    AddCoreSheet c, "50_VAULT_SCHEDULE_BOM"
    AddCoreSheet c, "51_AUTO_PLAN_SECTION"
    AddCoreSheet c, "52_BATCH_9_REGRESSION"
    FullReportSheets = CollectionToArray(c)
End Function

Private Function SummaryReportSheets() As Variant
    Dim c As New Collection
    AddCoreSheet c, "02_RESULTS"
    AddCoreSheet c, "03_LAYOUT"
    AddCoreSheet c, "04_REPORT"
    AddCoreSheet c, "16_DECISION_DASHBOARD"
    AddCoreSheet c, "17_ACTION_PLAN"
    AddCoreSheet c, "21_REQUIREMENTS_CLOSURE"
    AddCoreSheet c, "22_ELECTRICAL_SCREEN"
    AddCoreSheet c, "29_FINAL_REGRESSION"
    AddCoreSheet c, "32_AUDIT_FIX_REGRESSION"
    AddCoreSheet c, "33_MASTER_REQUIREMENTS"
    AddCoreSheet c, "35_BATCH_GATE"
    AddCoreSheet c, "36_DUCT_BANK_GEOMETRY"
    AddCoreSheet c, "38_BATCH_2_REGRESSION"
    AddCoreSheet c, "39_BATCH_3_REGRESSION"
    AddCoreSheet c, "40_ROUTE_PULL_INTEGRATION"
    AddCoreSheet c, "42_SUPPORT_STACK_ANCHOR"
    AddCoreSheet c, "41_BATCH_4_REGRESSION"
    AddCoreSheet c, "43_BATCH_5_REGRESSION"
    AddCoreSheet c, "44_SPLICE_TERM_WORKSPACE"
    AddCoreSheet c, "46_SITE_STRUCT_ENV"
    AddCoreSheet c, "45_BATCH_6_REGRESSION"
    AddCoreSheet c, "47_BATCH_7_REGRESSION"
    AddCoreSheet c, "48_STUDY_HANDOFF_RESERVE"
    AddCoreSheet c, "50_VAULT_SCHEDULE_BOM"
    AddCoreSheet c, "51_AUTO_PLAN_SECTION"
    AddCoreSheet c, "49_BATCH_8_REGRESSION"
    AddCoreSheet c, "52_BATCH_9_REGRESSION"
    SummaryReportSheets = CollectionToArray(c)
End Function

Private Function PullPackageSheets() As Variant
    Dim c As New Collection
    AddCoreSheet c, "02_RESULTS"
    AddCoreSheet c, "03_LAYOUT"
    AddCoreSheet c, "07_PULL_CARD"
    AddCoreSheet c, "09_PULL_SENSITIVITY"
    AddCoreSheet c, "10_REDESIGN_ADVISOR"
    AddCoreSheet c, "17_ACTION_PLAN"
    AddCoreSheet c, "20_TOLERANCE_SENSITIVITY"
    AddCoreSheet c, "37_PULL_SECTION_PLANNER"
    AddCoreSheet c, "40_ROUTE_PULL_INTEGRATION"
    AddCoreSheet c, "42_SUPPORT_STACK_ANCHOR"
    PullPackageSheets = CollectionToArray(c)
End Function

Private Sub PrintSheetArray(ByVal arr As Variant)
    ThisWorkbook.Worksheets(arr).Select
    ActiveWindow.SelectedSheets.PrintOut
    ThisWorkbook.Worksheets("14_PRINT_CENTER").Select
End Sub

Public Sub PrintFullReport()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.CalculateFullRebuild
    ValidateCoreSheets
    PrintSheetArray FullReportSheets()
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Full report print stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub

Public Sub ExportFullReportPDF()
    Dim fn As Variant, arr As Variant
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.CalculateFullRebuild
    ValidateCoreSheets
    fn = Application.GetSaveAsFilename( _
        InitialFileName:="Underground_Enclosure_Installation_Report.pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")
    If fn = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    arr = FullReportSheets()
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(fn), _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Worksheets("14_PRINT_CENTER").Select
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "PDF export stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub

Public Sub PrintSummaryReport()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.CalculateFullRebuild
    ValidateCoreSheets
    PrintSheetArray SummaryReportSheets()
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Summary print stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub

Public Sub PrintPullPackage()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.CalculateFullRebuild
    ValidateCoreSheets
    PrintSheetArray PullPackageSheets()
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Pull package print stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub

Private Sub DeleteReportButtons(ByVal ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left$(shp.Name, 14) = "VaultCalcBtn_" Then shp.Delete
    Next shp
    On Error GoTo 0
End Sub

Private Sub AddMacroButton(ByVal ws As Worksheet, ByVal buttonName As String, _
                           ByVal caption As String, ByVal macroName As String, _
                           ByVal leftPos As Double, ByVal topPos As Double, _
                           ByVal widthVal As Double, ByVal heightVal As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.Name = buttonName
    shp.TextFrame2.TextRange.Characters.Text = caption
    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Public Sub InstallPrintButtonsSilent()
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets("14_PRINT_CENTER")
    DeleteReportButtons ws
    AddMacroButton ws, "VaultCalcBtn_Refresh", "Refresh + Validate", "RefreshAndValidate", 28, 48, 130, 28
    AddMacroButton ws, "VaultCalcBtn_Full", "Print Full Report", "PrintFullReport", 170, 48, 130, 28
    AddMacroButton ws, "VaultCalcBtn_PDF", "Export Full PDF", "ExportFullReportPDF", 312, 48, 130, 28
    AddMacroButton ws, "VaultCalcBtn_Summary", "Print Summary", "PrintSummaryReport", 454, 48, 130, 28
    AddMacroButton ws, "VaultCalcBtn_Pull", "Print Pull Package", "PrintPullPackage", 596, 48, 145, 28
    Exit Sub
Fail:
    Err.Raise Err.Number, "InstallPrintButtonsSilent", Err.Description
End Sub

Public Sub InstallPrintButtons()
    On Error GoTo Fail
    InstallPrintButtonsSilent
    MsgBox "One-click report buttons installed on 14_PRINT_CENTER.", vbInformation, "Vault Calculator"
    Exit Sub
Fail:
    MsgBox "Button installation stopped: " & Err.Description, vbCritical, "Vault Calculator"
End Sub
