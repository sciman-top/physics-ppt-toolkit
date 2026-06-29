Attribute VB_Name = "PhysicsPptReportOnly"
Option Explicit

' Report-only macro. It does not modify the presentation.
' Requires: PhysicsPptCommon module (shared constants and utilities).

Public Sub ReportCurrentPresentationStyle()
    GuardActivePresentation

    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim shpText As String
    Dim report As Collection
    Dim reportPath As String

    Set pres = ActivePresentation
    Set report = New Collection

    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            If ShapeHasText(shp) Then
                shpText = GetShapeText(shp)
                If IsFormulaCandidateText(shpText) Then
                    report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "FormulaCandidate", shpText)
                End If
                If GetTextSize(shp) < SIZE_MINIMUM Then
                    report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "SmallText", CStr(GetTextSize(shp)) & " pt")
                End If
            End If
            If shp.Type = MSO_MEDIA Then
                report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "MediaObject", "This slide may need black video background.")
            End If
            On Error GoTo 0
        Next shp
    Next sld

    reportPath = GetReportPath(pres, "physics-ppt-vba-report-only.csv")
    SaveReport report, reportPath
    MsgBox "检查报告已保存：" & vbCrLf & reportPath, vbInformation
End Sub
