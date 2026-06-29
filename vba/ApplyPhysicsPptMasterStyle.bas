Attribute VB_Name = "ApplyPhysicsPptMasterStyle"
Option Explicit

' Apply a conservative master style to the active presentation.
' It updates slide master text defaults and does not move existing slide objects.
' Requires: PhysicsPptCommon module (shared constants and utilities).

Public Sub ApplyPhysicsMasterStyleToActivePresentation()
    GuardActivePresentation

    Dim pres As Presentation
    Dim master As Master
    Dim layout As CustomLayout
    Dim shp As Shape

    Set pres = ActivePresentation
    Set master = pres.SlideMaster

    On Error Resume Next
    master.Background.Fill.Solid
    master.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
    On Error GoTo 0

    For Each shp In master.Shapes
        NormalizeMasterShape shp
    Next shp

    For Each layout In master.CustomLayouts
        For Each shp In layout.Shapes
            NormalizeMasterShape shp
        Next shp
    Next layout

    MsgBox "母版基础样式已应用。建议进入""视图 > 幻灯片母版""人工检查各版式。", vbInformation
End Sub

Private Sub NormalizeMasterShape(ByVal shp As Shape)
    On Error GoTo Failed
    If shp.TextFrame2.HasText = MSO_TRUE_VAL Then
        With shp.TextFrame2.TextRange.Font
            .Name = FONT_LATIN
            .NameFarEast = FONT_CN
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            If shp.Top < 90 Then
                .Size = SIZE_TITLE
                .Bold = msoTrue
            Else
                .Size = SIZE_BODY
                .Bold = msoFalse
            End If
        End With
    End If
    Exit Sub
Failed:
    Err.Clear
End Sub
