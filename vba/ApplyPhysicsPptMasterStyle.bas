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
    Dim masterTextSize As Single

    Set pres = ActivePresentation
    Set master = pres.SlideMaster

    On Error Resume Next
    ' Opt-in like the slide background: config keeps SLIDE.BACKGROUND disabled by
    ' default, so the master background is only whitened when the operator has
    ' deliberately enabled background normalization in PhysicsPptCommon.
    If NORMALIZE_SLIDE_BACKGROUND Then
        master.Background.Fill.Solid
        master.Background.Fill.ForeColor.RGB = COLOR_WHITE
    End If
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
            .Fill.ForeColor.RGB = COLOR_BODY
            masterTextSize = GetSafeMasterTextSize(shp)
            If shp.Top < 90 Then
                If masterTextSize >= 0 And masterTextSize < SIZE_TITLE Then .Size = SIZE_TITLE
                .Bold = msoTrue
            Else
                If masterTextSize > SIZE_BODY_MAX Then .Size = SIZE_BODY_MAX
                .Bold = msoFalse
            End If
        End With
    End If
    Exit Sub
Failed:
    Err.Clear
End Sub

Private Function GetSafeMasterTextSize(ByVal shp As Shape) As Single
    Dim size As Single
    If TryGetTextSize(shp, size) Then
        GetSafeMasterTextSize = size
    Else
        GetSafeMasterTextSize = -1
    End If
End Function
