Attribute VB_Name = "PhysicsPptNormalize"
Option Explicit

' Low-risk normalization macro for junior-middle-school physics PPT.
' It changes style only. It does not change text content, positions, sizes, animations, or picture crops.
' Requires: PhysicsPptCommon module (shared constants and utilities).

Public Sub NormalizeCurrentPresentation()
    GuardActivePresentation

    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim shpText As String
    Dim isVideo As Boolean
    Dim report As Collection
    Dim reportPath As String

    Set pres = ActivePresentation
    Set report = New Collection

    For Each sld In pres.Slides
        isVideo = IsVideoSlide(sld)
        SetSlideBackground sld, isVideo
        report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), "(slide)", "SlideType", IIf(isVideo, "VideoOrMediaCandidate", "Normal"))

        For Each shp In sld.Shapes
            On Error GoTo ShapeFailed

            If shp.Type = MSO_GROUP Then
                report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "GroupShapeSkipped", "Grouped shapes are not modified.")
                GoTo NextShape
            End If

            If ShapeHasTable(shp) Then
                NormalizeTableShape shp
                GoTo NextShape
            End If

            If ShapeHasText(shp) Then
                shpText = GetShapeText(shp)
                If IsFormulaCandidateText(shpText) Then
                    report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "FormulaCandidate", shpText)
                End If
                NormalizeTextShape shp, isVideo
                NormalizeHighlightBox shp
                ClearDecorativeEffects shp
            End If

            GoTo NextShape

ShapeFailed:
            report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "ShapeFailed", Err.Description)
            Err.Clear
            Resume NextShape

NextShape:
            On Error GoTo 0
        Next shp
    Next sld

    reportPath = GetReportPath(pres, "physics-ppt-vba-normalize-report.csv")
    SaveReport report, reportPath
    MsgBox "规范化完成。报告已保存：" & vbCrLf & reportPath, vbInformation
End Sub

Private Function IsTitleShape(ByVal shp As Shape) As Boolean
    On Error GoTo Fallback
    If shp.Type = MSO_PLACEHOLDER Then
        If shp.PlaceholderFormat.Type = PP_PLACEHOLDER_TITLE Or shp.PlaceholderFormat.Type = PP_PLACEHOLDER_CENTER_TITLE Then
            IsTitleShape = True
            Exit Function
        End If
    End If
Fallback:
    On Error Resume Next
    IsTitleShape = (Len(Trim$(GetShapeText(shp))) <= 24 And shp.Top < 90)
End Function

Private Function IsVideoSlide(ByVal sld As Slide) As Boolean
    Dim shp As Shape
    Dim txt As String
    For Each shp In sld.Shapes
        On Error Resume Next
        If shp.Type = MSO_MEDIA Then
            IsVideoSlide = True
            Exit Function
        End If
        txt = GetShapeText(shp)
        If Len(txt) > 0 Then
            If InStr(1, txt, "视频", vbTextCompare) > 0 Then
                IsVideoSlide = True
                Exit Function
            End If
            If InStr(1, txt, "播放", vbTextCompare) > 0 Then
                IsVideoSlide = True
                Exit Function
            End If
            If InStr(1, txt, "观察视频", vbTextCompare) > 0 Then
                IsVideoSlide = True
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next shp
    IsVideoSlide = False
End Function

Private Sub SetSlideBackground(ByVal sld As Slide, ByVal isVideo As Boolean)
    On Error Resume Next
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.Solid
    If isVideo Then
        sld.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Else
        sld.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    On Error GoTo 0
End Sub

Private Sub NormalizeTextShape(ByVal shp As Shape, ByVal isVideo As Boolean)
    Dim isTitle As Boolean
    Dim targetSize As Single
    Dim targetColor As Long

    On Error GoTo Failed
    isTitle = IsTitleShape(shp)
    targetSize = IIf(isTitle, SIZE_TITLE, SIZE_BODY)
    targetColor = IIf(isVideo, RGB(255, 255, 255), RGB(0, 0, 0))

    With shp.TextFrame2.TextRange.Font
        .Name = FONT_LATIN
        .NameFarEast = FONT_CN
        .Size = targetSize
        .Bold = IIf(isTitle, msoTrue, msoFalse)
        .Fill.ForeColor.RGB = targetColor
    End With
    Exit Sub
Failed:
    Err.Clear
End Sub

Private Sub NormalizeTableShape(ByVal shp As Shape)
    Dim r As Long, c As Long
    Dim cellShape As Shape
    Dim isHeader As Boolean

    On Error GoTo Failed
    For r = 1 To shp.Table.Rows.Count
        For c = 1 To shp.Table.Columns.Count
            Set cellShape = shp.Table.Cell(r, c).Shape
            If ShapeHasText(cellShape) Then
                isHeader = (r = 1)
                With cellShape.TextFrame2.TextRange.Font
                    .Name = FONT_LATIN
                    .NameFarEast = FONT_CN
                    .Size = IIf(isHeader, SIZE_TABLE_HEADER, SIZE_TABLE_BODY)
                    .Bold = IIf(isHeader, msoTrue, msoFalse)
                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
                End With
            End If
        Next c
    Next r
    Exit Sub
Failed:
    Err.Clear
End Sub

Private Sub ClearDecorativeEffects(ByVal shp As Shape)
    On Error Resume Next
    shp.Shadow.Visible = msoFalse
    shp.Glow.Radius = 0
    shp.SoftEdge.Radius = 0
    On Error GoTo 0
End Sub

Private Function IsYellowishFill(ByVal rgb As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = rgb And &HFF&
    g = (rgb \ &H100&) And &HFF&
    b = (rgb \ &H10000) And &HFF&
    IsYellowishFill = (r > 200 And g > 200 And b < 180)
End Function

Private Sub NormalizeHighlightBox(ByVal shp As Shape)
    On Error Resume Next
    If shp.Fill.Visible = msoTrue Then
        If IsYellowishFill(shp.Fill.ForeColor.RGB) Then
            shp.Fill.ForeColor.RGB = RGB(255, 242, 204)
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = RGB(214, 163, 0)
            shp.Line.Weight = 1.75
        End If
    End If
    On Error GoTo 0
End Sub
