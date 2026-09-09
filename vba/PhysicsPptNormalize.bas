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
        If NORMALIZE_SLIDE_BACKGROUND Then
            SetSlideBackground sld, isVideo
        Else
            report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), "(slide)", "SlideBackgroundCandidate", "Background preserved by default; enable NORMALIZE_SLIDE_BACKGROUND only after visual review.")
        End If
        report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), "(slide)", "SlideType", IIf(isVideo, "VideoOrMediaCandidate", "Normal"))

        For Each shp In sld.Shapes
            On Error GoTo ShapeFailed

            If shp.Type = MSO_GROUP Then
                report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "GroupShapeSkipped", "Grouped shapes are not modified.")
                GoTo NextShape
            End If

            If ShapeHasTable(shp) Then
                If NORMALIZE_TABLE_STYLE Then
                    NormalizeTableShape shp
                Else
                    report.Add CsvLine(pres.Name, CStr(sld.SlideIndex), shp.Name, "TableStyleSkipped", "Table styles preserved to avoid cell overflow or row-height changes.")
                End If
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
    Dim beforeSize As Single
    Dim left As Single, top As Single, width As Single, height As Single
    Dim autoSize As Long
    Dim geometryDrift As Single

    On Error GoTo Failed
    isTitle = IsTitleShape(shp)
    If Not TryGetTextSize(shp, beforeSize) Then Err.Raise vbObjectError + 541, "NormalizeTextShape", "Text size could not be read safely."
    targetSize = beforeSize
    If isTitle Then
        If beforeSize < SIZE_TITLE Then targetSize = SIZE_TITLE
    ElseIf beforeSize > SIZE_BODY_MAX Then
        targetSize = SIZE_BODY_MAX
    End If
    left = shp.Left
    top = shp.Top
    width = shp.Width
    height = shp.Height
    autoSize = shp.TextFrame2.AutoSize
    ' White text is only written when the matching black video-slide background
    ' is actually applied (NORMALIZE_SLIDE_BACKGROUND), mirroring the PowerShell gate.
    If isVideo And NORMALIZE_SLIDE_BACKGROUND Then
        targetColor = COLOR_WHITE
    Else
        targetColor = COLOR_BODY
    End If

    With shp.TextFrame2.TextRange.Font
        .Name = FONT_LATIN
        .NameFarEast = FONT_CN
        If targetSize <> beforeSize Then .Size = targetSize
        .Bold = IIf(isTitle, msoTrue, msoFalse)
        .Fill.ForeColor.RGB = targetColor
    End With
    If autoSize <> 0 Then
        geometryDrift = Abs(shp.Left - left)
        If Abs(shp.Top - top) > geometryDrift Then geometryDrift = Abs(shp.Top - top)
        If Abs(shp.Width - width) > geometryDrift Then geometryDrift = Abs(shp.Width - width)
        If Abs(shp.Height - height) > geometryDrift Then geometryDrift = Abs(shp.Height - height)
        If geometryDrift > 0.05 Then
            shp.TextFrame2.TextRange.Font.Size = beforeSize
            shp.Left = left
            shp.Top = top
            shp.Width = width
            shp.Height = height
            Err.Raise vbObjectError + 542, "NormalizeTextShape", "AutoSize reflowed geometry; font size and bounds were rolled back."
        End If
    End If
    Exit Sub
Failed:
    Err.Raise Err.Number, "NormalizeTextShape", Err.Description
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
                    .Fill.ForeColor.RGB = COLOR_BODY
                End With
            End If
        Next c
    Next r
    Exit Sub
Failed:
    Err.Raise Err.Number, "NormalizeTableShape", Err.Description
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

' Matches the PowerShell mainline (STYLE.HIGHLIGHT.TEXT_COLOR, allowedProperties=[Color]):
' only the text color of yellow-highlight boxes is normalized to the body color;
' fill and border are preserved untouched. The previous fill/border rewrite
' (including the fixed 1.75pt border) exceeded the configured authorization set.
Private Sub NormalizeHighlightBox(ByVal shp As Shape)
    On Error Resume Next
    If shp.Fill.Visible = msoTrue Then
        If IsYellowishFill(shp.Fill.ForeColor.RGB) Then
            If ShapeHasText(shp) Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_BODY
            End If
        End If
    End If
    On Error GoTo 0
End Sub
