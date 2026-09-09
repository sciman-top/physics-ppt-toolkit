Attribute VB_Name = "PhysicsPptCommon"
Option Explicit

' Shared constants and utility functions for PhysicsPpt VBA macros.
'
' NOTE: VBA macros are an independent offline solution. They do NOT read config/physics-ppt-style.config.json.
' When updating the JSON config, sync these constants manually to keep PowerShell and VBA behavior consistent.
' Mapping (synced subset only):
'   FONT_CN -> fonts.chinese | FONT_LATIN -> fonts.latin | FONT_MATH -> fonts.math
'   SIZE_TITLE -> fontSizes.title1 | SIZE_BODY -> fontSizes.body | SIZE_TABLE_HEADER -> fontSizes.tableHeader
'   SIZE_TABLE_BODY -> fontSizes.tableBody | SIZE_BODY_MAX -> fontSizes.bodyMax
'   SIZE_MINIMUM -> fontSizes.minimum
'   COLOR_WHITE -> colors.white | COLOR_BODY -> colors.body
'   COLOR_YELLOW_FILL -> colors.yellowFill | COLOR_YELLOW_BORDER -> colors.yellowBorder
' The remaining fontSizes/colors keys (sectionTitle, title2, formula*, footer,
' sectionTitle/extensionTitle/formulaBlue/experimentGreen/darkGray) have no VBA counterpart:
' the VBA fallback intentionally normalizes to a smaller set than the PowerShell mainline.
' Background normalization is intentionally opt-in in the VBA fallback, matching the
' PowerShell default. Set this constant to True only after visual review of the deck.
Public Const NORMALIZE_SLIDE_BACKGROUND As Boolean = False

' --- Font & Size Constants ---
Public Const FONT_CN As String = "微软雅黑"
Public Const FONT_LATIN As String = "Arial"
Public Const FONT_MATH As String = "Cambria Math"

Public Const SIZE_TITLE As Single = 46
Public Const SIZE_BODY As Single = 32
Public Const SIZE_BODY_MAX As Single = 36
Public Const SIZE_TABLE_HEADER As Single = 30
Public Const SIZE_TABLE_BODY As Single = 28
Public Const SIZE_MINIMUM As Single = 24

' --- Color Constants (VBA Long = R + G*256 + B*65536) ---
Public Const COLOR_WHITE As Long = 16777215   ' RGB(255, 255, 255) = colors.white   #FFFFFF
Public Const COLOR_BODY As Long = 0           ' RGB(0, 0, 0)       = colors.body    #000000
Public Const COLOR_YELLOW_FILL As Long = 13431551   ' RGB(255, 242, 204) = colors.yellowFill  #FFF2CC
Public Const COLOR_YELLOW_BORDER As Long = 41942    ' RGB(214, 163, 0)   = colors.yellowBorder #D6A300

' Table cell styling diverges from the PowerShell mainline, which is report-only
' for tables ("TableStyleSkipped" to avoid cell overflow / row-height changes).
' Default False keeps the VBA fallback aligned with the mainline; set True only
' after visual review on a copy.
Public Const NORMALIZE_TABLE_STYLE As Boolean = False

' --- Office Enum Constants ---
Public Const MSO_TRUE_VAL As Long = -1
Public Const MSO_FALSE_VAL As Long = 0
Public Const MSO_PLACEHOLDER As Long = 14
Public Const MSO_TABLE As Long = 19
Public Const MSO_GROUP As Long = 6
Public Const MSO_MEDIA As Long = 16
Public Const PP_PLACEHOLDER_TITLE As Long = 1
Public Const PP_PLACEHOLDER_CENTER_TITLE As Long = 3

' --- Guard: ensure a presentation is active ---
Public Sub GuardActivePresentation()
    If Presentations.Count = 0 Then
        MsgBox "请先打开一个 PPT 文件再运行此宏。", vbExclamation
        Err.Raise vbObjectError + 513, "PhysicsPptCommon", "No active presentation."
    End If
End Sub

' --- Report path: fall back to Desktop if presentation is unsaved ---
Public Function GetReportPath(ByVal pres As Presentation, ByVal fileName As String) As String
    If Len(pres.Path) > 0 Then
        GetReportPath = pres.Path & "\" & fileName
    Else
        GetReportPath = Environ("USERPROFILE") & "\Desktop\" & fileName
    End If
End Function

' --- Shape helpers ---
Public Function ShapeHasText(ByVal shp As Shape) As Boolean
    On Error GoTo Failed
    ShapeHasText = (shp.TextFrame2.HasText = MSO_TRUE_VAL)
    Exit Function
Failed:
    ShapeHasText = False
End Function

Public Function GetShapeText(ByVal shp As Shape) As String
    On Error GoTo Failed
    If shp.TextFrame2.HasText = MSO_TRUE_VAL Then
        GetShapeText = CStr(shp.TextFrame2.TextRange.Text)
    Else
        GetShapeText = ""
    End If
    Exit Function
Failed:
    GetShapeText = ""
End Function

Public Function GetTextSize(ByVal shp As Shape) As Single
    On Error GoTo Failed
    GetTextSize = shp.TextFrame2.TextRange.Font.Size
    Exit Function
Failed:
    GetTextSize = 999
End Function

Public Function TryGetTextSize(ByVal shp As Shape, ByRef size As Single) As Boolean
    On Error GoTo Failed
    size = shp.TextFrame2.TextRange.Font.Size
    TryGetTextSize = True
    Exit Function
Failed:
    TryGetTextSize = False
End Function

Public Function ShapeHasTable(ByVal shp As Shape) As Boolean
    On Error GoTo Failed
    ShapeHasTable = CBool(shp.HasTable)
    Exit Function
Failed:
    ShapeHasTable = False
End Function

' --- Formula detection ---
Public Function IsFormulaCandidateText(ByVal txt As String) As Boolean
    Dim t As String
    t = Replace(Replace(Replace(Trim$(txt), vbCr, ""), vbLf, ""), " ", "")
    If Len(t) = 0 Or Len(t) > 80 Then
        IsFormulaCandidateText = False
        Exit Function
    End If

    If InStr(t, "=") > 0 Or InStr(t, "＝") > 0 Or InStr(t, "η") > 0 Or InStr(t, "Ω") > 0 Then
        IsFormulaCandidateText = True
        Exit Function
    End If

    If InStr(t, "W有") > 0 Or InStr(t, "W总") > 0 Or InStr(t, "W额") > 0 _
       Or InStr(t, "G物") > 0 Or InStr(t, "G动") > 0 _
       Or InStr(t, "R1") > 0 Or InStr(t, "R2") > 0 _
       Or InStr(t, "U1") > 0 Or InStr(t, "U2") > 0 _
       Or InStr(t, "I1") > 0 Or InStr(t, "I2") > 0 _
       Or InStr(t, "P1") > 0 Or InStr(t, "P2") > 0 Then
        IsFormulaCandidateText = True
        Exit Function
    End If

    IsFormulaCandidateText = False
End Function

' --- CSV helpers ---
Public Function CsvLine(ByVal fileName As String, ByVal slideNo As String, ByVal shapeName As String, ByVal issue As String, ByVal details As String) As String
    CsvLine = CsvEscape(fileName) & "," & CsvEscape(slideNo) & "," & CsvEscape(shapeName) & "," & CsvEscape(issue) & "," & CsvEscape(details)
End Function

Public Function CsvEscape(ByVal s As String) As String
    CsvEscape = """" & Replace(s, """", """""") & """"
End Function

' --- Report file output (UTF-8 BOM via ADODB.Stream) ---
Public Sub SaveReport(ByVal report As Collection, ByVal path As String)
    Dim stm As Object
    Dim sb As String
    Dim i As Long

    On Error GoTo Cleanup
    sb = "File,Slide,Shape,Issue,Details" & vbCrLf
    For i = 1 To report.Count
        sb = sb & CStr(report.Item(i)) & vbCrLf
    Next i

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText sb
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite

Cleanup:
    If Not stm Is Nothing Then
        If stm.State <> 0 Then stm.Close
    End If
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub
