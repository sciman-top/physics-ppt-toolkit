Attribute VB_Name = "PhysicsPptCommon"
Option Explicit

' Shared constants and utility functions for PhysicsPpt VBA macros.
'
' NOTE: VBA macros are an independent offline solution. They do NOT read config/physics-ppt-style.config.json.
' When updating the JSON config, sync these constants manually to keep PowerShell and VBA behavior consistent.
' Mapping: FONT_CN -> fonts.chinese | FONT_LATIN -> fonts.latin | FONT_MATH -> fonts.math
'          SIZE_* -> fontSizes.* | COLOR_* -> colors.*

' --- Font & Size Constants ---
Public Const FONT_CN As String = "微软雅黑"
Public Const FONT_LATIN As String = "Arial"
Public Const FONT_MATH As String = "Cambria Math"

Public Const SIZE_TITLE As Single = 46
Public Const SIZE_BODY As Single = 32
Public Const SIZE_TABLE_HEADER As Single = 30
Public Const SIZE_TABLE_BODY As Single = 28
Public Const SIZE_MINIMUM As Single = 24

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
