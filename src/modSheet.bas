Option Explicit
'==============================================================
' modSheet - Do layout bang diem + nhan dien Lop / Dot / Mon
'==============================================================
Public Type NX_Layout
    Ok As Boolean
    SubjectKey As String
    SubjectName As String
    HeaderRow As Long
    NameCol As Long
    NoteCols As String          ' cac cot Nhan xet, dang "7" hoac "7,9"
    ScoreCol As Long            ' cot Diem / KTDK
    LevelCol As Long            ' cot Muc dat / Xep loai
    FirstRow As Long
    LastRow As Long
    Warn As String
End Type

Public Function NX_Detect(ByVal ws As Object) As NX_Layout
    Dim res As NX_Layout
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim scanR As Long, scanC As Long
    Dim grid As Variant
    Dim p As String
    Dim bestRow As Long, bestCnt As Long, cnt As Long
    Dim noteCols As String, dataStart As Long
    Dim cel As Object, ma As Object, mr As Long, mc As Long

    On Error GoTo Fail
    res.Ok = False
    res.SubjectName = ws.Name
    res.SubjectKey = NX_SubjectCode(ws.Name)
    res.Warn = ""

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastRow < 2 Then
        res.Warn = "Sheet trong"
        NX_Detect = res
        Exit Function
    End If
    If lastRow > 100000 Then lastRow = 100000
    If lastCol < 1 Then lastCol = 1
    If lastCol > 100 Then lastCol = 100

    scanR = lastRow
    If scanR > 60 Then scanR = 60
    scanC = lastCol
    If scanC > 60 Then scanC = 60

    grid = ws.Range(ws.Cells(1, 1), ws.Cells(scanR, scanC)).Value

    '------------------------------------------------- 1. dong tieu de "Nhan xet"
    bestRow = 0
    bestCnt = 0
    For r = 1 To scanR
        cnt = 0
        For c = 1 To scanC
            p = NX_Plain(NX_GridText(grid, r, c))
            If InStr(1, p, "nhanxet", vbTextCompare) > 0 Then cnt = cnt + 1
        Next c
        If cnt > bestCnt Then
            bestCnt = cnt
            bestRow = r
        End If
    Next r

    If bestRow = 0 Then
        res.Warn = "Khong tim thay cot Nhan xet"
        NX_Detect = res
        Exit Function
    End If

    '------------------------------------------------- 2. cot Nhan xet + dong bat dau du lieu
    noteCols = ""
    dataStart = 0
    For c = 1 To lastCol
        p = NX_Plain(NX_GridText(grid, bestRow, c))
        If InStr(1, p, "nhanxet", vbTextCompare) > 0 Then
            Set cel = ws.Cells(bestRow, c)
            Set ma = cel.MergeArea
            mc = ma.Column
            mr = ma.Row
            If InStr(1, "," & noteCols & ",", "," & mc & ",") = 0 Then
                If Len(noteCols) = 0 Then
                    noteCols = CStr(mc)
                Else
                    noteCols = noteCols & "," & CStr(mc)
                End If
            End If
            If mr + ma.Rows.Count > dataStart Then dataStart = mr + ma.Rows.Count
            Set ma = Nothing
            Set cel = Nothing
        End If
    Next c

    '------------------------------------------------- 3. cot Ho ten / Diem / Muc dat
    res.NameCol = NX_FindHeaderCol(grid, bestRow, scanC, Array("hovaten", "hoten", "tenhocsinh", "hocsinh", "ten"))
    res.ScoreCol = NX_FindHeaderCol(grid, bestRow, scanC, Array("diemktdk", "ktdk", "diemtb", "dtb", "diem"))
    res.LevelCol = NX_FindHeaderCol(grid, bestRow, scanC, Array("mucdat", "xeploai", "danhhieu", "ketqua", "hocluc", "xl"))

    res.Ok = True
    res.HeaderRow = bestRow
    res.NoteCols = noteCols
    res.FirstRow = IIf(dataStart > 0, dataStart, bestRow + 1)
    res.LastRow = lastRow
    NX_Detect = res
    Exit Function

Fail:
    res.Ok = False
    res.Warn = Err.Description
    NX_Detect = res
End Function

' Tim cot theo tu khoa trong vung tieu de (loai tru cac cot nhan xet / diem manh / gop y)
Private Function NX_FindHeaderCol(ByVal grid As Variant, ByVal headerRow As Long, ByVal scanC As Long, ByVal keywords As Variant) As Long
    Dim r As Long, c As Long, p As String, kw As Variant
    Dim bestCol As Long, bestLen As Long, kwLen As Long

    ' 1) uu tien dong tieu de
    For c = 1 To scanC
        p = NX_Plain(NX_GridText(grid, headerRow, c))
        If Len(p) > 0 Then
            If InStr(1, p, "nhanxet", vbTextCompare) = 0 And _
               InStr(1, p, "diemmanh", vbTextCompare) = 0 And _
               InStr(1, p, "gopy", vbTextCompare) = 0 Then
                For Each kw In keywords
                    kwLen = Len(CStr(kw))
                    If Len(p) <= kwLen + 6 Then
                        If p = CStr(kw) Or InStr(1, p, CStr(kw), vbTextCompare) > 0 Then
                            If kwLen > bestLen Then
                                bestLen = kwLen
                                bestCol = c
                            End If
                        End If
                    End If
                Next kw
            End If
        End If
    Next c
    If bestCol > 0 Then
        NX_FindHeaderCol = bestCol
        Exit Function
    End If

    ' 2) cac dong phia tren dong tieu de (tieu de nhieu tang)
    bestCol = 0
    bestLen = 0
    For r = 1 To headerRow - 1
        For c = 1 To scanC
            p = NX_Plain(NX_GridText(grid, r, c))
            If Len(p) > 0 Then
                If InStr(1, p, "nhanxet", vbTextCompare) = 0 And _
                   InStr(1, p, "diemmanh", vbTextCompare) = 0 And _
                   InStr(1, p, "gopy", vbTextCompare) = 0 Then
                    For Each kw In keywords
                        kwLen = Len(CStr(kw))
                        If Len(p) <= kwLen + 6 Then
                            If p = CStr(kw) Or InStr(1, p, CStr(kw), vbTextCompare) > 0 Then
                                If kwLen > bestLen Then
                                    bestLen = kwLen
                                    bestCol = c
                                End If
                            End If
                        End If
                    Next kw
                End If
            End If
        Next c
    Next r

    NX_FindHeaderCol = bestCol
End Function

'-------------------------------------------------------------- nhan dien Lop
Public Function NX_DetectClass(ByVal ws As Object) As String
    Dim scanText As String, r As Long, c As Long
    Dim p As String, i As Long, ch As String, nextCh As String

    scanText = NX_Plain(ws.Name) & " "
    For r = 1 To 6
        For c = 1 To 10
            scanText = scanText & " " & NX_Plain(CStr(ws.Cells(r, c).Value))
        Next c
    Next r

    p = LCase$(scanText)

    For i = 1 To 5
        If InStr(p, "lop" & i) > 0 Or InStr(p, "lop " & i) > 0 Or InStr(p, "khoi" & i) > 0 Or InStr(p, "khoi " & i) > 0 Then
            NX_DetectClass = "Lop_" & i
            Exit Function
        End If
    Next i

    For i = 1 To Len(p) - 1
        ch = Mid$(p, i, 1)
        If ch >= "1" And ch <= "5" Then
            nextCh = Mid$(p, i + 1, 1)
            If (nextCh >= "a" And nextCh <= "z") Or nextCh = "/" Or nextCh = "." Or nextCh = "-" Then
                NX_DetectClass = "Lop_" & ch
                Exit Function
            End If
        End If
    Next i

    NX_DetectClass = ""
End Function

'-------------------------------------------------------------- nhan dien Dot danh gia
Public Function NX_DetectPeriod(ByVal ws As Object) As String
    Dim r As Long, c As Long, txt As String, code As String

    code = NX_PeriodCode(ws.Name)
    If Len(code) > 0 Then
        NX_DetectPeriod = code
        Exit Function
    End If

    For r = 1 To 8
        For c = 1 To 12
            txt = NX_CellText(ws.Cells(r, c))
            If Len(txt) > 0 Then
                code = NX_PeriodCode(txt)
                If Len(code) > 0 Then
                    NX_DetectPeriod = code
                    Exit Function
                End If
            End If
        Next c
    Next r

    NX_DetectPeriod = ""
End Function

Private Function NX_GridText(ByVal grid As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    If Not IsArray(grid) Then
        If r = 1 And c = 1 Then NX_GridText = NX_Clean(CStr(grid))
        Exit Function
    End If
    On Error Resume Next
    v = grid(r, c)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    NX_GridText = NX_Clean(CStr(v))
End Function

Public Function NX_IsTotalRow(ByVal s As String) As Boolean
    Dim p As String
    p = NX_Plain(s)
    Select Case p
    Case "tong", "tongcong", "cong", "ghichu", "hientai", "danhsach", "stt", "hovaten", "hoten", "ten"
        NX_IsTotalRow = True
    End Select
End Function
