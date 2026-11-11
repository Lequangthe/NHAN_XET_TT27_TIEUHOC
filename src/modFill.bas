Option Explicit
'==============================================================
' modFill - Dien nhan xet: do cot, chon dot/mon, ghi de / bo qua,
'           xem truoc, sao luu truoc khi ghi, hoan tac, nhat ky
'==============================================================
Public NX_LastReport As String
Private mUndoBuffer As Object   ' "Sheet!A1" -> gia tri goc
Private mUndoWbName As String

Public Function NX_Report() As String
    NX_Report = NX_LastReport
End Function

'-------------------------------------------------------------- cau hinh cot & dong
Public Sub NX_ConfigureColumnsAndRows()
    Dim curMode As String, curScore As String, curNote As String, curRow As String
    Dim ansMode As String, ansScore As String, ansNote As String, ansRow As String
    Dim t As String

    curMode = NX_GetFillMode()
    curScore = NX_Ini("CotDiemManual", "")
    curNote = NX_Ini("CotNhanXetManual", "")
    curRow = NX_Ini("HangDauTienManual", "")

    t = NX_T("cfg.modeText", "")
    If Len(t) = 0 Then t = "1 = Theo cot Diem/Xep loai" & vbCrLf & "2 = Xoay vong" & vbCrLf & "3 = Tat ca TOT" & vbCrLf & "4 = Tat ca HOAN THANH"
    ansMode = NX_InputBox(t, NX_T("cfg.modeTitle", "1. CHE DO DIEN"), curMode)
    If Len(ansMode) = 0 Then Exit Sub
    If Val(ansMode) >= 1 And Val(ansMode) <= 4 Then
        NX_SetIniKey "CheDoDienManual", Trim$(ansMode)
    Else
        NX_SetIniKey "CheDoDienManual", "1"
    End If

    If Trim$(NX_GetFillMode()) = "1" Then
        ansScore = NX_InputBox(NX_T("cfg.scoreText", "Cot diem (E, F hoac 5, 6). Trong = tu do."), NX_T("cfg.scoreTitle", "2. COT DIEM"), curScore)
        NX_SetIniKey "CotDiemManual", Trim$(ansScore)
    Else
        NX_SetIniKey "CotDiemManual", ""
    End If

    ansNote = NX_InputBox(NX_T("cfg.noteText", "Cot nhan xet (F, G hoac 6, 7). Trong = tu do."), NX_T("cfg.noteTitle", "3. COT NHAN XET"), curNote)
    NX_SetIniKey "CotNhanXetManual", Trim$(ansNote)

    ansRow = NX_InputBox(NX_T("cfg.rowText", "Hang hoc sinh dau tien (vi du 5). Trong = tu do."), NX_T("cfg.rowTitle", "4. HANG DAU TIEN"), curRow)
    NX_SetIniKey "HangDauTienManual", Trim$(ansRow)

    NX_RefreshRibbon

    NX_MsgBox NX_Fmt(NX_T("cfg.done", "Da luu cau hinh."), _
                  NX_GetFillMode(), _
                  IIf(Len(Trim$(ansScore)) > 0, ansScore, NX_T("cfg.auto", "tu dong")), _
                  IIf(Len(Trim$(ansNote)) > 0, ansNote, NX_T("cfg.auto", "tu dong")), _
                  IIf(Len(Trim$(ansRow)) > 0, ansRow, NX_T("cfg.auto", "tu dong"))), _
           vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
End Sub

'-------------------------------------------------------------- hoan tac
Public Sub NX_UndoLastFill()
    Dim wb As Object, ws As Object
    Dim k As Variant, parts() As String
    Dim countRestored As Long

    If mUndoBuffer Is Nothing Then GoTo NoData
    If mUndoBuffer.Count = 0 Then GoTo NoData

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb.Name <> mUndoWbName Then
        NX_MsgBox NX_T("undo.wrongBook", "File hien tai khong khop voi file da dien gan nhat."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    countRestored = 0
    For Each k In mUndoBuffer.Keys
        parts = Split(CStr(k), "!")
        If UBound(parts) = 1 Then
            Set ws = wb.Worksheets(parts(0))
            If Not ws Is Nothing Then
                ws.Range(parts(1)).Value = mUndoBuffer(k)
                countRestored = countRestored + 1
            End If
        End If
        Set ws = Nothing
    Next k

    Set mUndoBuffer = Nothing
    NX_MsgBox NX_Fmt(NX_T("undo.done", "Da hoan tac {0} o nhan xet."), countRestored), vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
    Exit Sub

NoData:
    NX_MsgBox NX_T("undo.none", "Chua co du lieu hoan tac nao."), vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
End Sub

'-------------------------------------------------------------- chay chinh
Public Sub NX_Run(ByVal previewOnly As Boolean, Optional ByVal noUi As Boolean = False)
    Dim wb As Object, ws As Object
    Dim lay As NX_Layout
    Dim recs() As String
    Dim n As Long, writeN As Long, skipN As Long
    Dim r As Long, firstR As Long, lastR As Long
    Dim scoreCol As Long, levelCol As Long
    Dim cols() As Long, nCols As Long, ci As Long
    Dim noteColsStr As String
    Dim manualScore As String, manualNote As String, manualRow As String
    Dim fillMode As String, fillModeDesc As String
    Dim selClass As String, selPeriod As String, selSubj As String
    Dim detClass As String, detPeriod As String, detSubj As String
    Dim lvl As String, sen As String, nm As String, scoreVal As String, tmp As String
    Dim targetCols As String
    Dim ans As VbMsgBoxResult
    Dim backupPath As String
    Dim ghiDe As Boolean

    On Error GoTo Fail
    NX_LastReport = ""

    If Application.Workbooks.Count = 0 Then
        If Not noUi Then NX_MsgBox NX_T("run.noBook", "Chua mo file Excel nao."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        NX_LastReport = "ERR no workbook"
        Exit Sub
    End If

    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb Is ThisWorkbook Then
        If Not noUi Then NX_MsgBox NX_T("run.noSheet", "Vui long mo file so diem can dien."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        NX_LastReport = "ERR addin active"
        Exit Sub
    End If

    NX_EnsureDataDir
    NX_IniReset
    NX_IniMigrate

    Set ws = wb.ActiveSheet
    If ws Is Nothing Then Exit Sub

    selClass = NX_GetSelectedClass()
    selPeriod = NX_GetSelectedPeriod()
    selSubj = NX_GetSelectedSubject()
    fillMode = NX_GetFillMode()

    '--------------------------------------------- nhan dien tu sheet
    detClass = NX_DetectClass(ws)
    detPeriod = NX_DetectPeriod(ws)
    detSubj = NX_SubjectCode(ws.Name)

    If Len(detClass) = 0 Then detClass = selClass
    If Len(detPeriod) = 0 Then detPeriod = selPeriod
    If Len(detSubj) = 0 Then detSubj = selSubj

    If StrComp(detClass, selClass, vbTextCompare) <> 0 Or _
       StrComp(detPeriod, selPeriod, vbTextCompare) <> 0 Or _
       StrComp(detSubj, selSubj, vbTextCompare) <> 0 Then

        If noUi Then
            selClass = detClass
            selPeriod = detPeriod
            selSubj = detSubj
            NX_SetIniKey "LopSelected", selClass
            NX_SetSelectedPeriod selPeriod
            NX_SetIniKey "MonSelected", selSubj
            NX_RefreshRibbon
        Else
            ans = NX_MsgBox(NX_Fmt(NX_T("detect.confirm", ""), _
                                NX_ClassName(detClass), NX_PeriodName(detPeriod), UCase$(NX_SubjectName(detSubj)), _
                                NX_ClassName(selClass), NX_PeriodName(selPeriod), UCase$(NX_SubjectName(selSubj))), _
                         vbYesNoCancel + vbQuestion, NX_T("detect.title", "Nhan dien tu sheet"))
            If ans = vbCancel Then Exit Sub
            If ans = vbYes Then
                selClass = detClass
                selPeriod = detPeriod
                selSubj = detSubj
                NX_SetIniKey "LopSelected", selClass
                NX_SetSelectedPeriod selPeriod
                NX_SetIniKey "MonSelected", selSubj
                NX_RefreshRibbon
            End If
        End If
    End If

    '--------------------------------------------- ngan hang mau
    NX_BankLoad True, selClass, selPeriod
    NX_BankEnsure selSubj

    If NX_BankLevelCount(selSubj, "T") + NX_BankLevelCount(selSubj, "H") + NX_BankLevelCount(selSubj, "C") = 0 Then
        If Not noUi Then
            NX_MsgBox NX_Fmt(NX_T("run.noBank", "Chua doc duoc mau nhan xet cho {0} - {1} - mon {2}."), _
                          NX_ClassName(selClass), NX_PeriodName(selPeriod), UCase$(NX_SubjectName(selSubj))), _
                   vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        End If
        NX_LastReport = "ERR no bank " & selClass & " " & selPeriod & " " & selSubj
        Exit Sub
    End If

    '--------------------------------------------- do layout sheet
    lay = NX_Detect(ws)
    If Not lay.Ok Then
        If Not noUi Then
            NX_MsgBox NX_Fmt(NX_T("run.noLayout", "Khong tim thay cot 'Nhan xet' trong sheet '{0}'.\n{1}"), ws.Name, lay.Warn), _
                   vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        End If
        NX_LastReport = "ERR layout " & ws.Name
        Exit Sub
    End If

    manualScore = Trim$(NX_Ini("CotDiemManual", ""))
    manualNote = Trim$(NX_Ini("CotNhanXetManual", ""))
    manualRow = Trim$(NX_Ini("HangDauTienManual", ""))
    ghiDe = (Trim$(NX_Ini("GhiDe", "0")) = "1")

    '--------------------------------------------- cot diem / muc dat
    scoreCol = 0
    levelCol = 0
    If fillMode = "1" Then
        fillModeDesc = NX_T("mode.desc1", "Theo cot Diem / Xep loai")
        If Len(manualScore) > 0 Then
            scoreCol = NX_ColLetterToNum(manualScore)
            levelCol = scoreCol
        Else
            scoreCol = lay.ScoreCol
            levelCol = lay.LevelCol
        End If

        If scoreCol = 0 And levelCol = 0 Then
            If noUi Then
                NX_LastReport = "ERR no score col " & ws.Name
                Exit Sub
            End If
            tmp = NX_InputBox(NX_T("run.askScoreCol", "Chua chon cot diem. Nhap E, F hoac 5, 6:"), NX_T("run.askScoreTitle", "CHON COT DIEM"), "")
            If Len(Trim$(tmp)) = 0 Then Exit Sub
            NX_SetIniKey "CotDiemManual", Trim$(tmp)
            scoreCol = NX_ColLetterToNum(Trim$(tmp))
            levelCol = scoreCol
        End If
    ElseIf fillMode = "2" Then
        fillModeDesc = NX_T("mode.desc2", "Xoay vong Tot & Hoan thanh")
    ElseIf fillMode = "3" Then
        fillModeDesc = NX_T("mode.desc3", "Tat ca muc TOT")
    Else
        fillModeDesc = NX_T("mode.desc4", "Tat ca muc HOAN THANH")
    End If

    '--------------------------------------------- cot nhan xet
    noteColsStr = NX_PickNoteCols(ws, lay, selPeriod, manualNote)
    If Len(noteColsStr) = 0 Then
        If Not noUi Then
            NX_MsgBox NX_Fmt(NX_T("run.noLayout", "Khong tim thay cot 'Nhan xet' trong sheet '{0}'.\n{1}"), ws.Name, lay.Warn), _
                   vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        End If
        NX_LastReport = "ERR note col " & ws.Name
        Exit Sub
    End If
    cols = NX_SplitCols(noteColsStr, nCols)

    '--------------------------------------------- hang bat dau / ket thuc
    If Len(manualRow) > 0 And Val(manualRow) > 0 Then
        firstR = CLng(Val(manualRow))
    Else
        firstR = lay.FirstRow
        If firstR < 2 Then firstR = lay.HeaderRow + 1
    End If

    lastR = lay.LastRow
    If lastR < firstR Then lastR = firstR
    If lay.NameCol > 0 Then
        tmp = CStr(ws.Cells(ws.Rows.Count, lay.NameCol).End(-4162).Row)
        If Val(tmp) > lastR Then lastR = CLng(Val(tmp))
    End If
    If lastR > firstR + 500 Then lastR = firstR + 500

    Randomize
    Set mUndoBuffer = CreateObject("Scripting.Dictionary")
    mUndoWbName = wb.Name

    ReDim recs(0 To 7, 0 To 0)
    n = -1
    writeN = 0
    skipN = 0

    For r = firstR To lastR
        nm = ""
        If lay.NameCol > 0 Then nm = NX_CellText(ws.Cells(r, lay.NameCol))
        If Len(nm) > 0 Then
            If Not NX_IsTotalRow(nm) Then

                Select Case fillMode
                Case "1"
                    scoreVal = ""
                    If levelCol > 0 Then scoreVal = NX_CellText(ws.Cells(r, levelCol))
                    If Len(scoreVal) = 0 Then
                        If scoreCol > 0 Then scoreVal = NX_CellText(ws.Cells(r, scoreCol))
                    End If
                    lvl = NX_LevelCode(scoreVal)
                Case "2"
                    If Rnd() < 0.7 Then lvl = "T" Else lvl = "H"
                Case "3"
                    lvl = "T"
                Case "4"
                    lvl = "H"
                Case Else
                    lvl = "H"
                End Select

                If Len(lvl) > 0 Then
                    targetCols = ""
                    For ci = 0 To nCols - 1
                        If ghiDe Or Len(NX_CellText(ws.Cells(r, cols(ci)))) = 0 Then
                            targetCols = targetCols & CStr(cols(ci)) & ","
                        Else
                            skipN = skipN + 1
                        End If
                    Next ci

                    If Len(targetCols) > 0 Then
                        targetCols = Left$(targetCols, Len(targetCols) - 1)
                        sen = NX_Draw(selSubj, lvl)
                        If Len(sen) > 0 Then
                            n = n + 1
                            ReDim Preserve recs(0 To 7, 0 To n)
                            recs(0, n) = ws.Name
                            recs(1, n) = CStr(n + 1)
                            recs(2, n) = nm
                            recs(3, n) = NX_LevelName(lvl)
                            recs(4, n) = NX_SubjectName(selSubj)
                            recs(5, n) = sen
                            recs(6, n) = targetCols
                            recs(7, n) = CStr(r)
                            writeN = writeN + NX_CountCols(targetCols)
                        End If
                    End If
                End If

            End If
        End If
    Next r

    If n < 0 Then
        If Not noUi Then
            NX_MsgBox NX_Fmt(NX_T("run.noStudent", "Khong tim thay hoc sinh nao hop le de dien.\n(Cot diem: {0} | Cot nhan xet: {1} | Hang dau: {2})"), _
                          IIf(scoreCol > 0, CStr(scoreCol), "-"), noteColsStr, CStr(firstR)), _
                   vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
        End If
        NX_LastReport = "OK " & selClass & " " & selPeriod & " " & selSubj & " write=0 skip=" & skipN
        Exit Sub
    End If

    If previewOnly Then
        NX_MakePreview wb, recs, n
        NX_LastReport = "PREVIEW " & selClass & " " & selPeriod & " " & selSubj & " write=" & writeN & " skip=" & skipN
        Exit Sub
    End If

    If Not noUi Then
        ans = NX_MsgBox(NX_Fmt(NX_T("run.confirm", "Dien nhan xet ngay?"), _
                            CStr(n + 1), NX_ClassName(selClass), NX_PeriodName(selPeriod), UCase$(NX_SubjectName(selSubj)), _
                            fillModeDesc, noteColsStr, CStr(skipN)), _
                     vbYesNo + vbQuestion, NX_T("app.title", "NX Tieu hoc TT27"))
        If ans <> vbYes Then Exit Sub
    End If

    backupPath = NX_BackupWorkbook(wb)

    NX_AppSave
    writeN = NX_Commit(wb, recs, n, skipN)
    NX_AppRestore

    NX_LastReport = "OK " & selClass & " " & selPeriod & " " & selSubj & " write=" & writeN & " skip=" & skipN
    NX_AppendLog "DIEN " & wb.Name & " | " & ws.Name & " | " & selClass & " " & selPeriod & " " & selSubj & " | write=" & writeN & " skip=" & skipN & " | backup=" & backupPath

    If Not noUi Then
        NX_MsgBox NX_Fmt(NX_T("run.done", "Xong! Da ghi {0} o nhan xet (bo qua {1} o da co chu).\n\n - Mau: {2} | {3} | MON {4}\n - Che do: {5}\n - Sao luu: {6}"), _
                      CStr(writeN), CStr(skipN), NX_ClassName(selClass), NX_PeriodName(selPeriod), UCase$(NX_SubjectName(selSubj)), _
                      fillModeDesc, IIf(Len(backupPath) > 0, backupPath, "-")), _
               vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    End If
    Exit Sub

Fail:
    On Error Resume Next
    NX_AppRestore
    NX_LastReport = "ERR " & Err.Description
    If Not noUi Then NX_MsgBox NX_Fmt(NX_T("run.error", "Loi: {0}"), Err.Description), vbCritical, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- chon cot nhan xet
Public Function NX_PickNoteCols(ByVal ws As Object, ByRef lay As NX_Layout, ByVal period As String, ByVal manual As String) As String
    Dim spec As String, cols() As Long, nCols As Long, i As Long
    Dim best As Long, bestScore As Long, c As Long, code As String, p As String

    manual = Trim$(manual)
    If Len(manual) > 0 Then
        c = NX_ColLetterToNum(manual)
        If c > 0 Then NX_PickNoteCols = CStr(c)
        Exit Function
    End If

    spec = LCase$(Trim$(NX_Ini("CotNhanXet", "auto")))
    If Len(spec) = 0 Then spec = "auto"

    If spec = "first" Then
        cols = NX_SplitCols(lay.NoteCols, nCols)
        If nCols > 0 Then NX_PickNoteCols = CStr(cols(0))
        Exit Function
    End If

    If spec = "all" Then
        NX_PickNoteCols = lay.NoteCols
        Exit Function
    End If

    If spec <> "auto" And spec <> "last" Then
        c = NX_ColLetterToNum(spec)
        If c > 0 Then
            NX_PickNoteCols = CStr(c)
            Exit Function
        End If
    End If

    ' auto: uu tien cot co tieu de dung dot dang chon, neu khong thi cot cuoi
    cols = NX_SplitCols(lay.NoteCols, nCols)
    If nCols = 0 Then Exit Function

    best = 0
    bestScore = 0
    For i = 0 To nCols - 1
        p = NX_CellText(ws.Cells(lay.HeaderRow, cols(i)))
        code = NX_PeriodCode(p)
        If Len(code) > 0 And StrComp(code, period, vbTextCompare) = 0 Then
            best = cols(i)
            Exit For
        End If
    Next i

    If best > 0 Then
        NX_PickNoteCols = CStr(best)
    Else
        NX_PickNoteCols = CStr(cols(nCols - 1))
    End If
End Function

'-------------------------------------------------------------- tien ich cot
Public Function NX_ColLetterToNum(ByVal colStr As String) As Long
    Dim i As Long, c As String, num As Long
    colStr = UCase$(Trim$(colStr))
    If Len(colStr) = 0 Then Exit Function

    If IsNumeric(colStr) Then
        NX_ColLetterToNum = CLng(Val(colStr))
        Exit Function
    End If

    num = 0
    For i = 1 To Len(colStr)
        c = Mid$(colStr, i, 1)
        If Asc(c) >= 65 And Asc(c) <= 90 Then
            num = num * 26 + (Asc(c) - 64)
        End If
    Next i
    NX_ColLetterToNum = num
End Function

Private Function NX_SplitCols(ByVal s As String, ByRef nCols As Long) As Long()
    Dim parts() As String, i As Long, k As Long
    Dim a() As Long, v As Long

    nCols = 0
    ReDim a(0 To 0)

    If Len(Trim$(s)) > 0 Then
        parts = Split(s, ",")
        ReDim a(0 To UBound(parts))
        k = 0
        For i = LBound(parts) To UBound(parts)
            v = NX_ColLetterToNum(Trim$(parts(i)))
            If v > 0 Then
                a(k) = v
                k = k + 1
            End If
        Next i
        If k > 0 Then
            ReDim Preserve a(0 To k - 1)
            nCols = k
        End If
    End If

    NX_SplitCols = a
End Function

Private Function NX_CountCols(ByVal s As String) As Long
    Dim i As Long, c As Long
    If Len(Trim$(s)) = 0 Then Exit Function
    c = 1
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "," Then c = c + 1
    Next i
    NX_CountCols = c
End Function

'-------------------------------------------------------------- sao luu file truoc khi ghi
Public Function NX_BackupWorkbook(ByVal wb As Object) As String
    Dim src As String, dst As String, base As String

    On Error Resume Next
    src = wb.FullName
    If Len(src) = 0 Then Exit Function
    If Dir$(src) = "" Then Exit Function

    NX_EnsureDataDir
    base = NX_FileBase(wb.Name)
    dst = NX_BackupDir() & "\" & base & "_" & NX_NowStamp() & ".xlsx"

    FileCopy src, dst
    If Err.Number <> 0 Then
        Err.Clear
        wb.SaveCopyAs dst
    End If
    If Err.Number = 0 Then NX_BackupWorkbook = dst
    Err.Clear
    On Error GoTo 0
End Function

'-------------------------------------------------------------- xem truoc
Private Sub NX_MakePreview(ByVal wb As Object, ByRef recs() As String, ByVal n As Long)
    Dim ws As Object, i As Long, j As Long, r As Long, hdr As Variant
    Dim oldAlerts As Boolean

    oldAlerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(NX_T("run.previewSheet", "NX_XEMTRUOC")).Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = NX_T("run.previewSheet", "NX_XEMTRUOC")

    hdr = Array(NX_T("preview.sheet", "Sheet"), NX_T("preview.stt", "STT"), NX_T("preview.name", "Ho va ten"), _
                NX_T("preview.level", "Muc"), NX_T("preview.subject", "Mon"), _
                NX_T("preview.sentence", "Cau nhan xet"), NX_T("preview.target", "O dich"))
    For j = 0 To UBound(hdr)
        ws.Cells(1, j + 1).Value = hdr(j)
    Next j

    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = -4108
    End With

    For i = 0 To n
        r = i + 2
        For j = 0 To 5
            ws.Cells(r, j + 1).Value = recs(j, i)
        Next j
        ws.Cells(r, 7).Value = NX_T("preview.row", "hang ") & recs(7, i) & " / cot " & recs(6, i)
    Next i

    ws.Columns.AutoFit
    ws.Activate
End Sub

'-------------------------------------------------------------- ghi that
Private Function NX_Commit(ByVal wb As Object, ByRef recs() As String, ByVal n As Long, ByRef skipN As Long) As Long
    Dim i As Long, w As Long, ci As Long, nCols As Long
    Dim ws As Object, cel As Object, cols() As Long
    Dim r As Long, c As Long, key As String

    For i = 0 To n
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(recs(0, i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            r = CLng(Val(recs(7, i)))
            cols = NX_SplitCols(recs(6, i), nCols)
            For ci = 0 To nCols - 1
                c = cols(ci)
                Set cel = ws.Cells(r, c)
                key = ws.Name & "!" & cel.Address(False, False)
                If Not mUndoBuffer.Exists(key) Then mUndoBuffer(key) = cel.Value

                If Trim$(NX_Ini("GhiDe", "0")) = "1" Or Len(NX_CellText(cel)) = 0 Then
                    cel.Value = recs(5, i)
                    w = w + 1
                Else
                    skipN = skipN + 1
                End If
            Next ci
        End If
        Set ws = Nothing
    Next i

    NX_Commit = w
End Function
