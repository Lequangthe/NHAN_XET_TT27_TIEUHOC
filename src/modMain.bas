Option Explicit
'==============================================================
' modMain - Nut bam Ribbon + nhan dien tu sheet + thong tin
'==============================================================
Public Const NX_VERSION As String = "3.0 (TT27 - 4 dot)"
Public mRibbon As Object

Public Sub Ribbon_OnLoad(ribbon As Object)
    Set mRibbon = ribbon
End Sub

Public Sub NX_RefreshRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- nhan tren Ribbon
Public Sub GetActiveTemplateLabel(control As Object, ByRef label As Variant)
    Dim m As String, modeTag As String

    On Error Resume Next
    m = NX_GetFillMode()
    Select Case m
    Case "1": modeTag = NX_T("mode.tag1", "CD1 (COT DIEM)")
    Case "2": modeTag = NX_T("mode.tag2", "CD2 (XOAY VONG)")
    Case "3": modeTag = NX_T("mode.tag3", "CD3 (MUC TOT)")
    Case "4": modeTag = NX_T("mode.tag4", "CD4 (MUC HT)")
    Case Else: modeTag = NX_T("mode.tag1", "CD1")
    End Select

    label = NX_Fmt(NX_T("label.current", "MAU: {0} | {1} | MON {2} | {3}"), _
                   UCase$(NX_ClassName(NX_GetSelectedClass())), _
                   NX_PeriodName(NX_GetSelectedPeriod()), _
                   UCase$(NX_SubjectName(NX_GetSelectedSubject())), _
                   modeTag)
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- dien & hoan tac
Public Sub DienNhanXet(control As Object)
    On Error Resume Next
    NX_Run False, False
    On Error GoTo 0
End Sub

Public Sub XemTruocNhanXet(control As Object)
    On Error Resume Next
    NX_Run True, False
    On Error GoTo 0
End Sub

Public Sub HoanTacNhanXet(control As Object)
    On Error Resume Next
    NX_UndoLastFill
    On Error GoTo 0
End Sub

Public Sub CauHinhCotDiem(control As Object)
    On Error Resume Next
    NX_ConfigureColumnsAndRows
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- che do dien
Public Sub ChonCheDoDien(control As Object)
    Dim curMode As String, ansMode As String, modeDesc As String, t As String

    On Error Resume Next
    curMode = NX_GetFillMode()

    t = NX_T("mode.text", "")
    If Len(t) = 0 Then
        t = "1 = Theo cot Diem / Xep loai" & vbCrLf & _
            "2 = Xoay vong ngau nhien Tot & Hoan thanh" & vbCrLf & _
            "3 = Tat ca deu MUC TOT" & vbCrLf & _
            "4 = Tat ca deu MUC HOAN THANH"
    End If

    ansMode = NX_InputBox(t, NX_T("mode.title", "CHON CHE DO DIEN NHAN XET"), curMode)
    If Len(ansMode) = 0 Then Exit Sub
    If Val(ansMode) >= 1 And Val(ansMode) <= 4 Then
        NX_SetIniKey "CheDoDienManual", Trim$(ansMode)
    Else
        NX_SetIniKey "CheDoDienManual", "1"
    End If

    NX_RefreshRibbon

    Select Case NX_GetFillMode()
    Case "1": modeDesc = NX_T("mode.desc1", "Theo cot Diem / Xep loai")
    Case "2": modeDesc = NX_T("mode.desc2", "Xoay vong Tot & Hoan thanh")
    Case "3": modeDesc = NX_T("mode.desc3", "Tat ca muc TOT")
    Case "4": modeDesc = NX_T("mode.desc4", "Tat ca muc HOAN THANH")
    Case Else: modeDesc = ""
    End Select

    NX_MsgBox NX_Fmt(NX_T("mode.done", "Da thiet lap Che do dien {0}: {1}"), NX_GetFillMode(), modeDesc), _
           vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- chon Lop -> Dot -> Mon
Public Sub ChonLopVaMon(control As Object)
    Dim classes As Variant, subs As Variant
    Dim s As String, ans As String, i As Long, n As Long
    Dim lopFolder As String, per As String, monKey As String
    Dim dfltPeriod As String

    On Error Resume Next

    classes = NX_ListClasses()
    n = NX_ArrCount(classes)
    If n = 0 Then
        NX_MsgBox NX_T("pick.noClass", "Khong tim thay thu muc lop nao trong thu muc mau."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    s = ""
    For i = 0 To n - 1
        s = s & CStr(i + 1) & ". " & NX_ClassName(CStr(classes(i))) & vbCrLf
    Next i

    ans = NX_InputBox(NX_Fmt(NX_T("pick.class", "CHON KHOI LOP (nhap so hoac ten):\n{0}"), s), _
                   NX_T("pick.classTitle", "1. CHON KHOI LOP"), _
                   Right$(NX_GetSelectedClass(), 1))
    If Len(ans) = 0 Then Exit Sub

    lopFolder = NX_PickFromList(classes, ans)
    If Len(lopFolder) = 0 Then
        NX_MsgBox NX_T("pick.invalidClass", "Lop khong hop le."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    '---- dot danh gia
    s = "1. " & NX_PeriodName("GK1") & vbCrLf & _
        "2. " & NX_PeriodName("HK1") & vbCrLf & _
        "3. " & NX_PeriodName("GK2") & vbCrLf & _
        "4. " & NX_PeriodName("CK2")
    dfltPeriod = CStr(NX_PeriodIndex(NX_GetSelectedPeriod()))

    ans = NX_InputBox(NX_Fmt(NX_T("pick.period", "CHON DOT DANH GIA (nhap 1-4):\n{0}"), s), _
                   NX_T("pick.periodTitle", "2. CHON DOT DANH GIA"), dfltPeriod)
    If Len(ans) = 0 Then Exit Sub

    per = NX_PeriodCode(ans)
    If Len(per) = 0 Then
        If IsNumeric(ans) Then
            Select Case CLng(Val(ans))
            Case 1: per = "GK1"
            Case 2: per = "HK1"
            Case 3: per = "GK2"
            Case 4: per = "CK2"
            End Select
        End If
    End If
    If Len(per) = 0 Then
        NX_MsgBox NX_T("pick.invalidPeriod", "Dot danh gia khong hop le."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    '---- mon hoc (quet theo thu muc mau)
    subs = NX_ListSubjects(lopFolder, per)
    n = NX_ArrCount(subs)
    If n = 0 Then
        NX_MsgBox NX_Fmt(NX_T("pick.noSubject", "Khong co file mau cho {0} - {1}.\nThu muc: {2}"), _
                      NX_ClassName(lopFolder), NX_PeriodName(per), NX_TemplateDir()), _
               vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    s = ""
    For i = 0 To n - 1
        s = s & CStr(i + 1) & ". " & NX_SubjectName(CStr(subs(i))) & vbCrLf
    Next i

    ans = NX_InputBox(NX_Fmt(NX_T("pick.subject", "CHON MON HOC (nhap so hoac ten):\n{0}"), s), _
                   NX_T("pick.subjectTitle", "3. CHON MON HOC"), NX_GetSelectedSubject())
    If Len(ans) = 0 Then Exit Sub

    monKey = NX_MatchSubject(subs, ans)
    If Len(monKey) = 0 Then monKey = CStr(subs(0))

    NX_SetIniKey "LopSelected", lopFolder
    NX_SetSelectedPeriod per
    NX_SetIniKey "MonSelected", monKey

    NX_BankLoad True, lopFolder, per
    NX_BankEnsure monKey
    NX_RefreshRibbon

    NX_MsgBox NX_Fmt(NX_T("pick.done", "Da thiet lap:\n - Lop: {0}\n - Dot: {1}\n - Mon: {2}\n\nFile mau: {3}"), _
                  NX_ClassName(lopFolder), NX_PeriodName(per), NX_SubjectName(monKey), _
                  NX_TemplateFilePath(lopFolder, per, monKey)), _
           vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- nhan dien tu sheet dang mo
Public Sub NXNhanDienTuSheet(control As Object)
    Dim ws As Object, wb As Object
    Dim detClass As String, detPeriod As String, detSubj As String

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb Is ThisWorkbook Then
        NX_MsgBox NX_T("run.noSheet", "Vui long mo file so diem can dien."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    Set ws = wb.ActiveSheet
    If ws Is Nothing Then Exit Sub

    detClass = NX_DetectClass(ws)
    detPeriod = NX_DetectPeriod(ws)
    detSubj = NX_SubjectCode(ws.Name)

    If Len(detClass) = 0 And Len(detPeriod) = 0 And Len(detSubj) = 0 Then
        NX_MsgBox NX_T("detect.none", "Khong nhan dien duoc Lop / Dot / Mon tu sheet dang mo."), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
        Exit Sub
    End If

    If Len(detClass) = 0 Then detClass = NX_GetSelectedClass()
    If Len(detPeriod) = 0 Then detPeriod = NX_GetSelectedPeriod()
    If Len(detSubj) = 0 Then detSubj = NX_GetSelectedSubject()

    NX_SetIniKey "LopSelected", detClass
    NX_SetSelectedPeriod detPeriod
    NX_SetIniKey "MonSelected", detSubj

    NX_BankLoad True, detClass, detPeriod
    NX_BankEnsure detSubj
    NX_RefreshRibbon

    NX_MsgBox NX_Fmt(NX_T("detect.done", "Da chuyen sang mau:\n - Lop: {0}\n - Dot: {1}\n - Mon: {2}"), _
                  NX_ClassName(detClass), NX_PeriodName(detPeriod), NX_SubjectName(detSubj)), _
           vbInformation, NX_T("detect.title", "Nhan dien tu sheet")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- du lieu mau Excel
Public Sub MoFileMauMon(control As Object)
    On Error Resume Next
    NX_EnsureDataDir
    NX_OpenCurrentTemplate NX_GetSelectedClass(), NX_GetSelectedPeriod(), NX_GetSelectedSubject()
    On Error GoTo 0
End Sub

Public Sub MoThuMucMauExcel(control As Object)
    On Error Resume Next
    NX_EnsureDataDir
    NX_OpenFolder NX_TemplateDir()
    On Error GoTo 0
End Sub

Public Sub ChonThuMucMauCustom(control As Object)
    On Error Resume Next
    NX_SetCustomTemplateDir
    NX_RefreshRibbon
    On Error GoTo 0
End Sub

Public Sub LamMoiNganHang(control As Object)
    On Error Resume Next
    NX_EnsureDataDir
    NX_IniReset
    NX_IniMigrate
    NX_BankLoad True, NX_GetSelectedClass(), NX_GetSelectedPeriod()
    NX_BankEnsure NX_GetSelectedSubject()
    NX_RefreshRibbon

    NX_MsgBox NX_Fmt(NX_T("data.reloaded", "Da nap lai ngan hang mau.\n\n - Dang chon: {0} | {1} | {2}\n - Thu muc: {3}\n - So nhom cau: {4}"), _
                  NX_ClassName(NX_GetSelectedClass()), NX_PeriodName(NX_GetSelectedPeriod()), _
                  NX_SubjectName(NX_GetSelectedSubject()), NX_TemplateDir(), CStr(NX_BankCount())), _
           vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- tro giup & thong tin
Public Sub HuongDanAddin(control As Object)
    Dim s As String
    On Error Resume Next
    s = NX_T("help.text", "")
    If Len(s) = 0 Then s = "NX Tieu hoc - huong dan su dung."
    NX_MsgBox s, vbInformation, NX_T("help.title", "Tro giup - NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

Public Sub ThongTinAddin(control As Object)
    Dim s As String
    On Error Resume Next
    NX_IniReset
    NX_IniMigrate
    s = NX_Fmt(NX_T("info.text", ""), NX_VERSION, NX_ClassName(NX_GetSelectedClass()), NX_PeriodName(NX_GetSelectedPeriod()), _
               NX_SubjectName(NX_GetSelectedSubject()), NX_TemplateDir(), NX_DataDir())
    If Len(s) = 0 Then s = "NX Tieu hoc v" & NX_VERSION
    NX_MsgBox s, vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- tien ich
Private Function NX_PickFromList(ByRef list As Variant, ByVal ans As String) As String
    Dim i As Long, n As Long, idx As Long
    Dim p As String, v As String

    n = NX_ArrCount(list)
    If n = 0 Then Exit Function

    ans = Trim$(ans)
    If Len(ans) = 0 Then Exit Function

    If IsNumeric(ans) Then
        idx = CLng(Val(ans))
        If idx >= 1 And idx <= n Then
            NX_PickFromList = CStr(list(idx - 1))
            Exit Function
        End If
    End If

    p = NX_Plain(ans)
    For i = 0 To n - 1
        v = CStr(list(i))
        If StrComp(v, ans, vbTextCompare) = 0 Then
            NX_PickFromList = v
            Exit Function
        End If
        If Len(p) > 0 Then
            If InStr(1, NX_Plain(v), p, vbTextCompare) > 0 Then
                NX_PickFromList = v
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub NX_OpenFolder(ByVal p As String)
    On Error Resume Next
    If Dir$(p, vbDirectory) = "" Then MkDir p
    Shell "explorer.exe """ & p & """", vbNormalFocus
    On Error GoTo 0
End Sub
