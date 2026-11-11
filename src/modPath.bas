Option Explicit
'==============================================================
' modPath - Thu muc du lieu, DOT DANH GIA (GK1/HK1/GK2/CK2),
'           thong diep UTF-8, cai dat INI, doc/ghi UTF-8
'==============================================================
Private mIni As Object
Private mMsg As Object

'-------------------------------------------------------------- thu muc du lieu
Public Function NX_DataDir() As String
    Dim p As String
    If Len(ThisWorkbook.Path) > 0 Then
        NX_DataDir = ThisWorkbook.Path & "\NXTiengViet_data"
        Exit Function
    End If

    p = Environ$("USERPROFILE") & "\Downloads\NXTiengViet_data"
    If Dir$(p, vbDirectory) <> "" Then
        NX_DataDir = p
        Exit Function
    End If

    NX_DataDir = Environ$("APPDATA") & "\NXTiengViet_data"
End Function

Public Function NX_TemplateDir() As String
    Dim customFolder As String
    customFolder = NX_Ini("ThuMucMau", "")
    If Len(customFolder) > 0 Then
        On Error Resume Next
        If Dir$(customFolder, vbDirectory) <> "" Then
            NX_TemplateDir = customFolder
            Exit Function
        End If
        On Error GoTo 0
    End If

    NX_TemplateDir = NX_DataDir() & "\Mau_Nhan_Xet"
End Function

Public Function NX_MsgPath() As String
    NX_MsgPath = NX_DataDir() & "\thongdiep.txt"
End Function

Public Function NX_IniPath() As String
    NX_IniPath = NX_DataDir() & "\caidat.ini"
End Function

Public Function NX_LogPath() As String
    NX_LogPath = NX_DataDir() & "\nhat_ky.txt"
End Function

Public Function NX_BackupDir() As String
    NX_BackupDir = NX_DataDir() & "\backup"
End Function

Public Sub NX_EnsureDataDir()
    Dim p As String, pMau As String, pBk As String
    p = NX_DataDir()
    pMau = NX_TemplateDir()
    pBk = NX_BackupDir()
    On Error Resume Next
    If Dir$(p, vbDirectory) = "" Then MkDir p
    If Dir$(pMau, vbDirectory) = "" Then MkDir pMau
    If Dir$(pBk, vbDirectory) = "" Then MkDir pBk
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- DOT DANH GIA
' GK1 = giua ky 1 | HK1 = cuoi ky 1 | GK2 = giua ky 2 | CK2 = cuoi ky 2
Public Function NX_PeriodCode(ByVal s As String) As String
    Dim p As String
    Dim yy As Long

    p = NX_Plain(s)
    If Len(p) = 0 Then Exit Function
    ' Bo nam hoc (2020-2030) de tranh nham so 2 trong nam thanh dot 2
    For yy = 2030 To 2020 Step -1
        p = Replace(p, CStr(yy), "")
    Next yy
    If Len(p) = 0 Then Exit Function

    Select Case p
    Case "gk1", "gki1", "giuaky1", "giuaki1", "giuahocky1", "giuahki1"
        NX_PeriodCode = "GK1": Exit Function
    Case "gk2", "gki2", "giuaky2", "giuaki2", "giuahocky2", "giuahki2"
        NX_PeriodCode = "GK2": Exit Function
    Case "hk1", "hki1", "hki", "hk i", "hocky1", "hocki1", "hockyi", "cuoiky1", "cuoiki1", "ck1", "cki1"
        NX_PeriodCode = "HK1": Exit Function
    Case "hk2", "hki2", "hkii", "hk ii", "hocky2", "hocki2", "hockyii", "cuoiky2", "cuoiki2", "ck2", "cki2"
        NX_PeriodCode = "CK2": Exit Function
    End Select

    If InStr(1, p, "giua", vbTextCompare) > 0 Then
        If InStr(1, p, "2", vbTextCompare) > 0 Or InStr(1, p, "kyii", vbTextCompare) > 0 Or Right$(p, 2) = "ii" Then
            NX_PeriodCode = "GK2"
        Else
            NX_PeriodCode = "GK1"
        End If
        Exit Function
    End If

    If InStr(1, p, "cuoi", vbTextCompare) > 0 Then
        If InStr(1, p, "2", vbTextCompare) > 0 Or InStr(1, p, "kyii", vbTextCompare) > 0 Or Right$(p, 2) = "ii" Then
            NX_PeriodCode = "CK2"
        Else
            NX_PeriodCode = "HK1"
        End If
        Exit Function
    End If

    If InStr(1, p, "ck2", vbTextCompare) > 0 Then NX_PeriodCode = "CK2": Exit Function
    If InStr(1, p, "ck1", vbTextCompare) > 0 Then NX_PeriodCode = "HK1": Exit Function
    If InStr(1, p, "hk2", vbTextCompare) > 0 Or InStr(1, p, "hocky2", vbTextCompare) > 0 Then NX_PeriodCode = "CK2": Exit Function
    If InStr(1, p, "hk1", vbTextCompare) > 0 Or InStr(1, p, "hocky1", vbTextCompare) > 0 Then NX_PeriodCode = "HK1": Exit Function
End Function

Public Function NX_PeriodName(ByVal code As String) As String
    Select Case UCase$(code)
    Case "GK1": NX_PeriodName = "Gi" & ChrW$(&H1EEF) & "a k" & ChrW$(&H1EF3) & " 1 (GK1)"
    Case "HK1": NX_PeriodName = "Cu" & ChrW$(&H1ED1) & "i k" & ChrW$(&H1EF3) & " 1 (HK1)"
    Case "GK2": NX_PeriodName = "Gi" & ChrW$(&H1EEF) & "a k" & ChrW$(&H1EF3) & " 2 (GK2)"
    Case "CK2": NX_PeriodName = "Cu" & ChrW$(&H1ED1) & "i k" & ChrW$(&H1EF3) & " 2 (CK2)"
    Case Else: NX_PeriodName = code
    End Select
End Function

' Ten hien thi cua khoi lop: Lop_3 -> LOP 3 (dung tieng Viet)
' Luu tru trong file (ini/thumuc) van la ma Lop_n.
Public Function NX_ClassName(ByVal folder As String) As String
    Dim p As Long, num As String, i As Long, ch As String
    folder = Trim$(folder)
    num = ""
    For i = Len(folder) To 1 Step -1
        ch = Mid$(folder, i, 1)
        If Asc(ch) >= 48 And Asc(ch) <= 57 Then
            num = ch & num
        Else
            Exit For
        End If
    Next i
    If Len(num) > 0 Then
        NX_ClassName = "L" & ChrW$(&H1EDB) & "p " & num
    Else
        NX_ClassName = folder
    End If
End Function

Public Function NX_PeriodIndex(ByVal code As String) As Long
    Select Case UCase$(code)
    Case "GK1": NX_PeriodIndex = 1
    Case "HK1": NX_PeriodIndex = 2
    Case "GK2": NX_PeriodIndex = 3
    Case "CK2": NX_PeriodIndex = 4
    Case Else: NX_PeriodIndex = 2
    End Select
End Function

Public Function NX_PeriodFallbacks(ByVal code As String) As Variant
    Select Case UCase$(code)
    Case "GK1": NX_PeriodFallbacks = Array("GK1", "HK1", "GK2", "CK2")
    Case "HK1": NX_PeriodFallbacks = Array("HK1", "GK1", "CK2", "GK2")
    Case "GK2": NX_PeriodFallbacks = Array("GK2", "CK2", "HK1", "GK1")
    Case "CK2": NX_PeriodFallbacks = Array("CK2", "GK2", "HK1", "GK1")
    Case Else:  NX_PeriodFallbacks = Array("HK1", "GK1", "CK2", "GK2")
    End Select
End Function

Private Function NX_PeriodDirNames(ByVal per As String) As Variant
    Select Case UCase$(per)
    Case "HK1": NX_PeriodDirNames = Array("HK1", "CK1")
    Case "CK2": NX_PeriodDirNames = Array("CK2", "HK2")
    Case Else:  NX_PeriodDirNames = Array(UCase$(per))
    End Select
End Function

'-------------------------------------------------------------- lua chon hien tai
Public Function NX_GetSelectedClass() As String
    Dim c As String
    c = Trim$(NX_Ini("LopSelected", "Lop_3"))
    If Len(c) = 0 Then c = "Lop_3"
    NX_GetSelectedClass = c
End Function

Public Function NX_GetSelectedPeriod() As String
    Dim p As String
    p = NX_PeriodCode(NX_Ini("DotSelected", ""))
    If Len(p) > 0 Then NX_GetSelectedPeriod = p: Exit Function
    p = NX_PeriodCode(NX_Ini("HocKySelected", ""))
    If Len(p) > 0 Then NX_GetSelectedPeriod = p: Exit Function
    NX_GetSelectedPeriod = "HK1"
End Function

Public Sub NX_SetSelectedPeriod(ByVal per As String)
    Dim p As String
    p = NX_PeriodCode(per)
    If Len(p) = 0 Then p = "HK1"
    NX_SetIniKey "DotSelected", p
    If p = "CK2" Then
        NX_SetIniKey "HocKySelected", "HK2"
    Else
        NX_SetIniKey "HocKySelected", "HK1"
    End If
End Sub

' giu ten cu cho tuong thich (tra ve HK1/HK2)
Public Function NX_GetSelectedSemester() As String
    If NX_GetSelectedPeriod() = "CK2" Then
        NX_GetSelectedSemester = "HK2"
    Else
        NX_GetSelectedSemester = "HK1"
    End If
End Function

Public Function NX_GetSelectedSubject() As String
    Dim m As String
    m = Trim$(NX_Ini("MonSelected", "Toan"))
    If Len(m) = 0 Then m = "Toan"
    NX_GetSelectedSubject = m
End Function

Public Function NX_GetFillMode() As String
    Dim mode As String
    mode = Trim$(NX_Ini("CheDoDienManual", "1"))
    If Len(mode) = 0 Then mode = "1"
    NX_GetFillMode = mode
End Function

'-------------------------------------------------------------- duong dan file mau
Public Function NX_SubjectFileName(ByVal key As String) As String
    Dim c As String
    If Len(key) = 0 Then
        NX_SubjectFileName = "Toan"
        Exit Function
    End If
    If InStr(key, "\") = 0 And InStr(key, "/") = 0 Then
        c = NX_SubjectCode(key)
        If Len(c) > 0 Then
            NX_SubjectFileName = c
            Exit Function
        End If
    End If
    NX_SubjectFileName = key
End Function

Public Function NX_TemplateFilePath(ByVal classFolder As String, ByVal period As String, ByVal subjectKey As String) As String
    Dim p As String, i As Long, j As Long
    Dim chain As Variant, dirs As Variant
    Dim per As String

    If Len(classFolder) = 0 Then classFolder = NX_GetSelectedClass()
    If Len(period) = 0 Then period = NX_GetSelectedPeriod()
    If Len(subjectKey) = 0 Then subjectKey = NX_GetSelectedSubject()
    subjectKey = NX_SubjectFileName(subjectKey)

    chain = NX_PeriodFallbacks(period)
    For i = LBound(chain) To UBound(chain)
        per = CStr(chain(i))
        dirs = NX_PeriodDirNames(per)
        For j = LBound(dirs) To UBound(dirs)
            p = NX_TemplateDir() & "\" & classFolder & "\" & CStr(dirs(j)) & "\" & subjectKey & ".xlsx"
            If Dir$(p) <> "" Then
                NX_TemplateFilePath = p
                Exit Function
            End If
        Next j
    Next i

    NX_TemplateFilePath = NX_TemplateDir() & "\" & classFolder & "\" & period & "\" & subjectKey & ".xlsx"
End Function

' Danh sach mon co file mau (quet theo thu muc, tu thich ung khi them/bot mon)
Public Function NX_ListSubjects(ByVal classFolder As String, ByVal period As String) As Variant
    Dim d As Object, chain As Variant, per As String
    Dim p As String, f As String, name As String
    Dim ks As Variant, arr() As String, i As Long, n As Long, j As Long, tmp As String

    Set d = CreateObject("Scripting.Dictionary")
    If Len(classFolder) = 0 Then classFolder = NX_GetSelectedClass()
    If Len(period) = 0 Then period = NX_GetSelectedPeriod()

    chain = NX_PeriodFallbacks(period)
    For i = LBound(chain) To UBound(chain)
        per = CStr(chain(i))
        p = NX_TemplateDir() & "\" & classFolder & "\" & per & "\*.xlsx"
        f = Dir$(p)
        Do While Len(f) > 0
            name = NX_FileBase(f)
            If Len(name) > 0 Then
                If Not d.Exists(LCase$(name)) Then d(LCase$(name)) = name
            End If
            f = Dir$
        Loop
        If d.Count > 0 Then Exit For
    Next i

    If d.Count = 0 Then
        NX_ListSubjects = Array()
        Exit Function
    End If

    ReDim arr(0 To d.Count - 1)
    ks = d.Keys
    n = 0
    For i = LBound(ks) To UBound(ks)
        arr(n) = CStr(d(ks(i)))
        n = n + 1
    Next i

    ' sap xep A-Z
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    NX_ListSubjects = arr
End Function

Public Function NX_ListClasses() As Variant
    Dim d As Object, f As String
    Dim ks As Variant, arr() As String, i As Long, n As Long, j As Long, tmp As String
    Dim base As String

    Set d = CreateObject("Scripting.Dictionary")
    base = NX_TemplateDir()
    f = Dir$(base & "\Lop_*", vbDirectory)
    Do While Len(f) > 0
        If Left$(f, 1) <> "." Then
            If Not d.Exists(LCase$(f)) Then d(LCase$(f)) = f
        End If
        f = Dir$
    Loop

    If d.Count = 0 Then
        NX_ListClasses = Array("Lop_1", "Lop_2", "Lop_3", "Lop_4", "Lop_5")
        Exit Function
    End If

    ReDim arr(0 To d.Count - 1)
    ks = d.Keys
    n = 0
    For i = LBound(ks) To UBound(ks)
        arr(n) = CStr(d(ks(i)))
        n = n + 1
    Next i
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    NX_ListClasses = arr
End Function

' Tim ten mon trong danh sach theo so thu tu hoac ten/ma mon
Public Function NX_MatchSubject(ByVal subs As Variant, ByVal ans As String) As String
    Dim i As Long, idx As Long, code As String, p As String
    Dim v As String

    ans = Trim$(ans)
    If Len(ans) = 0 Then Exit Function

    If IsNumeric(ans) Then
        idx = CLng(Val(ans))
        i = LBound(subs) + idx - 1
        If i >= LBound(subs) And i <= UBound(subs) Then
            NX_MatchSubject = CStr(subs(i))
            Exit Function
        End If
    End If

    code = NX_SubjectCode(ans)
    For i = LBound(subs) To UBound(subs)
        v = CStr(subs(i))
        If StrComp(v, ans, vbTextCompare) = 0 Then
            NX_MatchSubject = v
            Exit Function
        End If
        If Len(code) > 0 Then
            If StrComp(NX_SubjectCode(v), code, vbTextCompare) = 0 Then
                NX_MatchSubject = v
                Exit Function
            End If
        End If
    Next i

    p = NX_Plain(ans)
    If Len(p) >= 3 Then
        For i = LBound(subs) To UBound(subs)
            v = CStr(subs(i))
            If InStr(1, NX_Plain(v), p, vbTextCompare) > 0 Then
                NX_MatchSubject = v
                Exit Function
            End If
        Next i
    End If
End Function

Public Sub NX_OpenCurrentTemplate(ByVal classFolder As String, ByVal semester As String, ByVal subjectKey As String)
    Dim p As String
    NX_EnsureDataDir
    p = NX_TemplateFilePath(classFolder, semester, subjectKey)
    If Dir$(p) <> "" Then
        On Error Resume Next
        Application.Workbooks.Open p
        On Error GoTo 0
    Else
        NX_MsgBox NX_Fmt(NX_T("tpl.notFound", "Khong tim thay file mau Excel:\n{0}"), p), vbExclamation, NX_T("app.title", "NX Tieu hoc TT27")
    End If
End Sub

Public Sub NX_SetCustomTemplateDir()
    Dim fd As Object, selectedPath As String
    On Error Resume Next
    Set fd = Application.FileDialog(4)
    If fd Is Nothing Then Exit Sub

    fd.Title = NX_T("tpl.pickFolder", "CHON THU MUC CHUA MAU EXCEL (MAU_NHAN_XET)")
    fd.InitialFileName = NX_TemplateDir()

    If fd.Show = -1 Then
        selectedPath = fd.SelectedItems(1)
        If Len(selectedPath) > 0 Then
            NX_SetIniKey "ThuMucMau", selectedPath
            NX_MsgBox NX_Fmt(NX_T("tpl.folderSet", "Da thiet lap thu muc mau:\n{0}"), selectedPath), vbInformation, NX_T("app.title", "NX Tieu hoc TT27")
        End If
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- thong diep UTF-8
Public Function NX_T(ByVal key As String, Optional ByVal dflt As String = "") As String
    If mMsg Is Nothing Then Set mMsg = NX_LoadMsg()
    If mMsg.Exists(key) Then
        NX_T = CStr(mMsg(key))
    Else
        NX_T = dflt
    End If
End Function

Public Sub NX_TReset()
    Set mMsg = Nothing
End Sub

Public Function NX_Fmt(ByVal s As String, ParamArray args() As Variant) As String
    Dim i As Long
    For i = LBound(args) To UBound(args)
        s = Replace(s, "{" & CStr(i) & "}", CStr(args(i)))
    Next i
    NX_Fmt = s
End Function

Private Function NX_LoadMsg() As Object
    Dim d As Object, s As String, lines() As String, i As Long
    Dim ln As String, p As Long, k As String, v As String

    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    d.CompareMode = 1
    On Error GoTo 0

    s = NX_ReadUTF8(NX_MsgPath())
    If Len(s) = 0 Then
        Set NX_LoadMsg = d
        Exit Function
    End If

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    lines = Split(s, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" Then
                p = InStr(ln, "=")
                If p > 1 Then
                    k = LCase$(Trim$(Left$(ln, p - 1)))
                    v = Trim$(Mid$(ln, p + 1))
                    v = Replace(v, "\n", vbCrLf)
                    If Not d.Exists(k) Then d(k) = v
                End If
            End If
        End If
    Next i

    Set NX_LoadMsg = d
End Function

'-------------------------------------------------------------- doc / ghi UTF-8
Public Function NX_ReadUTF8(ByVal path As String) As String
    Dim st As Object, s As String, c0 As Long
    On Error Resume Next
    If Dir$(path) = "" Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    s = st.ReadText(-1)
    st.Close
    Set st = Nothing
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Len(s) > 0 Then
        c0 = AscW(Left$(s, 1))
        If c0 = &HFEFF Or c0 = -257 Then s = Mid$(s, 2)
    End If
    NX_ReadUTF8 = s
End Function

Public Sub NX_WriteUTF8(ByVal path As String, ByVal s As String)
    Dim st As Object
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.SaveToFile path, 2
    st.Close
    Set st = Nothing
    On Error GoTo 0
End Sub

Public Sub NX_AppendLog(ByVal s As String)
    Dim old As String
    On Error Resume Next
    old = NX_ReadUTF8(NX_LogPath())
    If Len(old) > 200000 Then old = ""
    NX_WriteUTF8 NX_LogPath(), old & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & s & vbCrLf
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- cai dat ini
Public Function NX_Ini(ByVal key As String, ByVal dflt As String) As String
    NX_LoadIni
    If mIni Is Nothing Then NX_Ini = dflt: Exit Function
    If mIni.Exists(LCase$(key)) Then
        NX_Ini = mIni(LCase$(key))
    Else
        NX_Ini = dflt
    End If
End Function

Public Sub NX_SetIniKey(ByVal key As String, ByVal val As String)
    NX_LoadIni
    If mIni Is Nothing Then Set mIni = CreateObject("Scripting.Dictionary")
    mIni(LCase$(key)) = val

    Dim k As Variant, s As String
    s = "; ============================================================" & vbCrLf & _
        ";  NX Tieu hoc - CAI DAT (TT27)" & vbCrLf & _
        "; ============================================================" & vbCrLf
    For Each k In mIni.Keys
        s = s & CStr(k) & "=" & CStr(mIni(k)) & vbCrLf
    Next k
    NX_WriteUTF8 NX_IniPath(), s
End Sub

Public Sub NX_IniReset()
    Set mIni = Nothing
End Sub

Private Sub NX_LoadIni()
    Dim s As String, lines() As String, i As Long, ln As String, p As Long
    If Not mIni Is Nothing Then Exit Sub
    Set mIni = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    mIni.CompareMode = 1
    On Error GoTo 0
    s = NX_ReadUTF8(NX_IniPath())
    If Len(s) = 0 Then Exit Sub
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    lines = Split(s, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> "#" And Left$(ln, 1) <> ";" And Left$(ln, 1) <> "[" Then
                p = InStr(ln, "=")
                If p > 1 Then mIni(LCase$(Trim$(Left$(ln, p - 1)))) = Trim$(Mid$(ln, p + 1))
            End If
        End If
    Next i
End Sub

Public Sub NX_IniCreateDefault()
    Dim s As String
    s = "; ============================================================" & vbCrLf & _
        ";  NX Tieu hoc - CAI DAT (TT27)" & vbCrLf & _
        ";  GhiDe=0: bo qua o da co noi dung | 1: ghi de" & vbCrLf & _
        ";  CotNhanXet: auto/last | first | all | so cot (vd 7)" & vbCrLf & _
        ";  DotSelected: GK1 | HK1 | GK2 | CK2" & vbCrLf & _
        "; ============================================================" & vbCrLf & _
        "GhiDe=0" & vbCrLf & _
        "CotNhanXet=auto" & vbCrLf & _
        "MucMacDinh=H" & vbCrLf & _
        "LopSelected=Lop_3" & vbCrLf & _
        "DotSelected=HK1" & vbCrLf & _
        "HocKySelected=HK1" & vbCrLf & _
        "MonSelected=Toan" & vbCrLf & _
        "CheDoDienManual=1" & vbCrLf & _
        "CotDiemManual=" & vbCrLf & _
        "CotNhanXetManual=" & vbCrLf & _
        "HangDauTienManual=" & vbCrLf & _
        "CauHinhVersion=3" & vbCrLf
    NX_WriteUTF8 NX_IniPath(), s
    NX_IniReset
End Sub

' Nang cap cai dat tu ban cu (HK1/HK2 -> 4 dot; hang dau 6 -> tu dong do)
Public Sub NX_IniMigrate()
    Dim v As String
    NX_IniReset

    If Trim$(NX_Ini("CauHinhVersion", "")) <> "3" Then
        If Len(Trim$(NX_Ini("DotSelected", ""))) = 0 Then
            v = NX_PeriodCode(NX_Ini("HocKySelected", "HK1"))
            If Len(v) = 0 Then v = "HK1"
            NX_SetIniKey "DotSelected", v
        End If

        If Trim$(NX_Ini("HangDauTienManual", "")) = "6" Then
            NX_SetIniKey "HangDauTienManual", ""
        End If

        NX_SetIniKey "CauHinhVersion", "3"
    End If

    ' bo sung khoa thieu (file cai dat cu bi rut gon)
    If Not mIni.Exists("ghide") Then NX_SetIniKey "GhiDe", "0"
    If Len(Trim$(NX_Ini("CotNhanXet", ""))) = 0 Then NX_SetIniKey "CotNhanXet", "auto"
    If Len(Trim$(NX_Ini("MucMacDinh", ""))) = 0 Then NX_SetIniKey "MucMacDinh", "H"
End Sub
