Option Explicit
'==============================================================
' modBank - Ngan hang cau nhan xet theo LOP / DOT / MON
'
'   Mau_Nhan_Xet\<Lop_n>\<GK1|HK1|GK2|CK2>\<Mon>.xlsx
'   Cot A = TOT, Cot B = HOAN THANH, Cot C = CAN CO GANG
'   Dong 1 la tieu de muc, cau bat dau tu dong 2.
'
'   Nap LUOI (lazy): chi mo file cua mon can dung.
'==============================================================
Private mBank As Object         ' "MON|T" -> String()
Private mSeq As Object          ' "MON|T" -> Long()
Private mPtr As Object          ' "MON|T" -> Long
Private mLastDrawn As Object    ' "MON|T" -> cau vua rut (chong lap lien ke)
Private mLoadedSubj As Object   ' "MON"   -> True
Private mLoadClass As String
Private mLoadPeriod As String
Private mInited As Boolean

'-------------------------------------------------------------- ma mon
Public Function NX_SubjectCode(ByVal s As String) As String
    Dim p As String, i As Long, best As String, bestLen As Long, a As Variant
    p = NX_Plain(s)
    If Len(p) = 0 Then Exit Function

    a = NX_AliasPairs()
    For i = LBound(a) To UBound(a) Step 2
        If p = a(i) Then NX_SubjectCode = a(i + 1): Exit Function
    Next i

    best = ""
    bestLen = 0
    For i = LBound(a) To UBound(a) Step 2
        If Len(a(i)) >= 4 Then
            If InStr(1, p, a(i), vbTextCompare) > 0 Then
                If Len(a(i)) > bestLen Then
                    bestLen = Len(a(i))
                    best = a(i + 1)
                End If
            End If
        End If
    Next i

    If bestLen > 0 Then NX_SubjectCode = best
End Function

Public Function NX_SubjectName(ByVal code As String) As String
    Select Case UCase$(code)
    Case "TOAN": NX_SubjectName = "To" & ChrW$(&H00E1) & "n"
    Case "TIENGVIET", "TV": NX_SubjectName = "Ti" & ChrW$(&H1EBF) & "ng Vi" & ChrW$(&H1EC7) & "t"
    Case "TNXH": NX_SubjectName = "T" & ChrW$(&H1EF1) & " nhi" & ChrW$(&H00EA) & "n v" & ChrW$(&H00E0) & " X" & ChrW$(&H00E3) & " h" & ChrW$(&H1ED9) & "i"
    Case "DAODUC": NX_SubjectName = ChrW$(&H0110) & ChrW$(&H1EA1) & "o " & ChrW$(&H0111) & ChrW$(&H1EE9) & "c"
    Case "HDTN": NX_SubjectName = "Ho" & ChrW$(&H1EA1) & "t " & ChrW$(&H0111) & ChrW$(&H1ED9) & "ng tr" & ChrW$(&H1EA3) & "i nghi" & ChrW$(&H1EC7) & "m"
    Case "AMNHAC": NX_SubjectName = ChrW$(&H00C2) & "m nh" & ChrW$(&H1EA1) & "c"
    Case "MYTHUAT": NX_SubjectName = "M" & ChrW$(&H1EF9) & " thu" & ChrW$(&H1EAD) & "t"
    Case "GDTC": NX_SubjectName = "Gi" & ChrW$(&H00E1) & "o d" & ChrW$(&H1EE5) & "c th" & ChrW$(&H1EC3) & " ch" & ChrW$(&H1EA5) & "t"
    Case "TINHOC": NX_SubjectName = "Tin h" & ChrW$(&H1ECD) & "c"
    Case "CONGNGHE": NX_SubjectName = "C" & ChrW$(&H00F4) & "ng ngh" & ChrW$(&H1EC7)
    Case "TIENGANH": NX_SubjectName = "Ti" & ChrW$(&H1EBF) & "ng Anh"
    Case "KHOAHOC": NX_SubjectName = "Khoa h" & ChrW$(&H1ECD) & "c"
    Case "LSDL": NX_SubjectName = "L" & ChrW$(&H1ECB) & "ch s" & ChrW$(&H1EED) & " v" & ChrW$(&H00E0) & " " & ChrW$(&H0110) & ChrW$(&H1ECB) & "a l" & ChrW$(&H00ED)
    Case Else: NX_SubjectName = code
    End Select
End Function

Private Function NX_AliasPairs() As Variant
    Dim s As String, items() As String, i As Long, p As Long
    Dim a() As String

    s = "hoatdongtrainghiem:HDTN,hdtn:HDTN," & _
        "tunhienvaxahoi:TNXH,tnxh:TNXH," & _
        "lichsuvadialy:LSDL,lichsuvalialy:LSDL,lsdl:LSDL," & _
        "tiengviet:TiengViet,tv:TiengViet," & _
        "congnghe:CongNghe,mythuat:MyThuat,khoahoc:KhoaHoc,amnhac:AmNhac," & _
        "tinhoc:TinHoc,daoduc:DaoDuc,dd:DaoDuc," & _
        "gdtc:GDTC,theduc:GDTC,tienganh:TiengAnh,ta:TiengAnh," & _
        "toanhoc:Toan,toan:Toan"

    items = Split(s, ",")
    ReDim a(0 To UBound(items) * 2 + 1)
    For i = LBound(items) To UBound(items)
        p = InStr(items(i), ":")
        If p > 1 Then
            a(i * 2) = Left$(items(i), p - 1)
            a(i * 2 + 1) = Mid$(items(i), p + 1)
        End If
    Next i
    NX_AliasPairs = a
End Function

'-------------------------------------------------------------- khoi tao / nap
Private Sub NX_BankInit()
    If mInited Then Exit Sub
    Set mBank = CreateObject("Scripting.Dictionary")
    Set mSeq = CreateObject("Scripting.Dictionary")
    Set mPtr = CreateObject("Scripting.Dictionary")
    Set mLastDrawn = CreateObject("Scripting.Dictionary")
    Set mLoadedSubj = CreateObject("Scripting.Dictionary")
    mInited = True
End Sub

Private Sub NX_BankClear()
    Set mBank = CreateObject("Scripting.Dictionary")
    Set mSeq = CreateObject("Scripting.Dictionary")
    Set mPtr = CreateObject("Scripting.Dictionary")
    Set mLastDrawn = CreateObject("Scripting.Dictionary")
    Set mLoadedSubj = CreateObject("Scripting.Dictionary")
End Sub

Public Sub NX_BankLoad(Optional ByVal force As Boolean = False, Optional ByVal classFolder As String = "", Optional ByVal period As String = "")
    Dim c As String, p As String

    c = Trim$(classFolder)
    If Len(c) = 0 Then c = NX_GetSelectedClass()
    p = NX_PeriodCode(period)
    If Len(p) = 0 Then p = NX_GetSelectedPeriod()

    NX_BankInit

    If force Or StrComp(c, mLoadClass, vbTextCompare) <> 0 Or StrComp(p, mLoadPeriod, vbTextCompare) <> 0 Then
        NX_BankClear
        mLoadClass = c
        mLoadPeriod = p
    End If
End Sub

Public Function NX_BankCount() As Long
    NX_BankInit
    NX_BankCount = mBank.Count
End Function

Public Function NX_BankInfo() As String
    NX_BankInfo = mLoadClass & " " & mLoadPeriod & " groups=" & NX_BankCount()
End Function

Public Function NX_BankLevelCount(ByVal subj As String, ByVal lvl As String) As Long
    Dim v As Variant, key As String

    NX_BankEnsure subj
    key = UCase$(Trim$(subj)) & "|" & UCase$(Trim$(lvl))
    If Not mBank.Exists(key) Then Exit Function
    v = mBank(key)
    NX_BankLevelCount = UBound(v) - LBound(v) + 1
End Function

' Mo file mau cua 1 mon (chi 1 lan cho moi mon)
Public Sub NX_BankEnsure(ByVal subj As String)
    Dim key As String, file As String
    Dim wb As Object, ws As Object
    Dim lastR As Long, r As Long, c As Long
    Dim txt As String, lvl As String
    Dim buf(1 To 3) As String
    Dim v As Variant, arr() As String

    NX_BankInit
    key = UCase$(Trim$(subj))
    If Len(key) = 0 Then Exit Sub
    If mLoadedSubj.Exists(key) Then Exit Sub
    mLoadedSubj(key) = True

    file = NX_TemplateFilePath(mLoadClass, mLoadPeriod, subj)
    If Dir$(file) = "" Then Exit Sub

    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=file, UpdateLinks:=0, ReadOnly:=True)
    If Err.Number <> 0 Or wb Is Nothing Then
        Err.Clear
        Set wb = Nothing
        On Error GoTo 0
        Exit Sub
    End If

    Set ws = wb.Worksheets(1)
    lastR = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   ' xlUp
    If lastR > 2000 Then lastR = 2000

    If lastR >= 2 Then
        v = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 3)).Value
        If IsArray(v) Then
            For r = 1 To lastR - 1
                For c = 1 To 3
                    txt = NX_Clean(CStr(v(r, c)))
                    If Len(txt) > 0 Then
                        If Len(buf(c)) = 0 Then
                            buf(c) = txt
                        Else
                            buf(c) = buf(c) & vbLf & txt
                        End If
                    End If
                Next c
            Next r
        Else
            txt = NX_Clean(CStr(v))
            If Len(txt) > 0 Then buf(1) = txt
        End If
    End If

    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing
    Err.Clear
    On Error GoTo 0

    For c = 1 To 3
        lvl = Mid$("THC", c, 1)
        If Len(buf(c)) > 0 Then
            arr = Split(buf(c), vbLf)
            mBank(key & "|" & lvl) = arr
        End If
    Next c
End Sub

'-------------------------------------------------------------- muc dat
Public Function NX_LevelCode(ByVal s As String) As String
    Dim p As String, num As Double
    p = NX_Plain(s)
    If Len(p) = 0 Then Exit Function

    If IsNumeric(p) Then
        num = Val(p)
        If num >= 9# Then
            NX_LevelCode = "T"
        ElseIf num >= 5# Then
            NX_LevelCode = "H"
        Else
            NX_LevelCode = "C"
        End If
        Exit Function
    End If

    Select Case p
    Case "t", "tot", "htt", "hoanthanhtot", "xuatxac", "gioi"
        NX_LevelCode = "T"
    Case "c", "cht", "chuahoanthanh", "cancogang", "yeu", "kem"
        NX_LevelCode = "C"
    Case "h", "ht", "hoanthanh", "dat", "kha", "tb", "trungbinh"
        NX_LevelCode = "H"
    Case Else
        If InStr(1, p, "hoanthanhtot", vbTextCompare) > 0 Or InStr(1, p, "xuatxac", vbTextCompare) > 0 Then
            NX_LevelCode = "T"
        ElseIf InStr(1, p, "chuahoanthanh", vbTextCompare) > 0 Or InStr(1, p, "cancogang", vbTextCompare) > 0 Then
            NX_LevelCode = "C"
        ElseIf InStr(1, p, "hoanthanh", vbTextCompare) > 0 Or InStr(1, p, "dat", vbTextCompare) > 0 Then
            NX_LevelCode = "H"
        End If
    End Select
End Function

Public Function NX_LevelName(ByVal code As String) As String
    Select Case UCase$(code)
    Case "T": NX_LevelName = "T" & ChrW$(&H1ED0) & "T"
    Case "C": NX_LevelName = "C" & ChrW$(&H1EA6) & "N C" & ChrW$(&H1ED0) & " G" & ChrW$(&H1EAE) & "NG"
    Case Else: NX_LevelName = "HO" & ChrW$(&H00C0) & "N TH" & ChrW$(&H00C0) & "NH"
    End Select
End Function

Private Function NX_LevelChain(ByVal lvl As String) As Variant
    Select Case UCase$(lvl)
    Case "T": NX_LevelChain = Array("T", "H", "C")
    Case "C": NX_LevelChain = Array("C", "H", "T")
    Case Else: NX_LevelChain = Array("H", "T", "C")
    End Select
End Function

'-------------------------------------------------------------- rut cau ngau nhien chong trung
Public Function NX_Draw(ByVal subj As String, ByVal lvl As String) As String
    Dim subsTry As Variant, lv As Variant
    Dim i As Long, j As Long, key As String, k As Variant, s As String

    If Len(lvl) = 0 Then Exit Function
    NX_BankEnsure subj

    s = UCase$(Trim$(subj))
    If Len(s) = 0 Then s = "TOAN"
    subsTry = Array(s, "TOAN")
    lv = NX_LevelChain(lvl)

    For i = LBound(subsTry) To UBound(subsTry)
        For j = LBound(lv) To UBound(lv)
            key = CStr(subsTry(i)) & "|" & CStr(lv(j))
            If mBank.Exists(key) Then
                NX_Draw = NX_Take(key)
                Exit Function
            End If
        Next j
    Next i

    ' cuu canh: bat ky mon nao da nap co cung muc
    For j = LBound(lv) To UBound(lv)
        For Each k In mBank.Keys
            If Right$(CStr(k), 2) = "|" & CStr(lv(j)) Then
                NX_Draw = NX_Take(CStr(k))
                Exit Function
            End If
        Next k
    Next j
End Function

Private Function NX_Take(ByVal key As String) As String
    Dim v As Variant, n As Long, p As Long
    Dim seq() As Long, r As Long, res As String

    v = mBank(key)
    n = UBound(v) - LBound(v) + 1
    If n < 1 Then Exit Function

    If Not mSeq.Exists(key) Then
        mSeq(key) = NX_Shuffle(n)
        mPtr(key) = 0
    End If

    seq = mSeq(key)
    p = CLng(mPtr(key))
    If p >= n Then
        seq = NX_Shuffle(n)
        mSeq(key) = seq
        p = 0
    End If

    r = LBound(v) + seq(p)
    If r > UBound(v) Then r = LBound(v)
    res = CStr(v(r))

    If mLastDrawn.Exists(key) Then
        If res = CStr(mLastDrawn(key)) And n > 1 Then
            p = (p + 1) Mod n
            r = LBound(v) + seq(p)
            res = CStr(v(r))
        End If
    End If

    mPtr(key) = p + 1
    mLastDrawn(key) = res
    NX_Take = res
End Function

Private Function NX_Shuffle(ByVal n As Long) As Long()
    Dim a() As Long, i As Long, j As Long, t As Long
    ReDim a(0 To n - 1)
    For i = 0 To n - 1
        a(i) = i
    Next i
    For i = n - 1 To 1 Step -1
        j = Int(Rnd() * (i + 1))
        t = a(i)
        a(i) = a(j)
        a(j) = t
    Next i
    NX_Shuffle = a
End Function
