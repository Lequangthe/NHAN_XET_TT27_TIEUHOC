Option Explicit
'==============================================================
' modUtil - ham tien ich: chuan hoa chuoi, doc o, ten file
'==============================================================

' Chuan hoa: bo dau tieng Viet, chu thuong, chi giu a-z va 0-9
Public Function NX_Plain(ByVal s As String) As String
    Dim i As Long, n As Long, r As String, c As String
    s = NX_Clean(s)
    n = Len(s)
    If n = 0 Then Exit Function
    r = ""
    For i = 1 To n
        c = NX_BaseChar(Mid$(s, i, 1))
        If Len(c) > 0 Then r = r & c
    Next i
    NX_Plain = r
End Function

' Lam sach: trim, bo khoang trang lien tiep, bo ky tu dieu khien
Public Function NX_Clean(ByVal s As String) As String
    Dim r As String
    If Len(s) = 0 Then Exit Function
    r = Replace(s, ChrW$(160), " ")
    r = Replace(r, vbTab, " ")
    r = Replace(r, vbCr, " ")
    r = Replace(r, vbLf, " ")
    r = Replace(r, ChrW$(&HFEFF), "")
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop
    NX_Clean = Trim$(r)
End Function

Private Function NX_BaseChar(ByVal c As String) As String
    Dim q As Long
    q = AscW(c)
    Select Case q
    Case 65 To 90, 97 To 122, 48 To 57
        NX_BaseChar = LCase$(ChrW$(q))
    Case &H00C0, &H00C1, &H00C2, &H00C3, &H00C4, &H00C5, &H00E0, &H00E1, &H00E2, &H00E3, &H00E4, &H00E5, &H0100, &H0101, &H0102, &H0103, &H0104, &H0105, &H01CD, &H01CE, &H01DE, &H01DF, &H01E0, &H01E1, &H01FA, &H01FB, &H0200, &H0201, &H0202, &H0203, &H0226, &H0227, &H1E00, &H1E01, &H1EA0, &H1EA1, &H1EA2, &H1EA3, &H1EA4, &H1EA5, &H1EA6, &H1EA7, &H1EA8, &H1EA9, &H1EAA, &H1EAB, &H1EAC, &H1EAD, &H1EAE, &H1EAF, &H1EB0, &H1EB1, &H1EB2, &H1EB3, &H1EB4, &H1EB5, &H1EB6, &H1EB7, &H212B
        NX_BaseChar = "a"
    Case &H010E, &H010F, &H0110, &H0111, &H1E0A, &H1E0B, &H1E0C, &H1E0D, &H1E0E, &H1E0F, &H1E10, &H1E11, &H1E12, &H1E13
        NX_BaseChar = "d"
    Case &H00C8, &H00C9, &H00CA, &H00CB, &H00E8, &H00E9, &H00EA, &H00EB, &H0112, &H0113, &H0114, &H0115, &H0116, &H0117, &H0118, &H0119, &H011A, &H011B, &H0204, &H0205, &H0206, &H0207, &H0228, &H0229, &H1E14, &H1E15, &H1E16, &H1E17, &H1E18, &H1E19, &H1E1A, &H1E1B, &H1E1C, &H1E1D, &H1EB8, &H1EB9, &H1EBA, &H1EBB, &H1EBC, &H1EBD, &H1EBE, &H1EBF, &H1EC0, &H1EC1, &H1EC2, &H1EC3, &H1EC4, &H1EC5, &H1EC6, &H1EC7
        NX_BaseChar = "e"
    Case &H00CC, &H00CD, &H00CE, &H00CF, &H00EC, &H00ED, &H00EE, &H00EF, &H0128, &H0129, &H012A, &H012B, &H012C, &H012D, &H012E, &H012F, &H0130, &H01CF, &H01D0, &H0208, &H0209, &H020A, &H020B, &H1E2C, &H1E2D, &H1E2E, &H1E2F, &H1EC8, &H1EC9, &H1ECA, &H1ECB
        NX_BaseChar = "i"
    Case &H00D2, &H00D3, &H00D4, &H00D5, &H00D6, &H00F2, &H00F3, &H00F4, &H00F5, &H00F6, &H014C, &H014D, &H014E, &H014F, &H0150, &H0151, &H01A0, &H01A1, &H01D1, &H01D2, &H01EA, &H01EB, &H01EC, &H01ED, &H020C, &H020D, &H020E, &H020F, &H022A, &H022B, &H022C, &H022D, &H022E, &H022F, &H0230, &H0231, &H1E4C, &H1E4D, &H1E4E, &H1E4F, &H1E50, &H1E51, &H1E52, &H1E53, &H1ECC, &H1ECD, &H1ECE, &H1ECF, &H1ED0, &H1ED1, &H1ED2, &H1ED3, &H1ED4, &H1ED5, &H1ED6, &H1ED7, &H1ED8, &H1ED9, &H1EDA, &H1EDB, &H1EDC, &H1EDD, &H1EDE, &H1EDF, &H1EE0, &H1EE1, &H1EE2, &H1EE3
        NX_BaseChar = "o"
    Case &H00D9, &H00DA, &H00DB, &H00DC, &H00F9, &H00FA, &H00FB, &H00FC, &H0168, &H0169, &H016A, &H016B, &H016C, &H016D, &H016E, &H016F, &H0170, &H0171, &H0172, &H0173, &H01AF, &H01B0, &H01D3, &H01D4, &H01D5, &H01D6, &H01D7, &H01D8, &H01D9, &H01DA, &H01DB, &H01DC, &H0214, &H0215, &H0216, &H0217, &H1E72, &H1E73, &H1E74, &H1E75, &H1E76, &H1E77, &H1E78, &H1E79, &H1E7A, &H1E7B, &H1EE4, &H1EE5, &H1EE6, &H1EE7, &H1EE8, &H1EE9, &H1EEA, &H1EEB, &H1EEC, &H1EED, &H1EEE, &H1EEF, &H1EF0, &H1EF1
        NX_BaseChar = "u"
    Case &H00DD, &H00FD, &H00FF, &H0176, &H0177, &H0178, &H0232, &H0233, &H1E8E, &H1E8F, &H1E99, &H1EF2, &H1EF3, &H1EF4, &H1EF5, &H1EF6, &H1EF7, &H1EF8, &H1EF9
        NX_BaseChar = "y"
    Case Else
        NX_BaseChar = ""
    End Select
End Function

' Doc text cua 1 o, an toan voi loi / ngay thang
Public Function NX_CellText(ByVal cel As Object) As String
    Dim v As Variant
    On Error Resume Next
    v = cel.Value
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        NX_CellText = Format$(v, "dd/mm/yyyy")
    ElseIf VarType(v) = vbBoolean Then
        NX_CellText = IIf(v, "TRUE", "FALSE")
    Else
        NX_CellText = NX_Clean(CStr(v))
    End If
End Function

Public Function NX_FileBase(ByVal f As String) As String
    Dim p As Long
    p = InStrRev(f, ".")
    If p > 0 Then NX_FileBase = Left$(f, p - 1) Else NX_FileBase = f
End Function

Public Function NX_FileExt(ByVal f As String) As String
    Dim p As Long
    p = InStrRev(f, ".")
    If p > 0 Then NX_FileExt = Mid$(f, p + 1) Else NX_FileExt = "xlsx"
End Function

' So phan tu cua mang (an toan voi Array() rong)
Public Function NX_ArrCount(ByRef a As Variant) As Long
    On Error Resume Next
    NX_ArrCount = UBound(a) - LBound(a) + 1
    If Err.Number <> 0 Then
        Err.Clear
        NX_ArrCount = 0
    End If
    On Error GoTo 0
End Function

' Dau thoi gian cho ten file backup: yyyymmdd_hhnnss
Public Function NX_NowStamp() As String
    NX_NowStamp = Format$(Now, "yyyymmdd_hhnnss")
End Function
