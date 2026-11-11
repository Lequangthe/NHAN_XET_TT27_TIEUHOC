Option Explicit
'==============================================================
' ThisWorkbook - khoi tao khi add-in duoc nap
'==============================================================

Private Sub Workbook_Open()
    On Error Resume Next
    NX_EnsureDataDir

    If Dir$(NX_IniPath()) = "" Then NX_IniCreateDefault

    NX_IniReset
    NX_IniMigrate
    NX_AppendLog "Open add-in v" & NX_VERSION
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    NX_AppRestore
    On Error GoTo 0
End Sub
