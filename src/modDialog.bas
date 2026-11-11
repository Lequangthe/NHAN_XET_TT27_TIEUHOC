Option Explicit
'==============================================================
' modDialog - Hop thoai Unicode (UserForm Segoe UI)
' VBA MsgBox/InputBox ANSI khong hien duoc tieng Viet tren
' Windows ACP 1252 (hien dau ?). Form rieng font Segoe UI
' hien dung tieng Viet. UI noi dung lay tu thongdiep.txt
' (UTF-8) qua NX_T. Module chi giu ASCII khi build.
'==============================================================

' Thay MsgBox: UserForm Unicode, ket qua VbMsgBoxResult
Public Function NX_MsgBox(ByVal txtPrompt As String, _
                          Optional ByVal Buttons As Long = 0, _
                          Optional ByVal txtTitle As String = "") As Long
    Dim frm As New frmNXMsg
    On Error Resume Next
    frm.Caption = "Nhap lieu"
    If Len(txtTitle) > 0 Then
        frm.lblMsg.Caption = txtTitle & vbCrLf & vbCrLf & txtPrompt
    Else
        frm.lblMsg.Caption = txtPrompt
    End If
    frm.ShowMsg txtPrompt, Buttons
    NX_MsgBox = frm.MsgResult
    Unload frm
    Set frm = Nothing
    On Error GoTo 0
End Function

' Thay InputBox: UserForm Unicode (Segoe UI) hien tieng Viet dung.
' Cancel tra ve "".
Public Function NX_InputBox(ByVal txtPrompt As String, _
                            ByVal txtTitle As String, _
                            ByVal txtDef As String) As String
    Dim frm As New frmNXInput
    On Error Resume Next
    frm.Caption = "Nhap lieu"
    frm.lblPrompt.Caption = txtTitle & vbCrLf & vbCrLf & txtPrompt
    frm.txtInput.Text = txtDef
    frm.Show vbModal
    If frm.Cancelled Then
        NX_InputBox = ""
    Else
        NX_InputBox = frm.InputValue
    End If
    Unload frm
    Set frm = Nothing
    On Error GoTo 0
End Function
