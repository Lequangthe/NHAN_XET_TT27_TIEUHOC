VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNXMsg 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmNXMsg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNXMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MsgResult As Long
Public MsgButtons As Long

Private Sub UserForm_Initialize()
    MsgResult = 0
    lblMsg.Caption = ""
    lblMsg.WordWrap = True
    lblMsg.Font.Name = "Segoe UI"
    lblMsg.Font.Size = 11
    Me.Width = 590
    Me.Height = 340
    cmdOK.Visible = False
    cmdYes.Visible = False
    cmdNo.Visible = False
    cmdCancel.Visible = False
End Sub

Public Sub ShowMsg(ByVal prompt As String, ByVal Buttons As Long)
    MsgButtons = Buttons
    lblMsg.Caption = prompt
    ' 0=OKOnly, 1=OKCancel, 3=YesNoCancel, 4=YesNo
    If Buttons = 0 Or Buttons = 64 Or Buttons = 48 Or Buttons = 16 Or Buttons = 32 Or Buttons = 64 Or Buttons = 48 Then
        cmdOK.Visible = True
    Else
        cmdYes.Visible = True
        cmdNo.Visible = True
    End If
    If Buttons = 1 Or Buttons = 3 Or Buttons = 1 Or Buttons = 3 Then
        cmdCancel.Visible = True
    End If
    Me.Show vbModal
End Sub

Private Sub cmdOK_Click()
    MsgResult = 1
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    MsgResult = 6
    Me.Hide
End Sub

Private Sub cmdNo_Click()
    MsgResult = 7
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    MsgResult = 2
    Me.Hide
End Sub
