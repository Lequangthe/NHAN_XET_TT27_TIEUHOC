VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNXInput 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmNXInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNXInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public InputValue As String
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = False
    InputValue = ""
    lblPrompt.Caption = ""
    txtInput.Text = ""
    lblPrompt.WordWrap = True
    lblPrompt.Font.Name = "Segoe UI"
    lblPrompt.Font.Size = 11
    txtInput.Font.Name = "Segoe UI"
    txtInput.Font.Size = 11
    Me.Width = 590
    Me.Height = 360
End Sub

Private Sub cmdOK_Click()
    Cancelled = False
    InputValue = txtInput.Text
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub
