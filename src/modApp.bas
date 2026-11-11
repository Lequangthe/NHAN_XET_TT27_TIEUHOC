Option Explicit
'==============================================================
' modApp - luu / khoi phuc trang thai ung dung (khong ro ri)
'==============================================================
Private mScreen As Boolean
Private mCalc As Long
Private mAlerts As Boolean
Private mEvents As Boolean
Private mSaved As Boolean

Public Sub NX_AppSave()
    On Error Resume Next
    mScreen = Application.ScreenUpdating
    mCalc = Application.Calculation
    mAlerts = Application.DisplayAlerts
    mEvents = Application.EnableEvents
    mSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = -4135            ' xlCalculationManual
    On Error GoTo 0
End Sub

Public Sub NX_AppRestore()
    If Not mSaved Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = mScreen
    Application.Calculation = mCalc
    Application.DisplayAlerts = mAlerts
    Application.EnableEvents = mEvents
    Application.StatusBar = False
    On Error GoTo 0
    mSaved = False
End Sub

Public Sub NX_Status(ByVal s As String)
    On Error Resume Next
    Application.StatusBar = s
    On Error GoTo 0
End Sub
