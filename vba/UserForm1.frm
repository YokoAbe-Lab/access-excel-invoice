VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "お待ちください…"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   5985
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=== UserForm1（点滅アニメーション）===
Private blinkState As Boolean
Private nextTime As Double

Private Sub UserForm_Activate()
    Me.lblProgress.Caption = "処理中です…"
    Me.lblProgress.Visible = True
    Me.lblProgress.ForeColor = vbBlack
    blinkState = True
    ScheduleNextBlink
End Sub

Public Sub ScheduleNextBlink()
    nextTime = Now + TimeValue("00:00:01") ' 1秒ごとに切り替え
    Application.OnTime nextTime, "BlinkProgressLabel"
End Sub

Public Sub StopBlink()
    On Error Resume Next
    Application.OnTime EarliestTime:=nextTime, Procedure:="BlinkProgressLabel", Schedule:=False
    Me.lblProgress.Visible = False
End Sub
Private Sub UserForm_Initialize()
    Dim w As Integer, h As Integer
    Dim x As Integer, y As Integer

    ' Excelウィンドウの中央にUserFormを表示（1モニター想定）
    w = Me.Width
    h = Me.Height

    With Application
        x = .Left + (.Width / 2) - (w / 2)
        y = .Top + (.Height / 2) - (h / 2)
    End With

    Me.StartUpPosition = 0 ' 手動位置指定
    Me.Left = x
    Me.Top = y
End Sub

