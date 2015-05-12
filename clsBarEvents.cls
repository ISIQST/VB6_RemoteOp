VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MotionEvents As BarCont.clsMotion

Dim TestingLastBar As Boolean
Dim CurHeadList() As String

Public Enum HeadType
    UP = 1
    DOWN = 0
End Enum

Private Sub Class_Initialize()
    Set MotionEvents = BarDriver.AutomationEvents
End Sub

Private Sub MotionEvents_LoadBarInitiate(ByVal TargetBar As Integer, AbortRtn As Boolean)
Dim traynum As Integer
Dim barnum As Integer
    ReDim CurHeadList(0)            'clear serial numbers
    'traynum = TargetBar \ BarDriver.TrayInfo.BarsPerTray
'    barnum = TargetBar Mod BarDriver.TrayInfo.BarsPerTray
'    If Not LotINFO.IsBarPresent(traynum, barnum) Then           'this is how to skip a bar from the tray
'        AbortRtn = True
'    End If
'    If TargetBar + 1 = MotionEvents.TotalBars Then
'        TestingLastBar = True
'    Else
'        TestingLastBar = False
'    End If
End Sub

Private Sub MotionEvents_EnableHeadsTerminate(HeadEnableAr() As Boolean)
Dim i As Integer
    If (MotionEvents.CurrentState = 0) Then Exit Sub         'no bar is loaded yet
    If UBound(CurHeadList) > 0 Then
        For i = 0 To UBound(CurHeadList)
            HeadEnableAr(i) = LotINFO.IsHeadEnabled(CurHeadList(i))            'disable heads based on previous testing or known wafer test devices
        Next i
    End If
End Sub

Private Sub MotionEvents_FillDesignatorsInitiate(SNs() As String)
    'here can validate and change the serial numbers
End Sub

Private Sub MotionEvents_FillDesignatorsTerminate(SNs() As String)
Dim i As Integer
Dim j As Integer
    For i = 0 To UBound(SNs, 1)
        SNs(i, 1) = LotINFO.IsHeadEnabled(SNs(i, 0))
    Next i
End Sub

Private Sub MotionEvents_LoadBarTerminate()
    If MotionEvents.CurrentState <> 0 Then
        'frmSimple.Status = "Testing Bar " & CStr(MotionEvents.CurBar) & " of tray " & CStr(MotionEvents.CurTray)
    End If
End Sub

Private Sub MotionEvents_ApplicationExit()
    Set MotionEvents = Nothing              'important event, do not delete
End Sub

Private Sub MotionEvents_UnloadBarTerminate()
    If TestingLastBar = True Then
        frmSimple.Status = "Lot testing complete, please scan a new lot id."
        frmSimple.EnableControls (True)
        frmSimple.txtLotNumber.Text = "SCAN NEW LOT ID!"
    End If
End Sub
