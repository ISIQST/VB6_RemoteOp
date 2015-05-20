Attribute VB_Name = "modMain"
Option Explicit

Global qst As quasi97.Quasi97_Application
Global BarDriver As Object 'BarContLib.Driver
Global QuasiEvents As clsQuasiEvents
Global LotINFO As New clsLotInfo
'Global BarEvents As clsBarEvents
Global Const LOGACCESS = 1
Global Const LOGEXCEL = 2
Global Const LOGCSV = 3
Global Const LOGCSVSINGLE = 4
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Enum WindowPos
    vbTopMost = -1&
    vbNotTopMost = -2&
End Enum

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOACTIVATE = &H10
Public EndTest As Boolean
Public objFrmSimple As New frmSimple
Global gNetHostCallBack As Object

Sub gShowForm(ByRef frPtr As Object)
   If gNetHostCallBack Is Nothing Then
      frPtr.Show
   Else
      Call gNetHostCallBack.ShowForm(frPtr)
   End If
End Sub

Sub LoadTestSetup(TestProgramFile As String)

On Error GoTo errorhandler
   'open a setup file
   If qst.QuasiParameters.SetupFileName <> TestProgramFile Then
        objFrmSimple.Status = "Loading Setup Database " & TestProgramFile
        qst.QuasiParameters.OpenSetupFile (TestProgramFile)
        DoEvents
   End If
   'Set the Operator ID etc (not necessary but usually desired)
   qst.SystemParameters.OperatorID = "Remote"

errorhandler:
Select Case errorhandler("Run Test Example Function")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
End Select
End Sub

Function errorhandler%(FuncName$)
    If Err <> 0 Then
        errorhandler = MsgBox(CStr(Err) + " [" & FuncName & "] : " + Error, vbAbortRetryIgnore, Err.Source)
    End If
End Function


