VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const errModID As String = "clsHOption: "  'identifier for errorhandlers
Private Notifier As New Collection                  'all the objects in this collection should be notified about user feedback and others - whatever this application deems necessary
Private mvarCurrentApp$                            'identifier or instance in Quasi97

Private Enum eEventInt
   eCheckHealth = 1
   eRecover = 2
   eStartStop = 4
   eConnectHead = 8
   eSetupOpenClose = &H10
End Enum
 
Public Sub Initialize3(currentApp As String, ByRef qs As Object)      'pointer to valid instance Quasi97.application
   Dim PartID%
   Dim Grade$
   
   On Error GoTo errorhandler
   
   mvarCurrentApp = currentApp
   Unload objFrmSimple
   
   On Error Resume Next
   'Set the remote flag in the Integral Solutions Int'l/common/Quasi.ini file
   'This is only necessary one time on each computer
   'This example is based on this app being in a subdirectory inside the Integral Solutions Int'l directory
   Call WritePrivateProfileString("Quasi", "RemoteMode", "1", App.Path & "\..\common\Quasi.ini")
   'in remote mode the quasi software will only allow Operator mode and will suppress
   'the open setup, save to database and Operator Login windows
   'The Quasi.ini file can be modified using a text editor
   '"1" = Remote mode, "0" = regular mode
   'Remote=n is also available as a command line item for the quasi software, n = 1 or 0
   objFrmSimple.Show
   'SetWindowPos frmSimple.hwnd, WindowPos.vbTopMost, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
   'create an instance of the application object to start the quasi software
   Set qst = qs
   'set the event handler
   Call UpdateUserFeedback("Initializing QuasiEvents", 0)
   Set QuasiEvents = New clsQuasiEvents
   Call UpdateUserFeedback("Connecting to Barcont", 0)
   Dim hopt As Object 'quasi97.clsHardwareOption
   Set hopt = qst.HOptionManager.GetPointerByFunction("MechDriver").GetHandle
   If Not hopt Is Nothing Then
      Set BarDriver = qst.HOptionManager.GetPointerByFunction("MechDriver").GetHandle
   Else
      Call UpdateUserFeedback("Failed to find MechDriver", 0)
   End If
'   Set BarEvents = New clsBarEvents
   
   Call objFrmSimple.EnableControls
    'function will be called to activate the module
      
errorhandler:
   Select Case errorhandler(errModID & "Initialize3")
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
End Sub

Public Sub Terminate()
   Set Notifier = New Collection
   Unload objFrmSimple
   
   Set BarDriver = Nothing
   Set qst = Nothing
End Sub

Public Sub AddNotifier(notobj As Object)
   If Not notobj Is Nothing Then
      Dim s As Object
      For Each s In Notifier
         If s Is notobj Then Exit Sub
      Next
      Notifier.Add notobj
   End If
End Sub

Public Sub RemoveNotifier(notobj As Object)
   If notobj Is Nothing Then Exit Sub
   Dim i%
   For i = 1 To Notifier.Count
      If Notifier(i) Is notobj Then
         Notifier.Remove i
         Exit Sub
      End If
   Next i
End Sub

Public Function SetChannels(ByRef Channels%())
''   mvarChannels = Channels       'copying the array here
''   Dim i%, j%
''
''   For i = 0 To UBound(channelLUT)
''      channelLUT(i) = False
''   Next i
''
''   For i = 0 To UBound(mvarChannels)
''      channelLUT(mvarChannels(i)) = True
''   Next i
''   curHDW.AssignedChannel = mvarChannels(0) Mod 2
End Function

Public Sub GetNewProperties(ByRef colobjects As Collection, propDetails() As String)
'propdetails has 1+colobjects.count number of rows and 4 columns.
   'Call AddCustomProperty("AC Gain (dB)", "ACGaindB", False, Round(MinACGaindB, 1) & ";" & Round(MaxACGaindB, 1), False, colobjects, propDetails, Me)
End Sub

'interface functions for the test
Public Function GetUserControlClass() As String
   GetUserControlClass = ""
End Function

Public Sub ShowDiagnostics()
   objFrmSimple.Show
End Sub

Sub ShowUserMenu()
   objFrmSimple.Show
End Sub

Sub UpdateUserFeedback(lastUserFeeback As String, prg!)
   On Error GoTo errorhandler
   Dim s As Object
   For Each s In Notifier
      s.Message = lastUserFeeback
      s.Progress = prg
   Next
errorhandler:
   Select Case errorhandler(errModID & "UpdateUserFeedback")
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
End Sub
  
Public Sub DetectAllNew(ByRef DevAr() As String)
'the items in the collection must be easily identifiable for "Initialize" function later on
       
   'read serial number from the board and add it to the collection
   ReDim DevAr(1)
   DevAr(1) = ""

End Sub

Public Function ID$()
   ID = "RemoteOp.clsHOption"
End Function

Public Property Get EventInterests() As Long
   EventInterests = &H14 'eEventInt.eCheckHealth Or eEventInt.eRecover Or eEventInt.eConnectHead Or eEventInt.eSetupOpenClose Or eEventInt.eStartStop
End Property

Function CheckHealth(ByRef usrDescr As String, PartLoadedState As Byte) As Integer
'the module will use this call as an opportunity to check itself
'if fails this would return non-zero value and change the status. partloadedstate non-zero indicates that the head is probing (perhaps some checks can be done only while not probing the head)
Dim i%, N%
   
   On Error Resume Next
''
''   If curHDW.VerifyPower(True) = 0 Then
''      CheckHealth = 1001
''      usrDescr = "Power loss has been detected"
''      Call gUserMsgBox(usrDescr, vbCritical Or vbOKOnly)
''
''      Dim s As Object
''      For Each s In Notifier
''         Call s.DoEvent(Me, "SuspendFlag", True)
''      Next
''      HDWSuspended = 2              'more critical error, need to reinitialize
''   End If
End Function

Function Recover(ByRef usrDescr As String) As Integer
'module will use this call as an opportunity to recover from checkhealth suspension. For example if interlock door was opened
'0 means success
''   If HDWSuspended = 1 Then
''      HDWSuspended = 0 'occurs for the first time after everything is aborted and stop button is pressed
''   ElseIf HDWSuspended = 2 Then
''      Unload objFrmPulserDiag
''      Unload objFrmUser
''      Call curHDW.Initialize(mvarCurrentApp)
''      Dim s As Object
''      For Each s In Notifier
''         Call s.DoEvent(Me, "SuspendFlag", False)
''      Next
''      HDWSuspended = 0
''   End If
   'no critical errors are yet detecetd
End Function

Public Function Status() As Integer    'Initialized, active, suspended
'retrieves the current status from the device
   Status = 0
End Function

Public Sub NotifyOptionsUpdated()
'signals to the module that other modules have changed (configuration)
   
End Sub

Public Function StartStop(ByRef DoStart As Boolean)
'user presses start, the module may need to prepare (turn on relays, initialize something etc) - all of the slow processes
      If DoStart Then
'todo: do something when operator starts test
   Else
'todo: do something when the test is finished
   End If
End Function

Public Function ConnectHead(ByRef doconnect As Byte)
'the head gets connected and disconnected - whatever module needs to do to avoid zapping the head

End Function

Public Function SetupOpenClose(ByVal doopen As Boolean)
'signals to the module that setup open or setup close is in progress
   If doopen Then
        Set QuasiEvents.TestSeqEvnts = qst.TestSequencer
   Else
        Set QuasiEvents.TestSeqEvnts = Nothing
   End If
End Function

Public Property Get Present() As Boolean
   Present = True
End Property
