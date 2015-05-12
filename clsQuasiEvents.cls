VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuasiEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents TestSeqEvnts As TestSequencer
Attribute TestSeqEvnts.VB_VarHelpID = -1

Private Sub QuasiEvnts_CloseLogFileInitiate(FileName As String, abort As Boolean)
   Set TestSeqEvnts = Nothing
End Sub

Private Sub TestSeqEvnts_RunTestInitiate()
   EndTest = False
End Sub

Private Sub TestSeqEvnts_RunTestTerminate()
   EndTest = True
End Sub
