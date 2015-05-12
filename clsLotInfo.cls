VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLotInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TestSetup As String
Public LotSize As String
Public FirstBar As String
Public BarType As Integer
Public lotID As String

Private BarMap() As Boolean
Private HeadMap() As Boolean
Private HeadsToSkip() As String

Function LoadLotInfo(LotNumber As String)
'TODO: Look up the lot information on the network.
lotID = LotNumber
If LotNumber = "LOT#1" Then
    TestSetup = "c:\Setups\1.mds"
    Call LoadBarMap(LotNumber)
    LotSize = 4                'there are ten bars in the lot, including the ones that are not present in the tray
    FirstBar = 3                'first bar is 0
'    BarType = HeadType.DOWN     'only down heads are in this lot
    ReDim HeadsToSkip(0 To 5)
    HeadsToSkip(0) = "5C6DA0F501"
    HeadsToSkip(1) = "5C6DA0F502"
    HeadsToSkip(2) = "5C6DA0F503"
    HeadsToSkip(3) = "5C6DA0F504"
    HeadsToSkip(4) = "5C6DA0F505"
    HeadsToSkip(5) = "5C6DA0F510"
ElseIf LotNumber = "LOT#2" Then
    TestSetup = "c:\Setups\2.mds"
    LotSize = 10                 'there are ten bars in the lot, including the ones that are not present in the tray
    FirstBar = 3                'first bar is 0
'    BarType = HeadType.UP       'only down heads are in this lot
    Call LoadBarMap(LotNumber)
    ReDim HeadsToSkip(0 To 5)
    HeadsToSkip(0) = "5C6DA0B720"
    HeadsToSkip(1) = "5C6DA0B721"
    HeadsToSkip(2) = "5C6DA0B722"
    HeadsToSkip(3) = "5C6DA0B723"
    HeadsToSkip(4) = "5C6DA0B724"
    HeadsToSkip(5) = "5C6DA0B725"
End If
End Function

Sub LoadBarMap(lotIDStr As String)
Dim i As Integer
Dim j As Integer
    If lotIDStr = "LOT#1" Then
        ReDim BarMap(0 To 0, 0 To 22)
        BarMap(0, 0) = True           'bar is present in tray 0, slot 0
        BarMap(0, 1) = True
        BarMap(0, 2) = True
        BarMap(0, 3) = True
        BarMap(0, 4) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 5) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 6) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 7) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 8) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 9) = False          'bar is not present in Tray 0, slot 4
        BarMap(0, 10) = False          'bar is not present in Tray 0, slot 4
    ElseIf lotIDStr = "LOT#1" Then
        ReDim BarMap(0 To 0, 0 To 22)
        BarMap(0, 0) = False 'bar is present in tray 0, slot 0
        BarMap(0, 1) = False
        BarMap(0, 2) = False
        BarMap(0, 3) = True
        BarMap(0, 4) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 5) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 6) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 7) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 8) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 9) = True          'bar is not present in Tray 0, slot 4
        BarMap(0, 10) = True          'bar is not present in Tray 0, slot 4
    End If
End Sub

Public Function IsBarPresent(traynum As Integer, barnum As Integer) As Boolean
    IsBarPresent = BarMap(traynum, barnum)
End Function

Function IsHeadEnabled(sn As String) As Boolean
Dim i As Integer
IsHeadEnabled = True
    For i = 0 To UBound(HeadsToSkip)
        If HeadsToSkip(i) = sn Then
            IsHeadEnabled = False
            Exit For
        End If
    Next i
End Function
