VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimple 
   Caption         =   "Custom User Interface"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "frmSimple.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstStatus 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   390
      TabIndex        =   5
      Top             =   1110
      Width           =   6405
   End
   Begin MSComctlLib.ProgressBar pbMainProgress 
      Height          =   585
      Left            =   390
      TabIndex        =   4
      Top             =   4200
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   1032
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdStop 
      Height          =   2175
      Left            =   7260
      Picture         =   "frmSimple.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      Height          =   2175
      Left            =   7260
      Picture         =   "frmSimple.frx":DFC8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   2175
   End
   Begin VB.TextBox txtLotNumber 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2820
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "AAA123"
      Top             =   330
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Lot Number"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   0
      Top             =   420
      Width           =   2055
   End
End
Attribute VB_Name = "frmSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    Status = "Loading Lot Information for LOT " & txtLotNumber.Text
    cmdStart.Enabled = False
    Call LotINFO.LoadLotInfo(txtLotNumber.Text)
    Call LoadTestSetup(LotINFO.TestSetup)
    'BarDriver.AutomationEvents.TotalBars = LotINFO.LotSize
    'BarDriver.AutomationEvents.MoveToNthBar (LotINFO.FirstBar)
    'BarDriver.AutomationEvents.BarType = LotINFO.BarType
    'qst.SystemParameters.lotID = txtLotNumber.Text
    ''Open a new log file
    qst.OptionsParameters.LogFileType = LOGCSV   'log file type definitions in globals
    qst.OptionsParameters.LogFileName = "c:\setups\" & LotINFO.TestSetup & CDate(Now) & LotINFO.lotID & " - NEW.csv"
    Call qst.QuasiParameters.OpenLogFile(qst.OptionsParameters.LogFileName)
   
    Status = "Starting the Test for LOT " & txtLotNumber.Text
    qst.QuasiParameters.Start = True
End Sub

Private Sub cmdStop_Click()
    cmdStart.Enabled = True
    qst.QuasiParameters.AbortTest = True
End Sub

Private Sub Form_Load()
    Call EnableControls(False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Set BarDriver = Nothing
    Set QuasiEvents = Nothing
    Call qst.CloseApplication
    Set qst = Nothing
    Exit Sub
End Sub

Public Property Let Status(NewStatus As String)
    lstStatus.AddItem NewStatus, 0
End Property

Private Sub txtLotNumber_Click()
    txtLotNumber.SelStart = 0
    txtLotNumber.SelLength = Len(txtLotNumber.Text)
End Sub

Private Sub txtLotNumber_GotFocus()
    txtLotNumber.SelStart = 0
    txtLotNumber.SelLength = Len(txtLotNumber.Text)
End Sub

Public Sub EnableControls(Optional DoEnable As Boolean = True)
    txtLotNumber.Enabled = DoEnable
    cmdStart.Enabled = DoEnable
    cmdStop.Enabled = DoEnable
End Sub
