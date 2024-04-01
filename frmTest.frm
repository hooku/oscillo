VERSION 5.00
Begin VB.Form frmToolbox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "oscilloscope console"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI Symbol"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLog 
      Height          =   1845
      Left            =   60
      TabIndex        =   12
      Top             =   7740
      Width           =   2775
   End
   Begin VB.Frame framePosition 
      Caption         =   "Position"
      Height          =   1755
      Left            =   60
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
      Begin VB.ComboBox cmbHPos 
         Height          =   375
         ItemData        =   "frmTest.frx":000C
         Left            =   120
         List            =   "frmTest.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
      End
      Begin VB.HScrollBar hscVPos 
         Height          =   255
         LargeChange     =   128
         Left            =   120
         Max             =   1023
         SmallChange     =   64
         TabIndex        =   8
         Top             =   540
         Value           =   511
         Width           =   2535
      End
      Begin VB.Label labHPos 
         Caption         =   "SEC/DIV"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   2475
      End
      Begin VB.Label labVPos 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame frameInputSel 
      Caption         =   "Input Selection"
      Height          =   1635
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox chkInput 
         Caption         =   "CH&3 - idle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   2415
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "CH&2 - idle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "CH&4 - idle"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "CH&1 - idle"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   2415
      End
   End
   Begin VB.Frame frameInputEmulation 
      Caption         =   "Input Emulation"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   6900
      Width           =   2775
      Begin VB.Timer tmrData 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2280
         Top             =   240
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_LOG_LEN = 64

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const WS_EX_TOPMOST = &H8

Dim WithEvents sckClient As CSocketMaster
Attribute sckClient.VB_VarHelpID = -1

Dim test_phase As Integer

Private Sub cmdStart_Click()
    If sckClient.State = sckClosed Then
        sckClient.Connect TEST_ADDR, TEST_PORT
    End If

    Me.tmrData.Enabled = True
End Sub

Private Sub Form_Load()
    Set sckClient = New CSocketMaster

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    Dim win_long As Long
    win_long = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    win_long = win_long Or WS_EX_NOACTIVATE Or WS_EX_TOPMOST
    SetWindowLong Me.hwnd, GWL_EXSTYLE, win_long
    log "app start"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sckClient.CloseSck
    Set sckClient = Nothing
End Sub

Public Sub log(Text As String)
    If Me.lstLog.ListCount > MAX_LOG_LEN Then
        Dim i As Integer
        For i = 0 To MAX_LOG_LEN / 4
            Me.lstLog.RemoveItem Me.lstLog.ListCount - 1
        Next i
    End If
    'Me.lstLog.AddItem Format(Timer, "0.00") & " " & text, 0
End Sub

' === send test packet ===

Private Sub tmrData_Timer()
    If sckClient.State = sckConnected Then
        'Me.tmrData.Enabled = False
        Send_Pkt
        'sckClient.CloseSck
    End If
End Sub
