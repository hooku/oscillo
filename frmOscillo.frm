VERSION 5.00
Begin VB.Form frmOscillo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscillo"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOscillo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox dbg 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   60
      TabIndex        =   9
      Top             =   3060
      Width           =   2835
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "test"
      Height          =   375
      Left            =   3060
      TabIndex        =   10
      Top             =   60
      Width           =   1395
   End
   Begin Oscillo.vbalImageList vbalImgLst 
      Left            =   120
      Top             =   6960
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   1148
      Images          =   "frmOscillo.frx":000C
      Version         =   65536
      KeyCount        =   1
      Keys            =   ""
   End
   Begin VB.CheckBox chkX 
      BackColor       =   &H00FFF3EF&
      Caption         =   "????????"
      ForeColor       =   &H00C65D21&
      Height          =   255
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   2400
   End
   Begin Oscillo.xpSlider xpSliVPos 
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2220
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   767
   End
   Begin VB.PictureBox picInputSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF3EF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2700
      Begin VB.CheckBox chkInput 
         BackColor       =   &H00FFF3EF&
         Caption         =   "CH&1 - idle"
         Height          =   255
         Index           =   0
         Left            =   60
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "Toggle to Enable/Disable Channel"
         Top             =   60
         Width           =   2400
      End
      Begin VB.CheckBox chkInput 
         BackColor       =   &H00FFF3EF&
         Caption         =   "CH&4 - idle"
         Height          =   255
         Index           =   3
         Left            =   60
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1140
         Width           =   2400
      End
      Begin VB.CheckBox chkInput 
         BackColor       =   &H00FFF3EF&
         Caption         =   "CH&2 - idle"
         Height          =   255
         Index           =   1
         Left            =   60
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   420
         Width           =   2400
      End
      Begin VB.CheckBox chkInput 
         BackColor       =   &H00FFF3EF&
         Caption         =   "CH&3 - idle"
         Height          =   255
         Index           =   2
         Left            =   60
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   780
         Width           =   2400
      End
   End
   Begin VB.Timer tmrPaint 
      Interval        =   100
      Left            =   720
      Top             =   7080
   End
   Begin Oscillo.vbalExplorerBarCtl vbalExp 
      Height          =   6000
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   10583
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin VB.PictureBox picGL 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   3000
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "frmOscillo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

'Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const MF_STRING = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_SEPARATOR = &H800&

Dim WithEvents sck As clsSock
Attribute sck.VB_VarHelpID = -1
Dim gl As clsGL
Dim data_gap As Integer

Private Sub chkInput_Click(Index As Integer)
    If Me.chkInput(Index).Value = Unchecked Then
        sck.Close_Server Index + 1
    End If
End Sub

Private Sub cmdTest_Click()
    gl.Paint
    'MsgBox Me.xpSliVPos(1).Value
End Sub

Private Sub Form_DblClick()
'    frmToolbox.Show
End Sub

Private Sub Form_Load()
    log "app start"

    Set sck = New clsSock
    Set gl = New clsGL

    ' init winsock:
    If sck.Create_Server(LISTEN_ADDR, LISTEN_PORT) = False Then
        MsgBox "socket server error"
    End If

    ' init opengl:
    If gl.Create(Me.picGL.hwnd, Me.picGL.hdc) = False Then
        MsgBox "opengl error"
    End If

    '
    Me.chkInput(0).ForeColor = CH1_RGB
    Me.chkInput(1).ForeColor = CH2_RGB
    Me.chkInput(2).ForeColor = CH3_RGB
    Me.chkInput(3).ForeColor = CH4_RGB

    ' append about menu
    Dim h_sysmenu As Long
    h_sysmenu = GetSystemMenu(Me.hwnd, False)
    AppendMenu h_sysmenu, MF_SEPARATOR, 0, vbNullString
    AppendMenu h_sysmenu, MF_STRING Or MF_GRAYED, 0, "(C) 2013 pengxiaojing"
    AppendMenu h_sysmenu, MF_STRING Or MF_GRAYED, 0, "win2000.howeb.cn"

    Me.Caption = App.ProductName & " " & App.Major & "." & App.Minor & " Build " & App.Revision

    init_vbal_exp
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Me.vbalExp.Height = Me.ScaleHeight
    Me.picGL.Width = Me.ScaleWidth - Me.vbalExp.Width
    Me.picGL.Height = Me.ScaleHeight

    gl.Repaint
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sck.Destory_Server
    gl.Destory

    Set sck = Nothing
    Set gl = Nothing

    app_close
End Sub

Private Sub picGL_Paint()
    gl.Repaint
End Sub

Private Sub sck_SCKConnect(channel As Integer, iP As String)
    Me.chkInput(channel - 1).Caption = iP
    Me.chkInput(channel - 1).Value = Checked
End Sub

Private Sub sck_SCKData(channel As Integer, data() As Byte, bytesTotal As Long)
' convert byte into packet array
' but we need one more extra memcpy (200 byte in 100ms doesn't matter)

' === extract data ===
    Dim pkt() As Packet
    Dim pkt_count As Integer

    Debug.Print "got!!!" & bytesTotal

    data_gap = data_gap + 1
    'If data_gap Mod 5 <> 0 Then
    '    Exit Sub
    'End If

    'pkt_count = bytesTotal / 2
    pkt_count = MAX_DATA_LEN / 2    ' 2 Byte/pkt
    ReDim pkt(pkt_count)

    Dim i As Integer

    On Error Resume Next

    ' encapsulate into data packet
    For i = 0 To pkt_count - 1
        pkt(i).hi = data(i * 2)
        pkt(i).low = data(i * 2 + 1)
    Next i

    ' push data into gl buffer
    log "pushed into opengl buffer"
    gl.Push channel, pkt

    ' for debug purpose only
    'gl.Push 2, pkt
    'gl.Push 3, pkt
    'gl.Push 4, pkt

    ' === extract timestamp ===
    'Dim reference_time As SYSTEMTIME, current_time As SYSTEMTIME
    Dim reference_time As Long, current_time As Long
    'CopyMemory reference_time, data(MAX_DATA_LEN), MAX_TIMESTAMP_LEN
    CopyMemory reference_time, data(MAX_DATA_LEN), MAX_TIMESTAMP_LEN
    current_time = GetTickCount

    'GetSystemTime current_time

    ' TODO: calc time diff

    log "time = " & (current_time - reference_time) & " ms"
End Sub

Private Sub sck_SCKDisconnect(channel As Integer)
    Me.chkInput(channel - 1).Caption = "CH&" & channel & " - idle"
    Me.chkInput(channel - 1).Value = Unchecked
End Sub

Private Sub tmrPaint_Timer()
    gl.Paint
    'sck.Send_Pkt
End Sub

Private Sub init_vbal_exp()
    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Dim i As Integer

    With Me.vbalExp
        .Redraw = False
        .ImageList = Me.vbalImgLst

        .Bars.Clear

        Set cBar = .Bars.Add(, "INPUTSEL", "Input Selection")
        'cBar.CanExpand = False
        cBar.IsSpecial = True

        Set cItem = cBar.Items.Add(, "INPUTSEL_CONTROL", , , eItemControlPlaceHolder)
        cItem.Control = Me.picInputSel
        cItem.CanClick = False

        Set cBar = .Bars.Add(, "INPUTEMU", "Input Emulation")

        Set cItem = cBar.Items.Add(, "IE_CONN", "Connect")
        Set cItem = cBar.Items.Add(, "IE_SEND", "Send Fake Data")
        Set cItem = cBar.Items.Add(, "IE_DISC", "Disconnect")

        For i = 1 To CHANNEL_COUNT
            'Load Me.xpSliHPos(i)
            Load Me.xpSliVPos(i)
            Set cBar = .Bars.Add(, "POS_" & i, "Channel " & i)
            'Set cItem = cBar.Items.Add(, "HPOS" & i & "_CAPTION", "SEC/DIV:", , eItemText)
            'Set cItem = cBar.Items.Add(, "HPOS_CONTROL", , , eItemControlPlaceHolder)
            'cItem.Control = Me.xpSliHPos(i)
            Set cItem = cBar.Items.Add(, "VPOS" & i & "_CAPTION", "Position:", , eItemText)
            Set cItem = cBar.Items.Add(, "VPOS_CONTROL", , , eItemControlPlaceHolder)
            cItem.Control = Me.xpSliVPos(i)
            'cItem.CanClick = False
        Next i

        Set cBar = .Bars.Add(, "DBGLOG", "Debug Log")
        Set cItem = cBar.Items.Add(, "DBGLOG_CONTROL", , , eItemControlPlaceHolder)
        cItem.Control = Me.dbg
        cItem.CanClick = False

        ' === apply color style ===
        .UseExplorerStyle = False
        .UseExplorerTransitionStyle = True

        ' TODO: add color definition to const variable

        .BackColorStart = &HE6AA8C    'rgb(140, 170, 230)
        .BackColorEnd = &HDC8764    'rgb(100, 135, 220)

        For i = 1 To .Bars.Count
            With .Bars(i)
                If .IsSpecial = True Then
                    .TitleBackColorDark = &HB54900    'rgb(0, 73, 181)
                    .TitleBackColorLight = &HCE5D29    'rgb(41, 93, 206)
                    .TitleForeColor = &HFFFFFF
                    .TitleForeColorOver = &HE2B598    'rgb(152, 181, 226)
                    .BackColor = &HFFF3EF    'rgb(239, 243, 255)
                Else
                    .TitleBackColorDark = &HFFFFFF
                    .TitleBackColorLight = &HF7D3C6    'rgb(198, 211, 247)
                    .TitleForeColor = &H945025    'rgb(37, 80, 148)
                    .TitleForeColorOver = &HB9642E    'rgb(46, 100, 185)
                    .BackColor = &HFCE3D6    'rgb(214, 227, 252)
                End If
            End With
        Next i

        Me.picInputSel.BackColor = &HFFF3EF
        For i = 1 To CHANNEL_COUNT
            'Me.xpSliHPos(i).BackColor = &HFCE3D6
            Me.xpSliVPos(i).BackColor = &HFCE3D6
            Me.xpSliVPos(i).Max = MAX_SLI_V
            Me.xpSliVPos(i).Value = MAX_SLI_V * (i - 1) / CHANNEL_COUNT + MAX_SLI_V / CHANNEL_COUNT / 2
            xpSliVPos_Change (i)
        Next i

        ' === apply ===
        Dim ctl As Control

        For Each ctl In Me.Controls
            If TypeOf ctl Is CheckBox Then
                ctl.MouseIcon = LoadResPicture("HAND", vbResCursor)
            End If
        Next

        ' init bar status
        For i = 1 To .Bars.Count
            If (InStr(.Bars(i).Key, "POS") > 0) Then
                'If StrComp(.Bars(i).Key, "POS_1") <> 0 Then
                .Bars(i).State = eBarCollapsed
                'End If
            End If
        Next i

        .Redraw = True

    End With
End Sub

Private Sub txtServerIP_Click()
'SendKeys "{home}+{end}"
End Sub

Private Sub vbalExp_ItemClick(itm As cExplorerBarItem)
    Select Case itm.Key
    Case "IE_CONN"
        sck.Create_Client TEST_ADDR, TEST_PORT
    Case "IE_SEND"
        'MsgBox "A"
        sck.Send_Packet
    Case "IE_DISC"
        sck.Destory_Client
    End Select
End Sub

Private Sub xpSliVPos_Change(Index As Integer)
    gl.SetVPos Index, Me.xpSliVPos(Index).Value
End Sub
