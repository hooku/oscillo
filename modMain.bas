Attribute VB_Name = "modMain"
Option Explicit

Public Const LISTEN_ADDR As String = "0.0.0.0"
Public Const LISTEN_PORT As Integer = 9000

Public Const MAX_DATA_BUFF As Integer = 2048   ' data item count

'
Public Const TEST_ADDR As String = "127.0.0.1"
Public Const TEST_PORT As String = 9000

Public Const MAX_DATA_VALUE As Integer = 4095
Public Const MAX_DATA_LEN As Integer = 200
Public Const MAX_PKT_HEAD_LEN As Integer = 4
Public Const MAX_TIMESTAMP_LEN As Integer = 4
Public Const MAX_PKT_LEN As Integer = MAX_PKT_HEAD_LEN + MAX_DATA_LEN + MAX_TIMESTAMP_LEN ' bytes
Public Const MAX_PKT_COUNT As Integer = MAX_PKT_LEN / 2

Public Const CHANNEL_COUNT As Integer = 4

Public Const USE_OPENGL As Boolean = False
Public Const PROMPT_PORT As Boolean = False

Public Const MAX_SLI_V = 100

Public Const CH1_RGB As Long = &HFF&
Public Const CH2_RGB As Long = &HFF00&
Public Const CH3_RGB As Long = &HFF0000
Public Const CH4_RGB As Long = &HFF00FF


' === PACKAGE Definition ===
'#define HOST_CONFIRM    0x01
'#define ID_CONFLICT     0x02
'#define DATA_REQUEST    0x03
'#define DATA_STOP       0x04
'#define CHECKSUM_ERROR  0x05
'
'#define WIFI_DATA_PACKAGE_SIZE  200
'#define WIFI_DATA_NUM   100
'
'#define ENDCHAR 0x0a
'#define STARTCHAR ':'
'typedef struct
'{
'   u8 Command;
'   u8 FrameEnd;
'} HostPackageDef;
'
'typedef struct
'{
'   u8 FrameHead;
'   u8 UserId;
'   u8 length;
'   u8 WifiSignal;
'   u8 checksum;
'} ClientPackageDef;

Public Const ESCAPE_CHAR As Byte = &H3A& ' ":"

Public Const HOST_CONFIRM As Byte = &H1&
Public Const DATA_REQUEST As Byte = &H3&

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type hostPkg
    command As Byte
    frameTail As Byte
End Type

Public Type clientPkg
    frameHead As Byte
    userID As Byte
    Length As Byte
    wifiSignal As Byte
    checkSum As Byte
End Type

' === UI Related ===

Public Const GL_HEIGHT As Integer = 4096 '640
Public Const GL_WIDTH As Integer = GL_HEIGHT * 2 '960

Public Const DEFAULT_SEC_DIV As Integer = 20

Public Enum logLevel
    LL_VERBOSE
    LL_WARN
    LL_ERROR
End Enum

Private Const MIN_LOG_LEVEL = LL_VERBOSE

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetProcessDEPPolicy Lib "kernel32" (ByVal dwFlags As Long) As Long

Public Type Packet
    hi As Byte
    low As Byte
End Type

Sub Main()
    
    InitCommonControls
    
    On Error Resume Next
    SetProcessDEPPolicy 0
    If Err Then
        'MsgBox "Please Turn DEP off or app would crash!", vbExclamation
    End If
    
    Randomize timer

    Load frmOscillo

    frmOscillo.Width = Screen.TwipsPerPixelX * (960 + 200)
    frmOscillo.Height = Screen.TwipsPerPixelY * (640 + 25)

    'center workspace
    'Dim ws_width As Integer, ws_height As Integer
    'ws_width = frmOscillo.width + frmToolbox.width
    'ws_height = frmOscillo.Height

    'frmOscillo.Left = (Screen.width - ws_width) / 2
    'frmOscillo.Top = (Screen.Height - ws_height) / 2

    'frmToolbox.Left = frmOscillo.Left + frmOscillo.width
    'frmToolbox.Top = frmOscillo.Top

    frmOscillo.Show
    'frmToolbox.Show
End Sub

Public Sub app_close()
    'Unload frmToolbox
End Sub

Public Sub app_about()
    MsgBox App.ProductName & " " & App.Major & "." & App.Minor & " Build " & App.Revision, , "About " & App.ProductName
End Sub

Public Sub log(Text As String, Optional level As logLevel = LL_VERBOSE)
    'If left(Text, 4) <> "time" Then Exit Sub
    If level >= MIN_LOG_LEVEL Then
        If frmOscillo.dbg.ListCount > 50 Then
            frmOscillo.dbg.Clear
        End If
        frmOscillo.dbg.AddItem Text, 0
    End If
End Sub
