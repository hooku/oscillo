VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' color
Private Const BG_RGB As Long = &HFFFFFF
Private Const SIN_RGB As Long = &HFF

Private Const AXIS_RGB As Long = &HDFDFDF
Private Const GRID_RGB As Long = &HEFEFEF

' param
Private Const AXIS_WIDTH As Integer = 2

Private Const X_AXIS_LEN As Integer = GL_WIDTH
Private Const Y_AXIS_LEN As Integer = GL_HEIGHT
Private Const X_AXIS_HALF As Integer = GL_WIDTH / 2
Private Const Y_AXIS_HALF As Integer = GL_HEIGHT / 2

Private g_hwnd As Long, g_hdc As Long
Private hRC As Long

Private scan_pos As Integer

' draw buffer is a 2 dimension array based ring buffer
Private draw_buff(CHANNEL_COUNT + 1, MAX_DATA_BUFF) As Integer
Private draw_buff_head(CHANNEL_COUNT + 1) As Integer
Private draw_buff_len(CHANNEL_COUNT + 1) As Integer
Private color_array(CHANNEL_COUNT + 1) As Long

Private channel_v_pos(CHANNEL_COUNT + 1) As Integer

Private Sub gl_color(rgb As Long)
    Dim r As Double, g As Double, B As Double
    ' find r,g,b value
    r = (rgb Mod &H100) / &H100
    g = ((rgb \ &H100) Mod &H100) / &H100
    B = (rgb \ &H10000) / &H100
    glColor3d r, g, B
End Sub

Private Sub gl_line(x0 As Double, y0 As Double, X1 As Double, y1 As Double, Optional Width As Integer = 1)
    glLineWidth Width
    glBegin bmLines
    glVertex2d x0, y0
    glVertex2d X1, y1
    glEnd
End Sub

Private Sub draw_grid()
' draw hline
    Dim grid_space As Double
    grid_space = Y_AXIS_LEN / DEFAULT_SEC_DIV

    gl_color GRID_RGB

    Dim grid_xy As Double
    For grid_xy = 0 To Y_AXIS_LEN Step grid_space
        gl_line 0, grid_xy, GL_WIDTH, grid_xy
    Next grid_xy

' draw vline
    grid_space = X_AXIS_LEN / DEFAULT_SEC_DIV
    For grid_xy = 0 To X_AXIS_LEN Step grid_space
        gl_line grid_xy, 0, grid_xy, GL_HEIGHT
    Next grid_xy

End Sub

Private Sub draw_axis()
' draw x axis
    gl_color AXIS_RGB
    
    gl_line 0, Y_AXIS_HALF, X_AXIS_LEN, Y_AXIS_HALF, AXIS_WIDTH
    'gl_line X_AXIS_HALF, 0, X_AXIS_HALF, Y_AXIS_LEN, AXIS_WIDTH
End Sub

Private Sub draw_reference()
' enable line attributes: dash, 1px
    glPushAttrib amEnableBit

    glLineStipple 1, &HF0F0
    glEnable glcLineStipple

    draw_grid
    draw_axis

    glPopAttrib
End Sub

Private Sub gl_sin(x As Double, y As Double)
    glBegin bmPoints
    glVertex2d x, y
    glEnd
End Sub

Private Sub draw_wave()
' draw wave for all channel
    Dim channel_scan_pos As Integer
    
    Dim i_channel As Integer
    
    For i_channel = 1 To CHANNEL_COUNT
        
        channel_scan_pos = scan_pos
        
        Dim i_byte As Integer
        Dim byte_to_eat As Integer
        ' exhaust a pkt len data
        ' TODO: we'd implement mutex here
        If draw_buff_len(i_channel) >= (MAX_DATA_LEN / 2) Then
            byte_to_eat = (MAX_DATA_LEN / 2)
        Else
            byte_to_eat = draw_buff_len(i_channel)
        End If
        
        ' select channel color
        gl_color color_array(i_channel)
        
        For i_byte = 0 To byte_to_eat
            gl_sin CDbl(channel_scan_pos), CDbl(draw_buff(i_channel, draw_buff_head(i_channel))) + channel_v_pos(i_channel)
            draw_buff_head(i_channel) = (draw_buff_head(i_channel) + 1) Mod MAX_DATA_BUFF
            channel_scan_pos = channel_scan_pos + 1.5
        Next i_byte
        draw_buff_len(i_channel) = draw_buff_len(i_channel) - byte_to_eat
        
        ' draw straight line
        For i_byte = byte_to_eat To (MAX_DATA_LEN / 2) - 1
            gl_sin CDbl(channel_scan_pos), CDbl(channel_v_pos(i_channel))
            channel_scan_pos = channel_scan_pos + 1.5
        Next i_byte
  
    Next i_channel
    
    ' FIX ME: do not add manual offset
    scan_pos = scan_pos + (MAX_DATA_LEN / 2) * 1.5 + 40
' check if we need reset the scan position
    If scan_pos >= X_AXIS_LEN Then
        log "clean canvas"
        
        scan_pos = 0
        Repaint
        Repaint
    End If
End Sub

Public Function Create(hwnd As Long, hdc As Long) As Boolean
    Create = True
    Dim ret As Long

    g_hwnd = hwnd
    g_hdc = hdc

    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long    ' pixel format to choose
    With pfd
        .nSize = Len(pfd)
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_SUPPORT_OPENGL Or PFD_TYPE_RGBA
        .iPixelType = PFD_TYPE_RGBA
        .cColorBits = 24
        .cDepthBits = 0
        .iLayerType = PFD_MAIN_PLANE
    End With

    pf = ChoosePixelFormat(g_hdc, pfd)
    ret = SetPixelFormat(g_hdc, pf, pfd)

    ' create gl context
    hRC = wglCreateContext(g_hdc)    ' rendering context
    ret = wglMakeCurrent(g_hdc, hRC)

    ' setup 2d projection
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0, X_AXIS_LEN, Y_AXIS_LEN, 0, 0, 1
    glMatrixMode mmModelView

    ' disable depth buffer
    glDisable glcDepthTest

    ' clear screen
    glClear clrColorBufferBit
    
    color_array(1) = CH1_RGB
    color_array(2) = CH2_RGB
    color_array(3) = CH3_RGB
    color_array(4) = CH4_RGB
End Function

Public Function Destory() As Boolean
    Destory = True

End Function

Public Sub SetVPos(channel As Integer, vpos As Integer)
    channel_v_pos(channel) = Y_AXIS_LEN * (vpos / MAX_SLI_V)
End Sub

Friend Sub Push(channel As Integer, pkt() As Packet)
' push data into draw buffer
' note that draw buffer isn't ring buffer and has a fixed length

' just use ref here to rid memcpy

    Dim pkt_len As Integer
    Dim data_value As Integer
    Dim draw_buff_tail As Integer
    
    ' FIXME: We'd calculate real packet length
    pkt_len = UBound(pkt)
    If draw_buff_len(channel) + pkt_len <= MAX_DATA_BUFF Then
        Dim i As Integer
        For i = 0 To pkt_len
            ' convert 2-byte packet into integer
            data_value = pkt(i).low * &H100 + pkt(i).hi
            'data_value = pkt(i).hi * &H100 + pkt(i).low
            draw_buff_tail = (draw_buff_head(channel) + draw_buff_len(channel) + i) Mod MAX_DATA_BUFF
            draw_buff(channel, draw_buff_tail) = data_value
        Next i

        draw_buff_len(channel) = draw_buff_len(channel) + pkt_len
        'draw_buff_head(channel) = (draw_buff_head(channel) + pkt_len) Mod MAX_DATA_LEN

        log "draw_buff" & channel & " = " & draw_buff_len(channel)
    Else
        log "draw buffer full, data is loss", LL_WARN
    End If
End Sub

Public Sub Repaint()
    glClearColor 1, 1, 1, 0
    glClear clrColorBufferBit

    draw_reference
    SwapBuffers g_hdc
End Sub

Public Sub Paint() ' to be fucked
    ' TODO: warning: we're not using double buffer here, it is stupid
    SwapBuffers g_hdc
    draw_wave
    SwapBuffers g_hdc
End Sub

'Public Sub paint() ' to be fucked
'Dim old_scan_pos As Integer
'' TODO: don't draw twice
'old_scan_pos = scan_pos
'    draw_wave (1)
'    SwapBuffers g_hdc
'
'scan_pos = old_scan_pos
'    draw_wave (1)
'    SwapBuffers g_hdc
'
'    draw_buff_len(1) = 0
'End Sub

Public Sub paint2()
    'glTranslatef 20, 0, 0
    glViewport 10, 20, 0, 0
'    glOrtho -62, 10, -12, 47, 451, 255
    SwapBuffers g_hdc
    'glTranslatef 20, 0, 0
    glViewport 10, 20, 0, 0
    SwapBuffers g_hdc
End Sub
