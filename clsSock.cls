VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Enum serverState
    SS_WAITSHAKE
    SS_INSHAKE
    SS_CONNECTED
End Enum

Dim WithEvents sckServer As CSocketPlus
Attribute sckServer.VB_VarHelpID = -1
Dim WithEvents sckClient As CSocketPlus
Attribute sckClient.VB_VarHelpID = -1

Public Event SCKConnect(channel As Integer, iP As String)
Public Event SCKDisconnect(channel As Integer)
Public Event SCKData(channel As Integer, data() As Byte, bytesTotal As Long)

Dim connection_count As Integer
Dim test_phase As Integer

Dim server_state(CHANNEL_COUNT + 1) As serverState

Public Function Create_Server(iP As String, port As Integer) As Boolean
    Create_Server = True
    Set sckServer = New CSocketPlus
        
    Dim i_channel As Integer
    For i_channel = 0 To CHANNEL_COUNT
        sckServer.ArrayAdd (i_channel)
    Next i_channel

    If PROMPT_PORT = True Then
        port = InputBox("port", , "9000")
    End If

    sckServer.Bind 0, port, iP
    sckServer.Listen 0
    DoEvents
    
    log "listen on " & iP & ":" & port
End Function

Public Function Destory_Server() As Boolean
    Destory_Server = True
    sckServer.CloseSck 1
    Set sckServer = Nothing
End Function

Public Function Create_Client(iP As String, port As Integer) As Boolean
    If sckClient Is Nothing Then
        Create_Client = True
        Set sckClient = New CSocketPlus

        sckClient.ArrayAdd 0

        If sckClient.State(0) = sckClosed Then
            sckClient.Connect 0, sckClient.LocalIP(0), port
        End If
        DoEvents
        log "connect to " & iP & ":" & port
    End If
End Function

Public Function Destory_Client() As Boolean
    If Not sckClient Is Nothing Then
        Destory_Client = True
        sckClient.CloseSck 0
        Set sckClient = Nothing
    End If
End Function

Public Sub Send_Packet()
    If Not sckClient Is Nothing Then
        log "send pkt"

        'Dim payload(MAX_PKT_LEN - 1) As Byte
        Dim payload(MAX_PKT_LEN + 6) As Byte
        Dim pkt As Packet

' === fill packet header ===
    payload(0) = &HFF
    payload(1) = &HFF
    payload(2) = &HFF
    payload(3) = &HFF

' === fill data payload ===
        ' generate random data
        Dim sample_data As Integer
        Dim i As Integer

        For i = MAX_PKT_HEAD_LEN / 2 To MAX_PKT_HEAD_LEN / 2 + MAX_DATA_LEN / 2 - 1
            'sample_data = Int(Rnd * MAX_DATA_VALUE)
            sample_data = Sin(test_phase / 360) * 500 + 500
            'sample_data = 2000
            test_phase = test_phase + 1
            pkt.hi = sample_data \ &H100&
            pkt.low = sample_data Mod &H100&
            payload(i * 2) = pkt.low
            payload(i * 2 + 1) = pkt.hi
        Next i
        
' === fill timestamp payload ===
        ' TODO: embed current time
        
        ' memcpy
        ' ...

        sckClient.SendData 0, payload
    End If
End Sub

Public Sub Close_Server(channel As Integer)
    sckServer_CloseSck (channel)
End Sub

Private Sub sckClient_Connect(ByVal Index As Variant)
    log ">> fake handshake to server"
    ' send handshake to server
    Dim fake_handshake As clientPkg
    Dim payload() As Byte

    ReDim payload(LenB(fake_handshake))

    With fake_handshake
        .frameHead = ESCAPE_CHAR
    End With
    
    CopyMemory payload(0), fake_handshake, LenB(fake_handshake)
    'payload = StrConv(fake_handshake, vbFromUnicode, &H409)
    
    sckClient.SendData 0, payload
End Sub

Private Sub sckServer_CloseSck(ByVal Index As Variant)
    log "close"
    
    server_state(Index) = SS_WAITSHAKE
    connection_count = connection_count - 1
    
    sckServer.CloseSck Index
    RaiseEvent SCKDisconnect(CInt(Index))
End Sub

Private Sub sckServer_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
    log "conn req, id=" & requestID
    'sckServer.CloseSck
    'sckServer1.Accept requestID
    
    ' TODO: remove connection count
    
    If connection_count < CHANNEL_COUNT Then
        connection_count = connection_count + 1
    
        Dim i_server As Integer
        For i_server = 1 To CHANNEL_COUNT
        ' try to find the next empty tcp socket handler
            'log "aaaaa" & sckServer.State(i_server)
            If sckServer.State(i_server) = sckClosed Then
            'If sckServer.ArrayIndexInUse = False Then
                sckServer.Accept i_server, requestID
                Exit For
            End If
        Next i_server
        
        ' TODO: handle unaccepted connections
        
        RaiseEvent SCKConnect(i_server, sckServer.RemoteHostIP(Index))
    End If
End Sub

Private Sub sckServer_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
    Dim vta_data As Variant, byte_data() As Byte
    
    log "data arrive, len=" & bytesTotal

    Select Case server_state(Index)
    Case SS_WAITSHAKE
        ' once in handshake, we'd retrieve all bytes in one pull
        sckServer.GetData Index, vta_data, vbArray + vbByte, bytesTotal
        byte_data = vta_data
        
        Dim i_escape_char As Integer
        
        For i_escape_char = 0 To UBound(byte_data)
            If byte_data(i_escape_char) = ESCAPE_CHAR Then
                
                log "channel " & Index & " initialize handshake"
                server_state(Index) = SS_INSHAKE
                ' send host confirm & data request
                Dim payload() As Byte
                Dim host_package As hostPkg
                
                ReDim payload(LenB(host_package) - 1)
                
                host_package.frameTail = ESCAPE_CHAR
                
                host_package.command = HOST_CONFIRM
                CopyMemory payload(0), host_package, LenB(host_package)
                sckServer.SendData Index, payload
                
                ' TODO: insert some sleep
                Sleep 200
                
                host_package.command = DATA_REQUEST
                CopyMemory payload(0), host_package, LenB(host_package)
                sckServer.SendData Index, payload
                
                ' handshake successful
                
                server_state(Index) = SS_CONNECTED
                log "channel " & Index & " handshake succeed"

                Exit For
            End If
        Next i_escape_char
        
        'Exit Sub
    Case SS_CONNECTED

        ' socket buffer may contain bytes larger than MAX_PKT_LEN,
        ' so we'd to obtain them in several rounds
        Dim bytes_obtained As Long
        
        If bytesTotal < MAX_PKT_LEN Then
            Exit Sub
        End If
        
        For bytes_obtained = 0 To (bytesTotal - MAX_PKT_LEN) Step MAX_PKT_LEN
            
            
            
            ' check packet header
            sckServer.PeekData Index, vta_data, vbArray + vbByte, bytesTotal
            byte_data = vta_data
            
            Dim i_ffff As Integer
            Dim has_junk_data As Boolean
            
            has_junk_data = False
            For i_ffff = 0 To bytesTotal - 4
                If byte_data(i_ffff) = &HFF And _
                    byte_data(i_ffff + 1) = &HFF And _
                    byte_data(i_ffff + 2) = &HFF And _
                    byte_data(i_ffff + 3) = &HFF _
                Then
                    Exit For
                Else
                    has_junk_data = True
                End If
            Next i_ffff
            
            If has_junk_data = True Then
                ' drop them
                sckServer.GetData Index, vta_data, vbArray + vbByte, i_ffff
                'Exit Sub
                bytesTotal = bytesTotal - i_ffff
            End If
                        
            
            
            sckServer.GetData Index, vta_data, vbArray + vbByte, MAX_PKT_HEAD_LEN
            byte_data = vta_data
            
            sckServer.GetData Index, vta_data, vbArray + vbByte, MAX_PKT_LEN - MAX_PKT_HEAD_LEN
            byte_data = vta_data
            ' TODO: should?
            'DoEvents
            RaiseEvent SCKData(CInt(Index), byte_data, MAX_PKT_LEN)
        Next bytes_obtained
    End Select
End Sub
