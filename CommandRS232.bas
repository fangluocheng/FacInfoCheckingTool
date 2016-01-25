Attribute VB_Name = "Module2"
Option Explicit

Public Sub ENTER_FAC_MODE()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE E1 A0 00 01 04
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE1
    SendDataBuf(6) = &HA0
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H1
    SendDataBuf(9) = &H4
    
    Log_Info "Enter Factory Mode"
    cmdIdentifyNum = 0
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        isCmdDataRecv = True
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub EXIT_FAC_MODE()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE E1 A0 00 00 05
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE1
    SendDataBuf(6) = &HA0
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H5
    
    Log_Info "Exit Factory Mode"
    cmdIdentifyNum = 1

    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        isCmdDataRecv = True
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_SYS_VERSION()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE E4 13 00 00 B1
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE4
    SendDataBuf(6) = &H13
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &HB1
    
    Log_Info "Read system version"
    cmdIdentifyNum = 2
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_FLASH_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 0F 00 00 3C
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &HF
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H3C
    
    Log_Info "Read Flash information"
    cmdIdentifyNum = 3
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_HARDWARE_VERSION()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 16 00 00 25
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H16
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H25
    
    Log_Info "Read hardware version"
    cmdIdentifyNum = 4
    isCmdDataRecv = False

    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_DIMENSION_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 19 00 00 2A
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H19
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H2A
    
    Log_Info "Support 3D or not"
    cmdIdentifyNum = 5
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_24G_VERSION()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 14 00 00 27
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H14
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H27
    
    Log_Info "Read 2.4G Version"
    cmdIdentifyNum = 6
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_PANEL_NAME()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 17 00 00 24
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H17
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H24
    
    Log_Info "Read panel name"
    cmdIdentifyNum = 7
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_CARRIER_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 18 00 00 2B
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H18
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H2B
    
    Log_Info "Read carrier information"
    cmdIdentifyNum = 8
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_HDCP_KEY()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE 77 05 00 00 34
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H5
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H34
    
    Log_Info "Read HDCP key"
    cmdIdentifyNum = 9
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_MODEL_NAME()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 15 00 00 26
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H15
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H26
    
    Log_Info "Read model name"
    cmdIdentifyNum = 10
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_RESOLUTION_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 20 00 00 13
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H20
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H13
    
    Log_Info "Support 4K or 2K"
    cmdIdentifyNum = 11
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_MAC_ADDRESS()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE F0 01 01 00 B6
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HF0
    SendDataBuf(6) = &H1
    SendDataBuf(7) = &H1
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &HB6
    
    Log_Info "Read MAC address"
    cmdIdentifyNum = 12
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_CHANNEL_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE 77 32 00 00 03
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H32
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H3
    
    Log_Info "Read channel information"
    cmdIdentifyNum = 13
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_PARTITION_VER()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 03 FE 77 13 00 00 20
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H13
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H20
    
    Log_Info "Read partition version(DDR)"
    cmdIdentifyNum = 14
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_AREA_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE 77 33 00 00 02
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H33
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H2
    
    Log_Info "Read area information"
    cmdIdentifyNum = 15
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

Public Sub READ_DEVICE_KEY()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE 77 34 00 00 05
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H34
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H5
    
    Log_Info "Read device key"
    cmdIdentifyNum = 16
    isCmdDataRecv = False
    
    If False Then
        Form1.MSComm1.Output = SendDataBuf
    Else
        Form1.tcpClient.SendData SendDataBuf
    End If
End Sub

