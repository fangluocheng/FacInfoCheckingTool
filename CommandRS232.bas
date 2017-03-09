Attribute VB_Name = "Module2"
Option Explicit

Private mSendDataBuf(0 To 9) As Byte

Private Sub SendCmd()
    If gblUartMode Then
        FormMain.MSComm1.Output = mSendDataBuf
    Else
        FormMain.tcpClient.SendData mSendDataBuf
    End If
End Sub

Public Sub ENTER_FAC_MODE()
    '6E 51 86 03 FE E1 A0 00 01 04
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &HE1
    mSendDataBuf(6) = &HA0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H1
    mSendDataBuf(9) = &H4
    
    Log_Info "Enter Factory Mode"
    cmdIdentifyNum = 0
    
    SendCmd
End Sub

Public Sub EXIT_FAC_MODE()
    '6E 51 86 03 FE E1 A0 00 00 05
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &HE1
    mSendDataBuf(6) = &HA0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H5
    
    Log_Info "Exit Factory Mode"
    cmdIdentifyNum = 1

    SendCmd
End Sub

Public Sub READ_MODEL_NAME()
    '6E 51 86 03 FE 77 15 00 00 26
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H15
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H26
    
    Log_Info "Read model name"
    cmdIdentifyNum = 2
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_SYS_VERSION()
    '6E 51 86 01 FE E4 13 00 00 B1
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &HE4
    mSendDataBuf(6) = &H13
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &HB1
    
    Log_Info "Read system version"
    cmdIdentifyNum = 3
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_FLASH_INFO()
    '6E 51 86 03 FE 77 0F 00 00 3C
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &HF
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H3C
    
    Log_Info "Read Flash information"
    cmdIdentifyNum = 4
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_HARDWARE_VERSION()
    '6E 51 86 03 FE 77 16 00 00 25
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H16
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H25
    
    Log_Info "Read hardware version"
    cmdIdentifyNum = 5
    gblCmdDataRecv = False

    SendCmd
End Sub

Public Sub READ_DIMENSION_INFO()
    '6E 51 86 03 FE 77 19 00 00 2A
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H19
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H2A
    
    Log_Info "Support 3D or not"
    cmdIdentifyNum = 6
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_CHANNEL_INFO()
    '6E 51 86 01 FE 77 32 00 00 03
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H32
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H3
    
    Log_Info "Read channel information"
    cmdIdentifyNum = 7
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_24G_VERSION()
    '6E 51 86 03 FE 77 14 00 00 27
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H14
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H27
    
    Log_Info "Read 2.4G Version"
    cmdIdentifyNum = 8
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_PANEL_NAME()
    '6E 51 86 03 FE 77 17 00 00 24
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H17
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H24
    
    Log_Info "Read panel name"
    cmdIdentifyNum = 9
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_CARRIER_INFO()
    '6E 51 86 03 FE 77 18 00 00 2B
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H18
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H2B
    
    Log_Info "Read carrier information"
    cmdIdentifyNum = 10
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_PARTITION_VER()
    '6E 51 86 03 FE 77 13 00 00 20
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H13
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H20
    
    Log_Info "Read partition version(DDR)"
    cmdIdentifyNum = 11
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_RESOLUTION_INFO()
    '6E 51 86 03 FE 77 20 00 00 13
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H3
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H20
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H13
    
    Log_Info "Support 4K or 2K"
    cmdIdentifyNum = 12
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_AREA_INFO()
    '6E 51 86 01 FE 77 33 00 00 02
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H33
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H2
    
    Log_Info "Read area information"
    cmdIdentifyNum = 13
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_HDCP_KEY()
    '6E 51 86 01 FE 77 05 00 00 34
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H5
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H34
    
    Log_Info "Read HDCP key"
    cmdIdentifyNum = 14
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_MAC_ADDRESS()
    '6E 51 86 01 FE F0 01 01 00 B6
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &HF0
    mSendDataBuf(6) = &H1
    mSendDataBuf(7) = &H1
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &HB6
    
    Log_Info "Read MAC address"
    cmdIdentifyNum = 15
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_DEVICE_KEY()
    '6E 51 86 01 FE 77 34 00 00 05
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H34
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H5
    
    Log_Info "Read device key"
    cmdIdentifyNum = 16
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_WIDEVINE_KEY()
    '6E 51 86 01 FE 77 38 00 00 09
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H38
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H9
    
    Log_Info "Read widevine key"
    cmdIdentifyNum = 17
    gblCmdDataRecv = False
    
    SendCmd
End Sub

Public Sub READ_PLAYREADY_KEY()
    '6E 51 86 01 FE 77 39 00 00 08
    mSendDataBuf(0) = &H6E
    mSendDataBuf(1) = &H51
    mSendDataBuf(2) = &H86
    mSendDataBuf(3) = &H1
    mSendDataBuf(4) = &HFE
    mSendDataBuf(5) = &H77
    mSendDataBuf(6) = &H39
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H8
    
    Log_Info "Read playready key"
    cmdIdentifyNum = 18
    gblCmdDataRecv = False
    
    SendCmd
End Sub
