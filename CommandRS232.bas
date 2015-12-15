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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Enter Factory Mode" & vbCrLf
    cmdIdentifyNum = 0
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Exit Factory Mode" & vbCrLf
    cmdIdentifyNum = 1
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read system version" & vbCrLf
    cmdIdentifyNum = 2
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read Flash information" & vbCrLf
    cmdIdentifyNum = 3
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read hardware version" & vbCrLf
    cmdIdentifyNum = 4
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Support 3D or not" & vbCrLf
    cmdIdentifyNum = 5
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read 2.4G Version" & vbCrLf
    cmdIdentifyNum = 6
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read panel name" & vbCrLf
    cmdIdentifyNum = 7
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read carrier information" & vbCrLf
    cmdIdentifyNum = 8
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read HDCP key" & vbCrLf
    cmdIdentifyNum = 9
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub READ_MODEL_NAME()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E 51 86 01 FE F0 18 04 00 AA
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HF0
    SendDataBuf(6) = &H18
    SendDataBuf(7) = &H4
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &HAA
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read model name" & vbCrLf
    cmdIdentifyNum = 10
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Support 4K or 2K" & vbCrLf
    cmdIdentifyNum = 11
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read MAC address" & vbCrLf
    cmdIdentifyNum = 12
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read channel information" & vbCrLf
    cmdIdentifyNum = 13
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read partition version(DDR)" & vbCrLf
    cmdIdentifyNum = 14
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read area information" & vbCrLf
    cmdIdentifyNum = 15
    Form1.MSComm1.Output = SendDataBuf
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
    
    Form1.TxtReceive.Text = Form1.TxtReceive.Text & "Read device key" & vbCrLf
    cmdIdentifyNum = 16
    Form1.MSComm1.Output = SendDataBuf
End Sub

