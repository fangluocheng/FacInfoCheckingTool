Attribute VB_Name = "Module2"
Option Explicit

Public Sub ENTER_FAC_MODE()

    Dim SendDataBuf(0 To 8) As Byte
    
    '6E  51  86  03  FE  E1  A0  00  01
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE1
    SendDataBuf(6) = &HA0
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H1
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub EXIT_FAC_MODE()

    Dim SendDataBuf(0 To 8) As Byte
    
    '6E  51  86  03  FE  E1  A0  00  00
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE1
    SendDataBuf(6) = &HA0
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_SYS_VERSION()

    Dim SendDataBuf(0 To 8) As Byte
    
    '6E  51  86  01  FE  E4  13  00  00
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE4
    SendDataBuf(6) = &H13
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_FLASH_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  0F  00  00  3C
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_HARDWARE_VERSION()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  16  00  00  25
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_DIMENSION_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  19  00  00  2A
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_24G_VERSION()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  14  00  00  27
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_PANEL_NAME()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  17  00  00  24
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_BROADCAST_CTRL_PLATFORM()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  18  00  00  2B
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub READ_HDCP_KEY()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  01  FE  77  05  00  00  36
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &H77
    SendDataBuf(6) = &H5
    SendDataBuf(7) = &H0
    SendDataBuf(8) = &H0
    SendDataBuf(9) = &H36
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_MODEL_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  15  00  00  26
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_RESOLUTION_INFO()

    Dim SendDataBuf(0 To 9) As Byte
    
    '6E  51  86  03  FE  77  20  00  00  13
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
    
    Form1.MSComm1.Output = SendDataBuf
End Sub

Public Sub GET_MAC_ADDRESS()

    Dim SendDataBuf(0 To 8) As Byte
    
    '6E  51  86  01  FE  F0  01  01  00
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H1
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HF0
    SendDataBuf(6) = &H1
    SendDataBuf(7) = &H1
    SendDataBuf(8) = &H0
    
    Form1.MSComm1.Output = SendDataBuf
End Sub
