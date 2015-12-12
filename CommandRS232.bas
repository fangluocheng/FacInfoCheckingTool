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

Public Sub SET_BURNING_MODE(flag As Long)

    Dim SendDataBuf(0 To 8) As Byte
    
    '6E  51  86  03  FE  E1  A2  00  Flag(01/00)
    SendDataBuf(0) = &H6E
    SendDataBuf(1) = &H51
    SendDataBuf(2) = &H86
    SendDataBuf(3) = &H3
    SendDataBuf(4) = &HFE
    SendDataBuf(5) = &HE1
    SendDataBuf(6) = &HA2
    SendDataBuf(7) = &H0
    SendDataBuf(8) = CByte(flag)
    
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
