Attribute VB_Name = "Module1"
Option Explicit


Public Sub ComInit()

On Error GoTo ErrExit
 
    If Form1.MSComm1.PortOpen = True Then
        Form1.MSComm1.PortOpen = False
    End If

    With Form1
        .MSComm1.CommPort = SetTVCurrentComID
        .MSComm1.Settings = SetTVCurrentComBaud & ",N,8,1"
        .MSComm1.InputLen = 0
        
        .MSComm1.InBufferCount = 0
        .MSComm1.OutBufferCount = 0
        .MSComm1.InputMode = comInputModeBinary            'comInputModeHex
        
        .MSComm1.NullDiscard = False
        .MSComm1.DTREnable = False
        .MSComm1.EOFEnable = False
        .MSComm1.RTSEnable = False
        .MSComm1.SThreshold = 1
        .MSComm1.RThreshold = 1
        .MSComm1.InBufferSize = 1024
        .MSComm1.OutBufferSize = 512
        
        .MSComm1.PortOpen = True
 
    End With
    Exit Sub

ErrExit:
        MsgBox Err.Description, vbCritical, Err.Source
End Sub

Public Sub ClearComBuf()
    If isUartMode Then
        Form1.MSComm1.InBufferCount = 0
        Form1.MSComm1.OutBufferCount = 0
    End If
End Sub
