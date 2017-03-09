Attribute VB_Name = "Module1"
Option Explicit

Public Sub ClearComBuf()
    If gblUartMode Then
        FormMain.MSComm1.InBufferCount = 0
        FormMain.MSComm1.OutBufferCount = 0
    End If
End Sub
