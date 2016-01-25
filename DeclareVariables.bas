Attribute VB_Name = "Module3"
Option Explicit

Public strBuff As String

Public i As Integer

'------------------------------------------------------------------------------
'    Variable mapping the check items in "CheckItem".
'------------------------------------------------------------------------------
Public IsModelSelected As Boolean
Public IsSysVerSelected As Boolean
Public IsFlashInfoSelected As Boolean
Public IsHardwareVerSelected As Boolean
Public IsDimensionSelected As Boolean
Public IsChannelSelected As Boolean
Public IsPartitionVerSelected As Boolean
Public Is24GVerSelected As Boolean
Public IsPanelSelected As Boolean
Public IsCarrierSelected As Boolean
Public IsAreaVerSelected As Boolean
Public IsHDCPSelected As Boolean
Public IsResolutionSelected As Boolean
Public IsMACAddrSelected As Boolean
Public IsDeviceKeySelected As Boolean
Public isUartMode As Boolean
Public isNetworkConnected As Boolean

'------------------------------------------------------------------------------
'    Variable mapping the items in "CheckItem".
'------------------------------------------------------------------------------
Public SetTVCurrentComBaud As Long                         'ComBaud
Public IsStepTime As Long                                  'Delayms
Public barcodeLen As Integer                               'SN_Len
Public ModelSpec As String                                 'ModelM
Public SysVerSpec As String                                'SysVerM
Public FlashInfoSpec As String                             'FlashInfoM
Public HardwareVerSpec As String                           'HardwareVerM
Public DimensionSpec As String                             'DimensionM
Public ChannelSpec As String                               'ChannelM
Public PartitionVerSpec As String                          'PartitionVerM
Public TwoPointFourGVerSpec As String                      '24GVerM
Public PanelSpec As String                                 'PanelM
Public CarrierSpec As String                               'CarrierM
Public AreaSpec As String                                  'AreaM
Public ResolutionSpec As String                            'ResolutionM


Public IsStop As Boolean
Public IsACK As Boolean

Public strCurrentModelName As String

Public SetTVCurrentComID As Integer
Public SetData As Integer
Public SetDay As Integer

Public IsSNWriteSuccess As Boolean
Public isCmdDataRecv As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long

Public cmdIdentifyNum As Integer

Public Const strChkBoxUnselected As String = "----"
Public Const strNoRecvData As String = "None"
Public Const cmdResendTimes As Integer = 2
Public Const cmdReceiveWaitS As Integer = 5
Public Const strRemoteHost As String = "192.168.1.11"
Public Const lngRemotePort As Long = 8888

Public Sub Log_Info(strLog As String)
    Form1.TxtReceive.Text = Form1.TxtReceive.Text + strLog + vbCrLf
    Form1.TxtReceive.SelStart = Len(Form1.TxtReceive)
End Sub

Public Sub Log_Clear()
    Form1.TxtReceive.Text = ""
End Sub

Public Sub DelaySWithCmdFlag(Sec As Long, flag As Boolean)
On Error GoTo ShowError
    Dim start As Single
    start = Timer
    While (Timer - start) < Sec
        DoEvents
   
        If flag = True Then
            Exit Sub
        End If
        
        If IsStop = True Then
            Exit Sub
        End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub


Public Sub DelayMS(mmSec As Long)
On Error GoTo ShowError
    Dim start As Single
    start = Timer
    While (Timer - start) < (mmSec / 1000#)
        DoEvents
   
        If IsStop = True Then
            Exit Sub
        End If
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
    Exit Sub
End Sub
