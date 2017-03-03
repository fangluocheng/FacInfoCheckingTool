Attribute VB_Name = "Module3"
Option Explicit

Public strBuff As String

'------------------------------------------------------------------------------
'    Variable mapping the check items in "CheckItem".
'------------------------------------------------------------------------------
Public chkTitleFlag(16) As Boolean
Public isUartMode As Boolean
Public isNetworkConnected As Boolean

'------------------------------------------------------------------------------
'    Variable mapping the items in "CheckItem".
'------------------------------------------------------------------------------
Public SetTVCurrentComBaud As Long                         'ComBaud
Public IsStepTime As Long                                  'Delayms
Public barcodeLen As Integer                               'SN_Len
Public strTvInfoSpec(11) As String

Public IsStop As Boolean
Public IsACK As Boolean

Public strCurrentModelName As String
Public strDataBindingExcelFileName As String
Public isOpenDataBindingExcelFile As Boolean

Public SetTVCurrentComID As Integer

Public IsSNWriteSuccess As Boolean
Public isCmdDataRecv As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long

Public cmdIdentifyNum As Integer

Public isConnect1730 As Boolean
Public delayMs01 As Long
Public delayMs02 As Long
Public port1730 As Long
Public strErpUrl As String
Public strErpOrganization As String

Public Const strChkBoxUnselected As String = "----"
Public Const strNoRecvData As String = "None"
Public Const cmdResendTimes As Integer = 2
Public Const cmdReceiveWaitS As Integer = 5
Public Const itemNumOfTvInfo As Integer = 16
Public Const strRemoteHost As String = "192.168.1.11"
Public Const lngRemotePort As Long = 8888

Public Sub Log_Info(strLog As String)
    Form1.TxtReceive.Text = Form1.TxtReceive.Text + strLog + vbCrLf
    Form1.TxtReceive.SelStart = Len(Form1.TxtReceive)
    
    SaveLogInFile strLog
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
    MsgBox err.Source & "------" & err.Description
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
    MsgBox err.Source & "------" & err.Description
    Exit Sub
End Sub

Public Sub SaveLogInFile(strLog As String)
    Dim logPath As String

    logPath = App.Path & "\" & "Logs\"
    If Right(logPath, 1) <> "\" Then logPath = logPath & "\"
    
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    
    Open (logPath & Format(Date, "YYYY-MM-DD") & ".log") For Append As #1
    Write #1, CStr(Time) & "> " & strLog
    Close #1
End Sub
