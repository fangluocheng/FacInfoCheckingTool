Attribute VB_Name = "Module3"
Option Explicit

Public IsSNWriteSuccess As Boolean
Public scanbarcode As String
Public strSerialNo As String
Public countTime As Long

Public cmdIdentifyNum As Integer

Public Const cmdResendTimes As Integer = 2
Public Const cmdReceiveWaitS As Integer = 5

Public Sub Log_Info(strLog As String)
    FormMain.TxtReceive.Text = FormMain.TxtReceive.Text + strLog + vbCrLf
    FormMain.TxtReceive.SelStart = Len(FormMain.TxtReceive)
    
    SaveLogInFile strLog
End Sub

Public Sub Log_Clear()
    FormMain.TxtReceive.Text = ""
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
    Wend
    Exit Sub

ShowError:
    MsgBox Err.Source & "------" & Err.Description
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
