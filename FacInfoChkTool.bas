Attribute VB_Name = "FacInfoChkTool"
Option Explicit


Public Type TypeProperty
    Items(11) As String
    ItemChk(17) As Boolean
End Type

Public Const XML_COMMODE_UART As String = "UART"
Public Const XML_COMMODE_NET As String = "NET"
Public Const XML_TRUE As String = "TRUE"
Public Const XML_FALSE As String = "FALSE"

Public Const TVINFO_ITEMUNCHECK As String = "----"
Public Const TVINFO_INIT As String = "None"
Public Const REMOTE_HOST As String = "192.168.1.11"
Public Const REMOTE_PORT As Long = 8888
Public Const ITEMS_NUM As Integer = 17
Public Const cmdResendTimes As Integer = 2
Public Const cmdReceiveWaitS As Integer = 5

Public gstrXmlPath As String
Public gblUartMode As Boolean
Public gblNetConnected As Boolean
Public glngDelayTime As Long
Public gintSNLen As Integer
Public gintMACLen As Integer
Public gutdPropertySetting As TypeProperty
Public gblExitFacCmd As Boolean
Public gblSaveData As Boolean
Public gblCmdDataRecv As Boolean
Public gintCmdId As Integer

Private mstrComBaud As String
Private mintComID As Integer

Public Sub Main()
    FormSplash.Show
End Sub

Public Sub LoadFormMain()
    Dim i As Integer

    ParseXml

    gblCmdDataRecv = False
    
    With FormMain
        .Show
        .Enabled = False
        .Label1.Caption = gutdPropertySetting.Items(0)
        .TextTvSN.Text = ""
        .TextMacSN.Text = ""
        .TextTvSN.Enabled = True
        .TextMacSN.Enabled = False

        If gblUartMode Then
            InitComPort
        Else
            InitNetwork
        End If

        For i = 0 To ITEMS_NUM
            If Not gutdPropertySetting.ItemChk(i) Then
                .lbTVInfo(i).Caption = TVINFO_ITEMUNCHECK
                .lbTVInfo(i).BackColor = &HE0E0E0
            End If
        Next i
        .Enabled = True
    End With
End Sub

Private Sub ParseXml()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(gstrXmlPath)

    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        If UCase(xmlDoc.selectSingleNode("/Settings/Communication").selectSingleNode("@mode").Text) = XML_COMMODE_UART Then
            gblUartMode = True
            mstrComBaud = xmlDoc.selectSingleNode("/Settings/Communication/Uart").selectSingleNode("@baud").Text
            mintComID = Val(xmlDoc.selectSingleNode("/Settings/Communication/Uart").selectSingleNode("@id").Text)
        Else
            gblUartMode = False
        End If
        glngDelayTime = Val(xmlDoc.selectSingleNode("/Settings/Delayms").Text)
        gintSNLen = Val(xmlDoc.selectSingleNode("/Settings/SNLen").Text)
        gintMACLen = Val(xmlDoc.selectSingleNode("/Settings/MACLen").Text)
        'Send Exit Factory Mode Command or not.
        If UCase(xmlDoc.selectSingleNode("/Settings/ExitFacCmd").selectSingleNode("@enable").Text) = XML_TRUE Then
            gblExitFacCmd = True
        Else
            gblExitFacCmd = False
        End If
        'Save data or not.
        If UCase(xmlDoc.selectSingleNode("/Settings/SaveData").selectSingleNode("@enable").Text) = XML_TRUE Then
            gblSaveData = True
        Else
            gblSaveData = False
        End If
        'Model
        gutdPropertySetting.Items(0) = xmlDoc.selectSingleNode("/Settings/Model").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Model").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(0) = True
        Else
            gutdPropertySetting.ItemChk(0) = False
        End If
        'System Version
        gutdPropertySetting.Items(1) = xmlDoc.selectSingleNode("/Settings/SysVer").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/SysVer").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(1) = True
        Else
            gutdPropertySetting.ItemChk(1) = False
        End If
        'Flash Information
        gutdPropertySetting.Items(2) = xmlDoc.selectSingleNode("/Settings/FlashInfo").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/FlashInfo").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(2) = True
        Else
            gutdPropertySetting.ItemChk(2) = False
        End If
        'Hardware Version
        gutdPropertySetting.Items(3) = xmlDoc.selectSingleNode("/Settings/HWVer").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/HWVer").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(3) = True
        Else
            gutdPropertySetting.ItemChk(3) = False
        End If
        'Dimension
        gutdPropertySetting.Items(4) = xmlDoc.selectSingleNode("/Settings/Dimension").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Dimension").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(4) = True
        Else
            gutdPropertySetting.ItemChk(4) = False
        End If
        'Channel
        gutdPropertySetting.Items(5) = xmlDoc.selectSingleNode("/Settings/Channel").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Channel").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(5) = True
        Else
            gutdPropertySetting.ItemChk(5) = False
        End If
        '2.4G Version
        gutdPropertySetting.Items(6) = xmlDoc.selectSingleNode("/Settings/RemoteVer").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/RemoteVer").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(6) = True
        Else
            gutdPropertySetting.ItemChk(6) = False
        End If
        'Panel
        gutdPropertySetting.Items(7) = xmlDoc.selectSingleNode("/Settings/Panel").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Panel").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(7) = True
        Else
            gutdPropertySetting.ItemChk(7) = False
        End If
        'Carrier
        gutdPropertySetting.Items(8) = xmlDoc.selectSingleNode("/Settings/Carrier").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Carrier").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(8) = True
        Else
            gutdPropertySetting.ItemChk(8) = False
        End If
        'Partition Version
        gutdPropertySetting.Items(9) = xmlDoc.selectSingleNode("/Settings/PartitionVer").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/PartitionVer").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(9) = True
        Else
            gutdPropertySetting.ItemChk(9) = False
        End If
        'Resolution
        gutdPropertySetting.Items(10) = xmlDoc.selectSingleNode("/Settings/Resolution").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Resolution").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(10) = True
        Else
            gutdPropertySetting.ItemChk(10) = False
        End If
        'Area
        gutdPropertySetting.Items(11) = xmlDoc.selectSingleNode("/Settings/Area").Text
        If UCase(xmlDoc.selectSingleNode("/Settings/Area").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(11) = True
        Else
            gutdPropertySetting.ItemChk(11) = False
        End If
        'HDCP
        If UCase(xmlDoc.selectSingleNode("/Settings/HDCP").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(12) = True
        Else
            gutdPropertySetting.ItemChk(12) = False
        End If
        'MAC
        If UCase(xmlDoc.selectSingleNode("/Settings/MAC").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(13) = True
        Else
            gutdPropertySetting.ItemChk(13) = False
        End If
        'Device Key
        If UCase(xmlDoc.selectSingleNode("/Settings/DeviceKey").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(14) = True
        Else
            gutdPropertySetting.ItemChk(14) = False
        End If
        'Widevien Key
        If UCase(xmlDoc.selectSingleNode("/Settings/WidevienKey").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(15) = True
        Else
            gutdPropertySetting.ItemChk(15) = False
        End If
        'Playready Key
        If UCase(xmlDoc.selectSingleNode("/Settings/PlayreadyKey").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(16) = True
        Else
            gutdPropertySetting.ItemChk(16) = False
        End If
        'SN_NUM
        If UCase(xmlDoc.selectSingleNode("/Settings/SNNum").selectSingleNode("@enable").Text) = XML_TRUE Then
            gutdPropertySetting.ItemChk(17) = True
        Else
            gutdPropertySetting.ItemChk(17) = False
        End If
    End If
End Sub

Private Sub InitComPort()
    On Error GoTo ErrExit

    With FormMain.MSComm1
        If .PortOpen = True Then
            .PortOpen = False
        End If
        .CommPort = mintComID
        .Settings = mstrComBaud & ",N,8,1"
        .InputLen = 0
        .InBufferCount = 0
        .OutBufferCount = 0
        .InputMode = comInputModeBinary
        .NullDiscard = False
        .DTREnable = False
        .EOFEnable = False
        .RTSEnable = False
        .SThreshold = 1
        .RThreshold = 1
        .InBufferSize = 1024
        .OutBufferSize = 512
    End With
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub InitNetwork()
    gblNetConnected = False
    With FormMain.tcpClient
        .Protocol = sckTCPProtocol
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = REMOTE_HOST
        .RemotePort = REMOTE_PORT
    End With
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
End Sub

Public Sub Log_Info(strLog As String)
    FormMain.TxtReceive.Text = FormMain.TxtReceive.Text & strLog & vbCrLf
    FormMain.TxtReceive.SelStart = Len(FormMain.TxtReceive)
    
    SaveLogInFile strLog
End Sub

Public Sub Log_Clear()
    FormMain.TxtReceive.Text = ""
End Sub

