VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ϢУ�鹤��"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12630
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   1125
      Left            =   4320
      TabIndex        =   32
      Top             =   6720
      Width           =   8175
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   33
         Text            =   "123456789"
         Top             =   240
         Width           =   7905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "���Խ��"
      Height          =   1125
      Left            =   120
      TabIndex        =   31
      Top             =   6720
      Width           =   4095
      Begin VB.Label lbResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Checking"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TV ��Ϣ"
      Height          =   5700
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   12375
      Begin VB.TextBox txtDeviceKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtMacAddr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtResolution 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtHdcpKey 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "None"
         Top             =   3945
         Width           =   4000
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtCarrier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPanelName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtTwoPointFourVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "None"
         Top             =   2865
         Width           =   4000
      End
      Begin VB.TextBox txtPartitionVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "None"
         Top             =   5040
         Width           =   4000
      End
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtDimension 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtHWVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "None"
         Top             =   1785
         Width           =   4000
      End
      Begin VB.TextBox txtFlashInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtSysVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4150
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.TextBox txtModelInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "None"
         Top             =   700
         Width           =   4000
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Device Key"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   29
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAC ��ַ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   27
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HDCP"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4K/2K"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   23
         Top             =   3480
         Width           =   4000
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   21
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����ƽ̨"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   19
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2.4G �汾"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���ͺ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   15
         Top             =   2400
         Width           =   4000
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�����汾"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   4000
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   11
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ӳ���汾"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D/3D"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   7
         Top             =   1320
         Width           =   4000
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flash ��Ϣ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   8190
         TabIndex        =   5
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ϵͳ�汾"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4150
         TabIndex        =   3
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4000
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11500
      Top             =   7000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu vbSet 
      Caption         =   "����"
      Begin VB.Menu tbSetComPort 
         Caption         =   "���ô���"
      End
      Begin VB.Menu vbSetSPEC 
         Caption         =   "�������ݹ��"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RES As Long
Dim Result As Boolean
Dim StepTime As Long

Private Sub Form_Load()

    i = 0
    SetTVCurrentComBaud = 115200
    StepTime = IsStepTime
    IsStop = False
    subInitComPort
    subInitInterface

    Label1 = strCurrentModelName

End Sub

Private Sub subInitInterface()
    txtInput.Text = ""
End Sub

Private Sub subInitComPort()
    sqlstring = "select * from CommonTable where Mark='ATS'"
    Executesql (sqlstring)

    If rs.EOF = False Then
        SetTVCurrentComID = rs("ComID")
    Else
        MsgBox "Read Data Error,Please Check Your Database!", vbOKOnly + vbInformation, "Warning!"
    End
    End If

    Set cn = Nothing
    Set rs = Nothing
    sqlstring = ""

    ComInit

End Sub

Private Function funSNWrite() As Boolean
strSerialNo = ""
scanbarcode = ""
strSerialNo = UCase$(txtInput.Text)
If subJudgeTheSNIsAvailable = True Then
  funSNWrite = True
  scanbarcode = strSerialNo
Else
  funSNWrite = False
End If
End Function

Private Sub subInitBeforeRunning()
    countTime = Timer
    IsSNWriteSuccess = True
    strSerialNo = ""
End Sub

Private Sub subInitAfterRunning()
countTime = CLng(Timer - countTime)

Label9.Caption = countTime & "S"
IsSNWriteSuccess = False

txtInput.Text = ""
txtInput.SetFocus

End Sub

Private Sub subMainProcesser()
    
    Dim i, j As Integer

On Error GoTo ErrExit
    subInitBeforeRunning
    If IsStop = True Then
        Exit Sub
    End If

    If IsSNWriteSuccess = funSNWrite Then
        If IsStop = True Then
            Exit Sub
        End If
        txtInput = ""
        Command2.SetFocus
    Else
        ShowError_Sys (6)
        GoTo FAIL
    End If

On Error GoTo ErrExit

If IsCa210ok = False Then
MsgBox "CA210 disconnected,Please click'Connect'->'Connect CA210'to do operation!", vbOKOnly + vbInformation, "warning"
txtInput.Text = ""
txtInput.SetFocus
Exit Sub
End If

    'Whether the CheckBox of database file(*.mdb) selected or not.
    'If not, config the TextBox
    If Not IsModel Then Check1.Value = 1 Else Check1.Value = 0
    End If
    If IsSysVer Then Check2.Value = 1 Else Check2.Value = 0
    End If
    If IsFlashInfo Then Check3.Value = 1 Else Check3.Value = 0
    End If
    If IsHardwareVer Then Check4.Value = 1 Else Check4.Value = 0
    End If
    If IsDimension Then Check5.Value = 1 Else Check5.Value = 0
    End If
    If IsChannel Then Check6.Value = 1 Else Check6.Value = 0
    End If
    If Is24GVer Then Check7.Value = 1 Else Check7.Value = 0
    End If
    If IsPanel Then Check8.Value = 1 Else Check8.Value = 0
    End If
    If IsCarrier Then Check9.Value = 1 Else Check9.Value = 0
    End If
    If IsHDCP Then Check10.Value = 1 Else Check10.Value = 0
    End If
    If IsResolution Then Check11.Value = 1 Else Check11.Value = 0
    End If
    If IsMACAddr Then Check12.Value = 1 Else Check12.Value = 0
    End If
    If IsPartitionVer Then Check13.Value = 1 Else Check13.Value = 0
    End If
    If IsArea Then Check14.Value = 1 Else Check14.Value = 0
    End If
    If IsDeviceKey Then Check15.Value = 1 Else Check15.Value = 0
    End If

End Sub

Private Sub tbSetComPort_Click()
    Form2.Show
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Command1_Click()
    IsStop = False
    subMainProcesser
     
    If IsStop = True Then
        Exit Sub
    End If
End Sub

Private Sub txtResult_Change()

End Sub

Private Sub vbSetSPEC_Click()
    frmSetData.Show
End Sub
